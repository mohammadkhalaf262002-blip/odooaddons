# -*- coding: utf-8 -*-

from odoo import models, fields, api
from odoo.exceptions import UserError
import requests
import json
import base64
import logging
import io

_logger = logging.getLogger(__name__)

# Try to import openpyxl for Excel parsing
try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False
    _logger.warning('openpyxl not installed. Excel import will not work.')


class TotalinePriceSearch(models.Model):
    _name = 'totaline.price.search'
    _inherit = ['mail.thread', 'mail.activity.mixin']
    _description = 'Totaline Price Search'
    _order = 'create_date desc'

    name = fields.Char(string='Search Name', required=True, default=lambda self: f"Search {fields.Datetime.now()}")
    state = fields.Selection([
        ('draft', 'Draft'),
        ('searching', 'Searching'),
        ('done', 'Completed'),
        ('error', 'Error'),
    ], string='Status', default='draft')

    # Excel Upload
    excel_file = fields.Binary(string='Excel File', attachment=True)
    excel_filename = fields.Char(string='Filename')

    @api.onchange('excel_file')
    def _onchange_excel_file(self):
        """Parse Excel file when uploaded and populate product lines"""
        if not self.excel_file:
            return

        if not OPENPYXL_AVAILABLE:
            raise UserError('Excel parsing library (openpyxl) is not installed.')

        try:
            # Decode base64 file
            file_content = base64.b64decode(self.excel_file)
            workbook = openpyxl.load_workbook(io.BytesIO(file_content))
            sheet = workbook.active

            # Clear existing lines
            self.search_line_ids = [(5, 0, 0)]

            # Find header row and map columns
            headers = {}
            header_row = 1
            for col_idx, cell in enumerate(sheet[1], 1):
                if cell.value:
                    header_name = str(cell.value).strip().lower()
                    headers[header_name] = col_idx

            # Map Turkish/English column names
            column_mapping = {
                'product_name': ['ürün adı', 'urun adi', 'product name', 'product', 'ürün', 'urun'],
                'brand': ['marka', 'brand'],
                'quantity': ['miktar', 'quantity', 'qty', 'adet'],
                'unit': ['birim', 'unit'],
                'description': ['açıklama', 'aciklama', 'description', 'desc', 'detay'],
            }

            # Find column indices
            col_indices = {}
            for field, names in column_mapping.items():
                for name in names:
                    if name in headers:
                        col_indices[field] = headers[name]
                        break

            if 'product_name' not in col_indices:
                raise UserError('Could not find product name column. Expected: "Ürün Adı" or "Product Name"')

            # Parse data rows
            lines = []
            for row_idx in range(2, sheet.max_row + 1):
                product_name = sheet.cell(row=row_idx, column=col_indices.get('product_name', 1)).value

                if not product_name:
                    continue  # Skip empty rows

                line_data = {
                    'product_name': str(product_name).strip(),
                    'brand': '',
                    'quantity': 1,
                    'unit': '',
                    'description': '',
                }

                if 'brand' in col_indices:
                    val = sheet.cell(row=row_idx, column=col_indices['brand']).value
                    line_data['brand'] = str(val).strip() if val else ''

                if 'quantity' in col_indices:
                    val = sheet.cell(row=row_idx, column=col_indices['quantity']).value
                    try:
                        line_data['quantity'] = float(val) if val else 1
                    except (ValueError, TypeError):
                        line_data['quantity'] = 1

                if 'unit' in col_indices:
                    val = sheet.cell(row=row_idx, column=col_indices['unit']).value
                    line_data['unit'] = str(val).strip() if val else ''

                if 'description' in col_indices:
                    val = sheet.cell(row=row_idx, column=col_indices['description']).value
                    line_data['description'] = str(val).strip() if val else ''

                lines.append((0, 0, line_data))

            self.search_line_ids = lines

            return {
                'type': 'ir.actions.client',
                'tag': 'display_notification',
                'params': {
                    'title': 'Excel Imported',
                    'message': f'{len(lines)} products loaded successfully.',
                    'type': 'success',
                    'sticky': False,
                }
            }

        except UserError:
            raise
        except Exception as e:
            _logger.exception('Failed to parse Excel file')
            raise UserError(f'Failed to parse Excel file: {str(e)}')

    # n8n Webhook URL
    webhook_url = fields.Char(
        string='Webhook URL',
        default='https://your-n8n-instance.com/webhook/totaline',
        help='n8n webhook URL for price search'
    )

    # Results
    search_line_ids = fields.One2many('totaline.price.search.line', 'search_id', string='Search Results')

    # Totals
    total_best_price = fields.Float(string='Total (Best Price)', compute='_compute_totals', store=True)
    total_second_price = fields.Float(string='Total (2nd Price)', compute='_compute_totals', store=True)
    total_third_price = fields.Float(string='Total (3rd Price)', compute='_compute_totals', store=True)

    # Info
    product_count = fields.Integer(string='Product Count', compute='_compute_product_count')
    search_date = fields.Datetime(string='Search Date', readonly=True)
    error_message = fields.Text(string='Error Message')

    # Scheduling fields
    is_scheduled = fields.Boolean(string='Scheduled Search', default=False,
                                   help='Enable automatic scheduled price searches')
    schedule_frequency = fields.Selection([
        ('daily', 'Daily'),
        ('weekly', 'Weekly'),
        ('monthly', 'Monthly'),
    ], string='Frequency', default='weekly')
    schedule_time = fields.Float(string='Search Time', default=8.0,
                                  help='Time of day to run search (24h format, e.g. 8.0 = 08:00)')
    schedule_day = fields.Selection([
        ('0', 'Monday'),
        ('1', 'Tuesday'),
        ('2', 'Wednesday'),
        ('3', 'Thursday'),
        ('4', 'Friday'),
        ('5', 'Saturday'),
        ('6', 'Sunday'),
    ], string='Day of Week', default='0', help='For weekly schedules')
    schedule_day_of_month = fields.Integer(string='Day of Month', default=1,
                                            help='For monthly schedules (1-28)')
    next_scheduled_date = fields.Datetime(string='Next Scheduled Run', compute='_compute_next_scheduled', store=True)
    last_scheduled_run = fields.Datetime(string='Last Scheduled Run', readonly=True)

    @api.depends('is_scheduled', 'schedule_frequency', 'schedule_time', 'schedule_day', 'schedule_day_of_month')
    def _compute_next_scheduled(self):
        from datetime import datetime, timedelta
        for record in self:
            if not record.is_scheduled:
                record.next_scheduled_date = False
                continue

            now = fields.Datetime.now()
            hour = int(record.schedule_time)
            minute = int((record.schedule_time % 1) * 60)

            if record.schedule_frequency == 'daily':
                next_run = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
                if next_run <= now:
                    next_run += timedelta(days=1)
            elif record.schedule_frequency == 'weekly':
                target_day = int(record.schedule_day or '0')
                days_ahead = target_day - now.weekday()
                if days_ahead < 0:
                    days_ahead += 7
                next_run = now + timedelta(days=days_ahead)
                next_run = next_run.replace(hour=hour, minute=minute, second=0, microsecond=0)
                if next_run <= now:
                    next_run += timedelta(days=7)
            else:  # monthly
                day = min(record.schedule_day_of_month or 1, 28)
                next_run = now.replace(day=day, hour=hour, minute=minute, second=0, microsecond=0)
                if next_run <= now:
                    if now.month == 12:
                        next_run = next_run.replace(year=now.year + 1, month=1)
                    else:
                        next_run = next_run.replace(month=now.month + 1)

            record.next_scheduled_date = next_run

    @api.depends('search_line_ids')
    def _compute_product_count(self):
        for record in self:
            record.product_count = len(record.search_line_ids)

    @api.depends('search_line_ids.price_1', 'search_line_ids.price_2', 'search_line_ids.price_3')
    def _compute_totals(self):
        for record in self:
            record.total_best_price = sum(line.price_1 or 0 for line in record.search_line_ids)
            record.total_second_price = sum(line.price_2 or 0 for line in record.search_line_ids)
            record.total_third_price = sum(line.price_3 or 0 for line in record.search_line_ids)

    def action_search_prices(self):
        """Send products to n8n webhook and get prices"""
        self.ensure_one()

        if not self.search_line_ids:
            self.state = 'error'
            self.error_message = 'No products to search. Please add products first.'
            return

        self.state = 'searching'
        self.search_date = fields.Datetime.now()
        self.error_message = False

        try:
            # Prepare products from search lines
            products = []
            for line in self.search_line_ids:
                products.append({
                    'id': str(line.id),
                    'name': line.product_name,
                    'brand': line.brand or '',
                    'quantity': line.quantity or 1,
                    'unit': line.unit or '',
                    'description': line.description or '',
                })

            if not products:
                self.state = 'error'
                self.error_message = 'No products to search. Please import products first.'
                return

            # Send to n8n webhook
            response = requests.post(
                self.webhook_url,
                json=products,
                headers={'Content-Type': 'application/json'},
                timeout=300  # 5 minutes timeout
            )

            if response.status_code == 200:
                results = response.json()
                _logger.info(f'Received response from n8n: {type(results)}')
                self._process_search_results(results)
                self._save_price_history()  # Save to history
                self.state = 'done'
            else:
                self.state = 'error'
                self.error_message = f'Webhook error: {response.status_code} - {response.text}'

        except requests.exceptions.Timeout:
            self.state = 'error'
            self.error_message = 'Search timed out. Please try again.'
        except Exception as e:
            self.state = 'error'
            self.error_message = str(e)
            _logger.exception('Price search failed')

    def _save_price_history(self):
        """Save current prices to history for tracking"""
        PriceHistory = self.env['totaline.price.history']
        for line in self.search_line_ids:
            if line.price_1 > 0:
                PriceHistory.create({
                    'search_id': self.id,
                    'search_line_id': line.id,
                    'product_name': line.product_name,
                    'brand': line.brand,
                    'quantity': line.quantity,
                    'unit': line.unit,
                    'price_1': line.price_1,
                    'store_1': line.store_1,
                    'url_1': line.url_1,
                    'price_2': line.price_2,
                    'store_2': line.store_2,
                    'price_3': line.price_3,
                    'store_3': line.store_3,
                    'search_date': self.search_date,
                })

    def _process_search_results(self, results):
        """Update search lines with results from n8n"""
        _logger.info(f'Processing search results: {type(results)}')

        # Handle nested response formats from n8n
        if isinstance(results, list) and len(results) > 0:
            first_item = results[0]
            if isinstance(first_item, dict) and 'results' in first_item:
                results = first_item.get('results', [])
                if isinstance(results, list) and len(results) > 0:
                    first_inner = results[0]
                    if isinstance(first_inner, dict) and 'results' in first_inner:
                        results = first_inner.get('results', [])

        if isinstance(results, dict):
            results = results.get('results', [results])

        if not isinstance(results, list):
            _logger.warning(f'Unexpected results format: {type(results)}')
            return

        # Build maps for matching
        line_map_by_name = {}
        line_map_by_id = {}
        for line in self.search_line_ids:
            key = line.product_name.lower().strip() if line.product_name else ''
            line_map_by_name[key] = line
            line_map_by_id[str(line.id)] = line

        for result in results:
            if not isinstance(result, dict):
                continue

            # Find matching line
            line = None
            line_id = result.get('id')
            if line_id:
                line = line_map_by_id.get(str(line_id))

            if not line:
                product_name = result.get('product', result.get('name', result.get('product_name', '')))
                product_key = product_name.lower().strip() if product_name else ''
                line = line_map_by_name.get(product_key)
                if not line:
                    continue

            # Get prices
            price_1 = result.get('price_1', 0)
            store_1 = result.get('store_1', '')
            url_1 = result.get('url_1', '')
            price_2 = result.get('price_2', 0)
            store_2 = result.get('store_2', '')
            url_2 = result.get('url_2', '')
            price_3 = result.get('price_3', 0)
            store_3 = result.get('store_3', '')
            url_3 = result.get('url_3', '')

            # Fallback to trusted_results array
            if not price_1:
                top_results = result.get('trusted_results', result.get('results', []))
                if top_results and len(top_results) > 0:
                    price_1 = top_results[0].get('totalPrice', 0)
                    store_1 = top_results[0].get('store', '')
                    url_1 = top_results[0].get('link', '')
                if top_results and len(top_results) > 1:
                    price_2 = top_results[1].get('totalPrice', 0)
                    store_2 = top_results[1].get('store', '')
                    url_2 = top_results[1].get('link', '')
                if top_results and len(top_results) > 2:
                    price_3 = top_results[2].get('totalPrice', 0)
                    store_3 = top_results[2].get('store', '')
                    url_3 = top_results[2].get('link', '')

            warning = result.get('warning', '')

            line.write({
                'price_1': price_1 or 0,
                'store_1': store_1 or '',
                'url_1': url_1 or '',
                'price_2': price_2 or 0,
                'store_2': store_2 or '',
                'url_2': url_2 or '',
                'price_3': price_3 or 0,
                'store_3': store_3 or '',
                'url_3': url_3 or '',
                'warning': warning or '',
            })

    def action_import_excel(self):
        """Open wizard to import Excel file"""
        return {
            'type': 'ir.actions.act_window',
            'name': 'Import Products',
            'res_model': 'totaline.import.wizard',
            'view_mode': 'form',
            'target': 'new',
            'context': {'default_search_id': self.id},
        }

    def action_reset_draft(self):
        """Reset to draft state"""
        self.state = 'draft'
        self.error_message = False

    def action_view_price_history(self):
        """View price history for this search"""
        self.ensure_one()
        return {
            'type': 'ir.actions.act_window',
            'name': 'Price History',
            'res_model': 'totaline.price.history',
            'view_mode': 'list,graph,pivot',
            'domain': [('search_id', '=', self.id)],
            'context': {'default_search_id': self.id},
        }

    def action_create_purchase_order(self):
        """Create purchase order from selected prices"""
        self.ensure_one()

        if self.state != 'done':
            raise UserError('Please complete the price search first.')

        # Check if there are any results with prices
        lines_with_prices = self.search_line_ids.filtered(lambda l: l.selected_price_value > 0)
        if not lines_with_prices:
            raise UserError('No products with prices found. Please run the price search first.')

        # Get or create the vendor
        Partner = self.env['res.partner']
        vendor = Partner.search([('name', '=', 'Online Price Search')], limit=1)
        if not vendor:
            vendor = Partner.create({
                'name': 'Online Price Search',
                'supplier_rank': 1,
                'company_type': 'company',
                'comment': 'Auto-created vendor for price search purchases',
            })

        # Get product model
        Product = self.env['product.product']

        # Prepare PO lines
        po_lines = []
        notes = []
        total_selected = 0

        for line in lines_with_prices:
            # Use linked Odoo product if available, otherwise find or create
            if line.product_id:
                product = line.product_id
            else:
                product = Product.search([('name', 'ilike', line.product_name)], limit=1)

                if not product:
                    product = Product.create({
                        'name': f"{line.product_name} ({line.brand})" if line.brand else line.product_name,
                        'type': 'consu',
                        'purchase_ok': True,
                        'sale_ok': False,
                    })

            # Get selected price, store, and URL
            selected_price = line.selected_price_value
            selected_store = line.selected_store_value
            selected_url = line.selected_url_value
            qty = line.quantity or 1

            # Build description
            description = f"{line.product_name}"
            if line.brand:
                description += f" - {line.brand}"
            if line.description:
                description += f" ({line.description})"
            description += f"\nStore: {selected_store}"
            if selected_url:
                description += f"\nURL: {selected_url}"
            if line.warning:
                description += f"\nNote: {line.warning}"

            po_lines.append((0, 0, {
                'product_id': product.id,
                'name': description,
                'product_qty': qty,
                'price_unit': selected_price / qty if qty else selected_price,
                'date_planned': fields.Datetime.now(),
            }))

            total_selected += selected_price
            notes.append(f"• {line.product_name}: ₺{selected_price:,.2f} from {selected_store} (Price {line.selected_price})")

        # Create the Purchase Order
        PurchaseOrder = self.env['purchase.order']
        po = PurchaseOrder.create({
            'partner_id': vendor.id,
            'date_order': fields.Datetime.now(),
            'notes': f"Created from Price Search: {self.name}\n\nProducts:\n" + "\n".join(notes),
            'order_line': po_lines,
        })

        self.message_post(
            body=f"Purchase Order <a href='#' data-oe-model='purchase.order' data-oe-id='{po.id}'>{po.name}</a> created with {len(po_lines)} products. Total: ₺{total_selected:,.2f}",
            message_type='notification',
        )

        return {
            'type': 'ir.actions.act_window',
            'name': 'Purchase Order',
            'res_model': 'purchase.order',
            'res_id': po.id,
            'view_mode': 'form',
            'target': 'current',
        }

    @api.model
    def _cron_run_scheduled_searches(self):
        """Cron job to run scheduled price searches"""
        now = fields.Datetime.now()
        scheduled_searches = self.search([
            ('is_scheduled', '=', True),
            ('next_scheduled_date', '<=', now),
        ])

        for search in scheduled_searches:
            try:
                _logger.info(f'Running scheduled search: {search.name}')
                search.action_search_prices()
                search.last_scheduled_run = now
                # Recompute next scheduled date
                search._compute_next_scheduled()
            except Exception as e:
                _logger.exception(f'Scheduled search failed: {search.name}')
                search.message_post(
                    body=f"Scheduled search failed: {str(e)}",
                    message_type='notification',
                )


class TotalinePriceSearchLine(models.Model):
    _name = 'totaline.price.search.line'
    _description = 'Totaline Price Search Line'
    _order = 'sequence, id'

    search_id = fields.Many2one('totaline.price.search', string='Search', ondelete='cascade')
    sequence = fields.Integer(string='Sequence', default=10)

    # Product Info (from Excel)
    product_name = fields.Char(string='Product Name', required=True)
    brand = fields.Char(string='Brand')
    quantity = fields.Float(string='Quantity', default=1)
    unit = fields.Char(string='Unit')
    description = fields.Char(string='Description')

    # Link to existing Odoo product (optional)
    product_id = fields.Many2one('product.product', string='Odoo Product',
                                  help='Link to existing product in Odoo. Leave empty to create new.')

    # Price Selection for PO
    selected_price = fields.Selection([
        ('1', '1. Fiyat (En İyi)'),
        ('2', '2. Fiyat'),
        ('3', '3. Fiyat'),
    ], string='Seçilen Fiyat', default='1', help='Select which price to use for Purchase Order')

    # Price Results
    price_1 = fields.Float(string='1st Price')
    store_1 = fields.Char(string='1st Store')
    url_1 = fields.Char(string='1st URL')
    store_link_1 = fields.Html(string='1. Store', compute='_compute_store_links', sanitize=False)

    price_2 = fields.Float(string='2nd Price')
    store_2 = fields.Char(string='2nd Store')
    url_2 = fields.Char(string='2nd URL')
    store_link_2 = fields.Html(string='2. Store', compute='_compute_store_links', sanitize=False)

    price_3 = fields.Float(string='3rd Price')
    store_3 = fields.Char(string='3rd Store')
    url_3 = fields.Char(string='3rd URL')
    store_link_3 = fields.Html(string='3. Store', compute='_compute_store_links', sanitize=False)

    # Calculated
    price_diff_percent = fields.Float(string='% Diff', compute='_compute_price_diff')
    warning = fields.Char(string='Warning')

    # Computed: Selected price and store for PO
    selected_price_value = fields.Float(string='Selected Price Value', compute='_compute_selected_values')
    selected_store_value = fields.Char(string='Selected Store Value', compute='_compute_selected_values')
    selected_url_value = fields.Char(string='Selected URL Value', compute='_compute_selected_values')

    # Price history link
    history_count = fields.Integer(string='History Count', compute='_compute_history_count')

    def _compute_history_count(self):
        PriceHistory = self.env['totaline.price.history']
        for line in self:
            line.history_count = PriceHistory.search_count([('search_line_id', '=', line.id)])

    def action_view_line_history(self):
        """View price history for this specific product"""
        self.ensure_one()
        return {
            'type': 'ir.actions.act_window',
            'name': f'Price History - {self.product_name}',
            'res_model': 'totaline.price.history',
            'view_mode': 'list,graph',
            'domain': [('product_name', '=', self.product_name), ('brand', '=', self.brand or False)],
        }

    @api.depends('selected_price', 'price_1', 'price_2', 'price_3',
                 'store_1', 'store_2', 'store_3', 'url_1', 'url_2', 'url_3')
    def _compute_selected_values(self):
        for line in self:
            if line.selected_price == '2':
                line.selected_price_value = line.price_2
                line.selected_store_value = line.store_2
                line.selected_url_value = line.url_2
            elif line.selected_price == '3':
                line.selected_price_value = line.price_3
                line.selected_store_value = line.store_3
                line.selected_url_value = line.url_3
            else:  # Default to price 1
                line.selected_price_value = line.price_1
                line.selected_store_value = line.store_1
                line.selected_url_value = line.url_1

    @api.depends('store_1', 'url_1', 'store_2', 'url_2', 'store_3', 'url_3')
    def _compute_store_links(self):
        for line in self:
            # Store 1
            if line.url_1 and line.store_1:
                line.store_link_1 = f'<a href="{line.url_1}" target="_blank">{line.store_1} ↗</a>'
            else:
                line.store_link_1 = line.store_1 or ''

            # Store 2
            if line.url_2 and line.store_2:
                line.store_link_2 = f'<a href="{line.url_2}" target="_blank">{line.store_2} ↗</a>'
            else:
                line.store_link_2 = line.store_2 or ''

            # Store 3
            if line.url_3 and line.store_3:
                line.store_link_3 = f'<a href="{line.url_3}" target="_blank">{line.store_3} ↗</a>'
            else:
                line.store_link_3 = line.store_3 or ''

    @api.depends('price_1', 'price_2')
    def _compute_price_diff(self):
        for line in self:
            if line.price_1 and line.price_2:
                line.price_diff_percent = ((line.price_2 - line.price_1) / line.price_1) * 100
            else:
                line.price_diff_percent = 0


class TotalinePriceHistory(models.Model):
    _name = 'totaline.price.history'
    _description = 'Price History'
    _order = 'search_date desc'

    search_id = fields.Many2one('totaline.price.search', string='Search', ondelete='cascade')
    search_line_id = fields.Many2one('totaline.price.search.line', string='Search Line', ondelete='set null')

    # Product info snapshot
    product_name = fields.Char(string='Product Name', required=True, index=True)
    brand = fields.Char(string='Brand', index=True)
    quantity = fields.Float(string='Quantity')
    unit = fields.Char(string='Unit')

    # Price data
    price_1 = fields.Float(string='Best Price')
    store_1 = fields.Char(string='Best Store')
    url_1 = fields.Char(string='Best URL')
    price_2 = fields.Float(string='2nd Price')
    store_2 = fields.Char(string='2nd Store')
    price_3 = fields.Float(string='3rd Price')
    store_3 = fields.Char(string='3rd Store')

    # Timing
    search_date = fields.Datetime(string='Search Date', required=True, index=True)

    # Computed for analytics
    price_per_unit = fields.Float(string='Price/Unit', compute='_compute_price_per_unit', store=True)

    @api.depends('price_1', 'quantity')
    def _compute_price_per_unit(self):
        for record in self:
            if record.quantity and record.price_1:
                record.price_per_unit = record.price_1 / record.quantity
            else:
                record.price_per_unit = record.price_1 or 0
