# -*- coding: utf-8 -*-

from odoo import models, fields, api
from odoo.exceptions import UserError
import requests
import json
import re
import math
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

# ============================================
# CONSTANTS (ported from n8n workflow)
# ============================================

TRUSTED_STORES = [
    "trendyol", "hepsiburada", "amazon.com.tr", "amazon", "migros",
    "migros sanalmarket", "a101", "bim", "sok", "carrefour",
    "carrefoursa", "gratis", "watsons", "avansas", "ofix",
    "bizim toptan", "metro", "makro", "koctas", "bauhaus",
    "n11", "gittigidiyor", "ciceksepeti", "pttavm", "morhipo",
    "teknosa", "mediamarkt", "vatan", "getir", "istegelsin",
    "rossmann", "eve", "flo", "boyner", "lcw", "defacto",
    "walmart", "costco", "sanalmarket", "marketpaketi", "happy center",
]

ACCESSORY_KEYWORDS = [
    # Cups and drinkware
    'bardak', 'bardagi', 'cup', 'mug', 'fincan',
    # Gifts and extras
    'hediye', 'hediyeli', 'gift', 'bonus', 'promosyon',
    # Accessories
    'kasik', 'spoon', 'tabak', 'plate',
    # Combo indicators
    '+ ', ' + ', 'ile birlikte', 'set icerigi', 'kullan at',
    # Free items
    'ucretsiz', 'bedava', 'free', 'gratis',
]

# Generic product names that need description-based search
GENERIC_PRODUCT_NAMES = [
    'rulo pecete', 'rulo pecete', 'kagit havlu', 'kagit havlu',
    'pecete', 'pecete', 'havlu',
]

# Product keyword categories for relevance checking
PRODUCT_KEYWORDS = {
    'turk kahvesi': ['turk kahvesi', 'turk kahvesi', 'turkish coffee'],
    'filtre kahve': ['filtre kahve', 'filter coffee'],
    'cay': ['cay', 'cay', 'tea'],
    'deterjan': ['deterjan', 'tablet', 'detergent'],
    'kagit havlu': ['kagit havlu', 'kagit havlu', 'paper towel', 'rulo'],
    'islak mendil': ['islak', 'mendil', 'havlu', 'wet wipe'],
}

# Wrong product categories to penalize
WRONG_CATEGORIES = ['bebek bezi', 'kedi mamasi', 'kopek mamasi', 'oyuncak']

# Turkish character mapping for normalization
TURKISH_CHAR_MAP = {
    '\u0131': 'i',  # ı → i
    '\u011f': 'g',  # ğ → g
    '\u00fc': 'u',  # ü → u
    '\u015f': 's',  # ş → s
    '\u00f6': 'o',  # ö → o
    '\u00e7': 'c',  # ç → c
    '\u0130': 'i',  # İ → i
    '\u011e': 'g',  # Ğ → g
    '\u00dc': 'u',  # Ü → u
    '\u015e': 's',  # Ş → s
    '\u00d6': 'o',  # Ö → o
    '\u00c7': 'c',  # Ç → c
}


# ============================================
# HELPER FUNCTIONS (ported from n8n JS)
# ============================================

def normalize_turkish(text):
    """Normalize Turkish characters to ASCII equivalents."""
    if not text:
        return ''
    result = str(text).lower()
    for turkish_char, ascii_char in TURKISH_CHAR_MAP.items():
        result = result.replace(turkish_char, ascii_char)
    return result


def is_trusted_store(source):
    """Check if a store is in the trusted stores list."""
    if not source:
        return False
    normalized = normalize_turkish(source)
    return any(normalize_turkish(store) in normalized for store in TRUSTED_STORES)


def is_accessory_bundle(title):
    """Check if title contains accessory/bundle indicators."""
    if not title:
        return False
    normalized_title = normalize_turkish(title)

    for keyword in ACCESSORY_KEYWORDS:
        if normalize_turkish(keyword) in normalized_title:
            return True

    # Check for "X + Y" pattern (combo products) e.g. "100gr + 50 bardak"
    if re.search(r'\d+\s*(?:gr|g|ml)\s*\+\s*\d+', title, re.IGNORECASE):
        return True

    return False


def is_relevant_product(title, searched_name, searched_brand):
    """Validate product title actually matches what we're searching for."""
    if not title:
        return False

    normalized_title = normalize_turkish(title)
    normalized_name = normalize_turkish(searched_name)
    normalized_brand = normalize_turkish(searched_brand)

    # Must contain the brand (if specified)
    if normalized_brand and len(normalized_brand) > 2:
        brand_words = [w for w in normalized_brand.split() if len(w) > 2]
        brand_match = any(word in normalized_title for word in brand_words)
        if not brand_match:
            return False

    # Must contain key product words
    for key, keywords in PRODUCT_KEYWORDS.items():
        if normalize_turkish(key) in normalized_name:
            has_keyword = any(normalize_turkish(kw) in normalized_title for kw in keywords)
            if not has_keyword:
                return False
            break

    return True


def extract_quantity_from_title(title, unit, product_name):
    """
    Extract quantity from product title.
    Returns dict with qty, confident, is_bundle, pattern_used.
    """
    if not title:
        return {'qty': 1, 'confident': False}

    normalized_title = normalize_turkish(title)
    original_title = title.lower()
    detected_qty = 1
    confident = False

    # ========== FIRST: Check if this is an accessory bundle ==========
    if is_accessory_bundle(title):
        before_plus = title.split('+')[0]
        main_match = re.search(r'(\d+)\s*(?:gr|g|kg|ml|l)(?:\s|$)', before_plus, re.IGNORECASE)
        if main_match:
            return {'qty': 1, 'confident': True, 'is_bundle': True}
        return {'qty': 1, 'confident': False, 'is_bundle': True}

    # ========== UNIT-SPECIFIC PATTERNS ==========

    # For KG unit
    if unit == 'kg':
        kg_patterns = [
            r'(\d+(?:[.,]\d+)?)\s*kg(?:r|ram)?',
            r'(\d+(?:[.,]\d+)?)\s*kilo',
        ]
        for pattern in kg_patterns:
            match = re.search(pattern, normalized_title, re.IGNORECASE)
            if match:
                detected_qty = float(match.group(1).replace(',', '.'))
                confident = True
                break

        if not confident:
            gram_match = re.search(r'(\d+)\s*gr(?:am)?', normalized_title, re.IGNORECASE)
            if gram_match and 'kg' not in normalized_title:
                detected_qty = int(gram_match.group(1)) / 1000
                confident = True

        return {'qty': detected_qty, 'confident': confident}

    # For LITRE unit
    if unit in ('litre', 'lt'):
        litre_patterns = [
            r'(\d+(?:[.,]\d+)?)\s*l(?:t|itre)?(?!\w)',
            r'(\d+(?:[.,]\d+)?)\s*litre',
        ]
        for pattern in litre_patterns:
            match = re.search(pattern, normalized_title, re.IGNORECASE)
            if match:
                detected_qty = float(match.group(1).replace(',', '.'))
                confident = True
                break

        if not confident:
            ml_match = re.search(r'(\d+)\s*ml', normalized_title, re.IGNORECASE)
            if ml_match:
                detected_qty = int(ml_match.group(1)) / 1000
                confident = True

        return {'qty': detected_qty, 'confident': confident}

    # ========== BULK PACK PATTERNS (High confidence) ==========
    bulk_patterns = [
        # "x 50 Adet" or "X50 Adet"
        {'pattern': r'[x\u00d7]\s*(\d+)\s*adet', 'name': 'X N adet', 'multiply': False},
        # "50 Adet" at start or after size
        {'pattern': r'(?:^|\d+\s*(?:gr|g|ml|kg)\s*)(\d+)\s*adet', 'name': 'N adet after size', 'multiply': False},
        # "50'li Set" or "50'li Paket"
        {'pattern': r"(\d+)\s*['\u00b4\u2018\u2019`]?\s*l[iiuu]\s*(?:set|paket|koli)", 'name': "N'li set/paket", 'multiply': False},
        # "(50 Adet)" or "(50'li)" in parentheses
        {'pattern': r"\((\d+)\s*(?:adet|'?l[iiuu])\)", 'name': 'parentheses count', 'multiply': False},
        # "100gr x 50" - size times count
        {'pattern': r'\d+\s*(?:gr|g|ml)\s*[x\u00d7]\s*(\d+)(?:\s|$|adet)', 'name': 'size x count', 'multiply': False},
        # "2 Koli*25" or "2 Koli x 25"
        {'pattern': r'(\d+)\s*koli\s*[x\u00d7\*]\s*(\d+)', 'name': 'koli pattern', 'multiply': True},
    ]

    for bp in bulk_patterns:
        match = re.search(bp['pattern'], original_title, re.IGNORECASE)
        if not match:
            match = re.search(bp['pattern'], normalized_title, re.IGNORECASE)
        if match:
            if bp['multiply'] and match.lastindex and match.lastindex >= 2:
                detected_qty = int(match.group(1)) * int(match.group(2))
            else:
                detected_qty = int(match.group(1))

            if 2 <= detected_qty <= 500:
                return {'qty': detected_qty, 'confident': True, 'pattern_used': bp['name']}

    # ========== STANDARD PACK PATTERNS (Medium confidence) ==========
    pack_patterns = [
        # Turkish: 10'lu, 8'li
        {'pattern': r"(\d+)\s*['\u00b4\u2018\u2019`]\s*l[iiuu]", 'name': 'Turkish apostrophe'},
        # Turkish without apostrophe: 10lu, 8li (but NOT after + sign)
        {'pattern': r'(?<!\+)(?<!\+ )(\d+)\s*l[iiuu](?:\s|$|,|\)|paket|koli|set|kutu)', 'name': 'Turkish no apostrophe'},
        # "N adet" standalone
        {'pattern': r'(?:^|\s)(\d+)\s*adet(?:\s|$|,|\))', 'name': 'N adet'},
        # Tablet/Capsule
        {'pattern': r'(\d+)\s*(?:tablet|kapsul|kapsul)', 'name': 'Tablet'},
        # Rulo
        {'pattern': r'(\d+)\s*(?:rulo|roll)', 'name': 'Rulo'},
    ]

    for pp in pack_patterns:
        match = re.search(pp['pattern'], original_title, re.IGNORECASE)
        if not match:
            match = re.search(pp['pattern'], normalized_title, re.IGNORECASE)
        if match:
            qty = int(match.group(1))
            if 2 <= qty <= 500:
                return {'qty': qty, 'confident': True, 'pattern_used': pp['name']}

    return {'qty': detected_qty, 'confident': confident}


def calculate_relevance_score(title, source, original_name, original_brand):
    """Calculate relevance score for a search result."""
    if not title:
        return 0

    normalized_title = normalize_turkish(title)
    normalized_name = normalize_turkish(original_name)
    normalized_brand = normalize_turkish(original_brand)

    score = 100

    # PENALTY: Accessory bundles
    if is_accessory_bundle(title):
        score -= 40

    # Check if product name matches
    name_words = [w for w in normalized_name.split() if len(w) > 2]
    name_match_count = sum(1 for w in name_words if w in normalized_title)

    if name_words:
        name_match_ratio = name_match_count / len(name_words)
        if name_match_ratio < 0.5:
            score -= 50
        elif name_match_ratio < 1:
            score -= 20

    # Check if brand matches
    if normalized_brand and len(normalized_brand) > 2:
        brand_words = normalized_brand.split()
        brand_match = any(
            len(word) > 2 and word in normalized_title
            for word in brand_words
        )
        if not brand_match:
            score -= 35

    # Penalty for wrong product categories
    for cat in WRONG_CATEGORIES:
        normalized_cat = normalize_turkish(cat)
        if normalized_cat in normalized_title and normalize_turkish(cat.split()[0]) not in normalized_name:
            score -= 60

    # BONUS: Bulk indicators
    bulk_indicators = ['toptan', 'koli', 'set', 'paket', 'x 50', 'x50', '50 adet']
    for indicator in bulk_indicators:
        if normalize_turkish(indicator) in normalized_title:
            score += 15
            break

    # Bonus for trusted stores
    if is_trusted_store(source):
        score += 10

    return max(0, score)


def build_search_query(product_name, brand, description, unit, quantity):
    """
    Build an optimized search query for Google Shopping.
    Ported from n8n Parse Products node.
    """
    name = (product_name or '').strip()
    brand = (brand or '').strip()
    description = (description or '').strip()
    unit = (unit or 'adet').lower().strip()
    quantity = quantity or 1

    name_lower = normalize_turkish(name)

    # Check if this is a generic product that needs description-based search
    use_description = False
    for generic in GENERIC_PRODUCT_NAMES:
        if normalize_turkish(generic) in name_lower and description:
            use_description = True
            break

    if use_description:
        search_query = f"{description} {brand}".strip()
    else:
        search_query = f"{name} {brand}".strip()

    # Add unit context for bulk searches
    if quantity > 1:
        unit_context_map = {
            'kg': 'kg toptan',
            'litre': 'litre toptan',
            'lt': 'litre toptan',
            'koli': 'koli toptan',
            'paket': 'paket',
            'adet': 'toptan',
        }
        context = unit_context_map.get(unit, 'toptan')
        search_query += f" {context}"

    # Extract size hints from description
    if description:
        size_match = re.search(r'(\d+)\s*(?:gr|g|kg|ml|lt|l)\b', description, re.IGNORECASE)
        if size_match and size_match.group(0).lower() not in search_query.lower():
            search_query += f" {size_match.group(0)}"

    # Add "fiyat" to help Google Shopping results
    search_query += " fiyat"

    return search_query.strip()


# ============================================
# ODOO MODELS
# ============================================

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
            for col_idx, cell in enumerate(sheet[1], 1):
                if cell.value:
                    header_name = str(cell.value).strip().lower()
                    headers[header_name] = col_idx

            # Map Turkish/English column names
            column_mapping = {
                'product_name': ['\u00fcr\u00fcn ad\u0131', 'urun adi', 'product name', 'product', '\u00fcr\u00fcn', 'urun'],
                'brand': ['marka', 'brand'],
                'quantity': ['miktar', 'quantity', 'qty', 'adet'],
                'unit': ['birim', 'unit'],
                'description': ['a\u00e7\u0131klama', 'aciklama', 'description', 'desc', 'detay'],
            }

            # Find column indices
            col_indices = {}
            for field_name, names in column_mapping.items():
                for col_name in names:
                    if col_name in headers:
                        col_indices[field_name] = headers[col_name]
                        break

            if 'product_name' not in col_indices:
                raise UserError('Could not find product name column. Expected: "\u00dcr\u00fcn Ad\u0131" or "Product Name"')

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

    # ============================================
    # SERPAPI DIRECT INTEGRATION
    # ============================================

    def _get_serpapi_key(self):
        """Get SerpAPI key from Odoo System Parameters."""
        api_key = self.env['ir.config_parameter'].sudo().get_param('totaline.serpapi_key', '')
        if not api_key:
            raise UserError(
                'SerpAPI key not configured.\n\n'
                'Go to Settings > Technical > Parameters > System Parameters\n'
                'and create a parameter:\n'
                '  Key: totaline.serpapi_key\n'
                '  Value: your SerpAPI key'
            )
        return api_key

    def _search_serpapi(self, search_query, api_key):
        """Call SerpAPI Google Shopping directly and return results."""
        params = {
            'engine': 'google_shopping',
            'q': search_query,
            'location': 'Turkey',
            'gl': 'tr',
            'hl': 'tr',
            'api_key': api_key,
        }

        try:
            response = requests.get(
                'https://serpapi.com/search.json',
                params=params,
                timeout=30,
            )

            if response.status_code == 200:
                return response.json()
            else:
                _logger.warning(f'SerpAPI returned status {response.status_code}: {response.text[:200]}')
                return {}

        except requests.exceptions.Timeout:
            _logger.warning(f'SerpAPI timeout for query: {search_query}')
            return {}
        except Exception as e:
            _logger.exception(f'SerpAPI request failed: {e}')
            return {}

    def _format_search_results(self, serp_results, product_name, brand, quantity, unit, description):
        """
        Process SerpAPI results and return top 3 price results.
        Ported from n8n Format Response V3 node.

        Returns dict with: price_1, store_1, url_1, price_2, ..., price_3, ..., warning
        """
        original_name = product_name or ''
        original_brand = brand or ''
        original_quantity = int(quantity or 1)
        original_unit = (unit or 'adet').lower()

        shopping_results = serp_results.get('shopping_results', [])
        processed_results = []

        for item in shopping_results:
            title = item.get('title', '')
            price = 0
            try:
                price = float(item.get('extracted_price', 0) or 0)
            except (ValueError, TypeError):
                continue

            source = item.get('source', 'Unknown')
            link = item.get('link', '')

            if price <= 0:
                continue

            # Check relevance
            if not is_relevant_product(title, original_name, original_brand):
                continue

            # Check if accessory bundle
            is_bundle_product = is_accessory_bundle(title)

            # Extract quantity from title
            quantity_info = extract_quantity_from_title(title, original_unit, original_name)
            detected_qty = quantity_info.get('qty', 1)
            confident = quantity_info.get('confident', False)

            # Calculate multiplier
            multiplier = 1
            warning = None

            if is_bundle_product:
                if original_quantity > 1:
                    multiplier = original_quantity
                    warning = f"\u26a0\ufe0f Hediye setli \u00fcr\u00fcn - {multiplier} adet al\u0131nmal\u0131"
            elif original_quantity > 1:
                if detected_qty < original_quantity:
                    multiplier = math.ceil(original_quantity / detected_qty)
                    if detected_qty == 1:
                        warning = f"\u26a0\ufe0f Tekli sat\u0131\u015f - {multiplier} adet al\u0131nmal\u0131"
                    else:
                        warning = f"\u26a0\ufe0f {detected_qty}'li paket - {multiplier} paket al\u0131nmal\u0131"
                elif detected_qty > original_quantity:
                    warning = f"\u2139\ufe0f Pakette {detected_qty} {original_unit} var ({original_quantity} istendi)"

                # Safety check for unconfident detections
                if not confident and original_quantity > 1 and multiplier == 1:
                    multiplier = original_quantity
                    warning = f"\u26a0\ufe0f Tekli sat\u0131\u015f varsay\u0131ld\u0131 - {multiplier} adet al\u0131nmal\u0131"

            total_price = price * multiplier
            relevance_score = calculate_relevance_score(title, source, original_name, original_brand)
            trusted = is_trusted_store(source)

            processed_results.append({
                'title': title,
                'source': source,
                'link': link,
                'original_price': price,
                'detected_qty': detected_qty,
                'confident': confident,
                'is_bundle': is_bundle_product,
                'multiplier': multiplier,
                'total_price': total_price,
                'warning': warning,
                'relevance_score': relevance_score,
                'trusted': trusted,
            })

        # ============================================
        # SELECT TOP 3 RESULTS
        # ============================================

        # Filter trusted results with good relevance
        trusted_results = [
            r for r in processed_results
            if r['trusted'] and r['relevance_score'] >= 40
        ]
        # Sort: non-bundles first, then by relevance (if big diff), then by price
        trusted_results.sort(key=lambda r: (
            r['is_bundle'],  # False < True, so non-bundles first
            -(r['relevance_score'] if abs(r.get('relevance_score', 0)) > 25 else 0),
            r['total_price'],
        ))

        untrusted_results = [r for r in processed_results if not r['trusted']]
        untrusted_results.sort(key=lambda r: r['total_price'])

        # Combine: trusted first, then untrusted
        all_sorted = trusted_results + untrusted_results
        top_3 = all_sorted[:3]

        # Build result dict
        result = {
            'price_1': top_3[0]['total_price'] if len(top_3) > 0 else 0,
            'store_1': top_3[0]['source'] if len(top_3) > 0 else '',
            'url_1': top_3[0]['link'] if len(top_3) > 0 else '',
            'price_2': top_3[1]['total_price'] if len(top_3) > 1 else 0,
            'store_2': top_3[1]['source'] if len(top_3) > 1 else '',
            'url_2': top_3[1]['link'] if len(top_3) > 1 else '',
            'price_3': top_3[2]['total_price'] if len(top_3) > 2 else 0,
            'store_3': top_3[2]['source'] if len(top_3) > 2 else '',
            'url_3': top_3[2]['link'] if len(top_3) > 2 else '',
            'warning': '',
        }

        # Collect warnings from top results
        warnings = []
        for r in top_3:
            if r.get('warning'):
                warnings.append(r['warning'])
                break  # Only first warning

        if not top_3:
            warnings.append('Sonu\u00e7 bulunamad\u0131')  # No results found

        result['warning'] = warnings[0] if warnings else ''

        _logger.info(
            f'Product "{original_name}": found {len(shopping_results)} raw, '
            f'{len(processed_results)} processed, {len(trusted_results)} trusted, '
            f'top price: {result["price_1"]}'
        )

        return result

    # ============================================
    # MAIN SEARCH ACTION
    # ============================================

    def action_search_prices(self):
        """Search prices directly via SerpAPI (no n8n dependency)."""
        self.ensure_one()

        if not self.search_line_ids:
            self.state = 'error'
            self.error_message = 'No products to search. Please add products first.'
            return

        self.state = 'searching'
        self.search_date = fields.Datetime.now()
        self.error_message = False

        try:
            api_key = self._get_serpapi_key()

            success_count = 0
            error_count = 0

            for line in self.search_line_ids:
                try:
                    # Build search query (ported from n8n Parse Products)
                    search_query = build_search_query(
                        product_name=line.product_name,
                        brand=line.brand,
                        description=line.description,
                        unit=line.unit,
                        quantity=line.quantity,
                    )

                    _logger.info(f'Searching: "{search_query}" for product "{line.product_name}"')

                    # Call SerpAPI directly
                    serp_results = self._search_serpapi(search_query, api_key)

                    # Format results (ported from n8n Format Response V3)
                    formatted = self._format_search_results(
                        serp_results=serp_results,
                        product_name=line.product_name,
                        brand=line.brand,
                        quantity=line.quantity,
                        unit=line.unit,
                        description=line.description,
                    )

                    # Write results to the line
                    line.write({
                        'price_1': formatted.get('price_1', 0),
                        'store_1': formatted.get('store_1', ''),
                        'url_1': formatted.get('url_1', ''),
                        'price_2': formatted.get('price_2', 0),
                        'store_2': formatted.get('store_2', ''),
                        'url_2': formatted.get('url_2', ''),
                        'price_3': formatted.get('price_3', 0),
                        'store_3': formatted.get('store_3', ''),
                        'url_3': formatted.get('url_3', ''),
                        'warning': formatted.get('warning', ''),
                    })

                    success_count += 1

                except Exception as e:
                    error_count += 1
                    _logger.exception(f'Failed to search product: {line.product_name}')
                    line.write({
                        'warning': f'Arama hatas\u0131: {str(e)[:100]}',
                    })

            # Save to price history
            self._save_price_history()

            if error_count == 0:
                self.state = 'done'
            elif success_count > 0:
                self.state = 'done'
                self.error_message = f'{error_count} product(s) had search errors. {success_count} succeeded.'
            else:
                self.state = 'error'
                self.error_message = 'All product searches failed.'

        except UserError:
            self.state = 'error'
            raise
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
        """
        Legacy: Update search lines with results from n8n.
        Kept for backward compatibility with any existing n8n workflows.
        """
        _logger.info(f'Processing search results (legacy): {type(results)}')

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
            notes.append(f"\u2022 {line.product_name}: \u20ba{selected_price:,.2f} from {selected_store} (Price {line.selected_price})")

        # Create the Purchase Order
        PurchaseOrder = self.env['purchase.order']
        po = PurchaseOrder.create({
            'partner_id': vendor.id,
            'date_order': fields.Datetime.now(),
            'notes': f"Created from Price Search: {self.name}\n\nProducts:\n" + "\n".join(notes),
            'order_line': po_lines,
        })

        self.message_post(
            body=f"Purchase Order <a href='#' data-oe-model='purchase.order' data-oe-id='{po.id}'>{po.name}</a> created with {len(po_lines)} products. Total: \u20ba{total_selected:,.2f}",
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
        ('1', '1. Fiyat (En \u0130yi)'),
        ('2', '2. Fiyat'),
        ('3', '3. Fiyat'),
    ], string='Se\u00e7ilen Fiyat', default='1', help='Select which price to use for Purchase Order')

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
                line.store_link_1 = f'<a href="{line.url_1}" target="_blank">{line.store_1} \u2197</a>'
            else:
                line.store_link_1 = line.store_1 or ''

            # Store 2
            if line.url_2 and line.store_2:
                line.store_link_2 = f'<a href="{line.url_2}" target="_blank">{line.store_2} \u2197</a>'
            else:
                line.store_link_2 = line.store_2 or ''

            # Store 3
            if line.url_3 and line.store_3:
                line.store_link_3 = f'<a href="{line.url_3}" target="_blank">{line.store_3} \u2197</a>'
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
