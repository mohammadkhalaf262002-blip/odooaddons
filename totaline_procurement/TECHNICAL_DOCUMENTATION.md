# Totaline Procurement Module - Technical Documentation

## 1. Module Overview

| Property | Value |
|----------|-------|
| **Module Name** | `totaline_procurement` |
| **Version** | 18.0.2.0.0 |
| **License** | LGPL-3 |
| **Category** | Purchase |
| **Odoo Compatibility** | Odoo 18 |

---

## 2. File Structure

```
totaline_procurement/
├── __init__.py
├── __manifest__.py
├── models/
│   ├── __init__.py
│   └── price_search.py          # Main business logic (704 lines)
├── views/
│   ├── price_search_views.xml   # Form, List, Graph, Pivot views
│   └── menu_views.xml           # App menu configuration
├── security/
│   └── ir.model.access.csv      # Access control rules
├── data/
│   └── cron_data.xml            # Scheduled job configuration
└── static/
    └── src/css/
        └── totaline.css         # Custom styling
```

---

## 3. Dependencies

| Module | Purpose |
|--------|---------|
| `base` | Core Odoo functionality |
| `mail` | Chatter/messaging system |
| `purchase` | Purchase Order creation |
| `product` | Product management |

**Python Libraries:**

| Library | Purpose | Required |
|---------|---------|----------|
| `requests` | HTTP calls to n8n webhook | Yes (built-in) |
| `openpyxl` | Excel file parsing | Yes (needs installation) |

---

## 4. Database Models

### 4.1 `totaline.price.search` (Main Model)

| Field | Type | Description |
|-------|------|-------------|
| `name` | Char | Search name (auto-generated) |
| `state` | Selection | draft/searching/done/error |
| `excel_file` | Binary | Uploaded Excel file |
| `webhook_url` | Char | n8n webhook URL |
| `search_line_ids` | One2many | Product lines |
| `total_best_price` | Float | Computed total |
| `is_scheduled` | Boolean | Enable scheduling |
| `schedule_frequency` | Selection | daily/weekly/monthly |
| `next_scheduled_date` | Datetime | Next run time |

### 4.2 `totaline.price.search.line` (Product Lines)

| Field | Type | Description |
|-------|------|-------------|
| `product_name` | Char | Product name from Excel |
| `brand` | Char | Brand name |
| `quantity` | Float | Quantity needed |
| `price_1/2/3` | Float | Top 3 prices found |
| `store_1/2/3` | Char | Store names |
| `url_1/2/3` | Char | Product URLs |
| `selected_price` | Selection | Which price to use for PO |

### 4.3 `totaline.price.history` (Price Tracking)

| Field | Type | Description |
|-------|------|-------------|
| `product_name` | Char | Product name snapshot |
| `price_1` | Float | Best price at search time |
| `store_1` | Char | Best store |
| `search_date` | Datetime | When search was performed |
| `price_per_unit` | Float | Computed price/unit |

---

## 5. Security & Access Control

### 5.1 Access Rights (ir.model.access.csv)

| Model | Group | Read | Write | Create | Delete |
|-------|-------|------|-------|--------|--------|
| `totaline.price.search` | Purchase User | ✅ | ✅ | ✅ | ✅ |
| `totaline.price.search` | Purchase Manager | ✅ | ✅ | ✅ | ✅ |
| `totaline.price.search.line` | Purchase User | ✅ | ✅ | ✅ | ✅ |
| `totaline.price.search.line` | Purchase Manager | ✅ | ✅ | ✅ | ✅ |
| `totaline.price.history` | Purchase User | ✅ | ✅ | ✅ | ✅ |
| `totaline.price.history` | Purchase Manager | ✅ | ✅ | ✅ | ✅ |

**Who can access:**
- Only users in `purchase.group_purchase_user` or `purchase.group_purchase_manager` groups
- Regular users without Purchase access **cannot** see the module

### 5.2 Security Measures Implemented

| Security Feature | Status | Details |
|-----------------|--------|---------|
| **Role-based Access** | ✅ | Only Purchase users/managers |
| **Input Validation** | ✅ | Excel file validation, UserError exceptions |
| **SQL Injection Protection** | ✅ | Uses Odoo ORM (no raw SQL) |
| **XSS Protection** | ⚠️ | HTML fields use `sanitize=False` for links (low risk - data from trusted source) |
| **CSRF Protection** | ✅ | Handled by Odoo framework |
| **File Upload Validation** | ✅ | Only Excel files parsed by openpyxl |
| **Timeout Protection** | ✅ | 5-minute timeout on webhook calls |
| **Error Handling** | ✅ | Try/except blocks with logging |

### 5.3 Potential Security Considerations

| Risk | Level | Mitigation |
|------|-------|------------|
| **Webhook URL exposed** | Low | URL stored in DB, visible to Purchase users |
| **External API calls** | Medium | n8n webhook is external; use HTTPS |
| **HTML links in results** | Low | `sanitize=False` allows clickable links |
| **Excel macro execution** | None | openpyxl doesn't execute macros |

---

## 6. External Integrations

### 6.1 n8n Webhook Integration

**Data Flow:**
```
Odoo → n8n Webhook → Google Shopping API → n8n → Odoo
```

**Request Format (Odoo → n8n):**
```json
[
  {
    "id": "123",
    "name": "Çay Çaykur Altınbaş",
    "brand": "Çaykur",
    "quantity": 15,
    "unit": "paket",
    "description": "500gr paket"
  }
]
```

**Response Format (n8n → Odoo):**
```json
[
  {
    "id": "123",
    "product": "Çay Çaykur Altınbaş",
    "price_1": 2529.34,
    "store_1": "Trendyol",
    "url_1": "https://...",
    "price_2": 2530.00,
    "store_2": "Hepsiburada",
    "price_3": 2549.00,
    "store_3": "Amazon"
  }
]
```

### 6.2 Webhook Configuration for Production

**For Odoo.sh deployment, you'll need:**
1. n8n instance accessible from internet (not localhost)
2. HTTPS webhook URL
3. Update `webhook_url` field in each search record

---

## 7. Cron Job (Scheduled Searches)

| Property | Value |
|----------|-------|
| **Name** | Totaline: Run Scheduled Price Searches |
| **Model** | `totaline.price.search` |
| **Method** | `_cron_run_scheduled_searches()` |
| **Interval** | Every 1 hour |
| **Active** | Yes |

**What it does:**
1. Finds all searches with `is_scheduled=True` and `next_scheduled_date <= now`
2. Runs the price search automatically
3. Updates `last_scheduled_run`
4. Recomputes `next_scheduled_date`

---

## 8. Key Business Logic

### 8.1 Excel Parsing
- Supports Turkish/English column headers
- Auto-detects: Ürün Adı, Marka, Miktar, Birim, Açıklama
- Skips empty rows
- Shows notification on success

### 8.2 Price Search Flow
1. User uploads Excel → Products parsed into lines
2. User clicks "Search Prices" → Data sent to n8n webhook
3. n8n searches Google Shopping → Returns top 3 prices
4. Results displayed with store links
5. User selects preferred price for each product
6. User clicks "Create Purchase Order"

### 8.3 Purchase Order Creation
- Auto-creates vendor "Online Price Search" if not exists
- Creates/finds products in Odoo
- Includes store name and URL in PO line description
- Posts notification in chatter

---

## 9. Known Limitations

| Limitation | Impact | Workaround |
|------------|--------|------------|
| Generic product names may not find prices | Medium | Use specific names with brand+size |
| n8n must be accessible from Odoo server | High | Deploy n8n to cloud for Odoo.sh |
| No rate limiting | Low | Google Shopping API has its own limits |
| Prices in TRY (₺) only | Low | Can be modified for other currencies |

---

## 10. Deployment Checklist for Production

- [ ] **Install Purchase module** in Odoo if not installed
- [ ] **Deploy n8n to cloud** (not localhost)
- [ ] **Update webhook URL** to HTTPS production URL
- [ ] **Test with sample products** before full deployment
- [ ] **Configure user permissions** (Purchase User/Manager groups)
- [ ] **Set up scheduled searches** if needed
- [ ] **Monitor logs** for errors after deployment

---

## 11. Performance Considerations

| Metric | Value |
|--------|-------|
| **Webhook timeout** | 5 minutes (300 seconds) |
| **Cron interval** | 1 hour |
| **Max products per search** | No hard limit (depends on n8n) |
| **Price history storage** | Unlimited (grows over time) |

---

## 12. Support & Maintenance

**Logs Location:** Odoo server logs

**Key Log Messages:**
- `Received response from n8n: <class 'list'>` - Successful webhook response
- `Processing search results: <class 'list'>` - Results being processed
- `Running scheduled search: {name}` - Cron job executing
- `Scheduled search failed: {name}` - Cron job error

**Common Issues:**
1. **Prices showing 0.00** - Product name too generic, make it more specific
2. **Webhook timeout** - n8n workflow taking too long, check n8n logs
3. **Module not visible** - User not in Purchase User/Manager group

---

## 13. Version History

| Version | Date | Changes |
|---------|------|---------|
| 18.0.1.0.0 | Initial | Basic price search functionality |
| 18.0.2.0.0 | Current | Added price history, scheduled searches, analytics views |

---

*Documentation generated for Totaline Procurement Module v18.0.2.0.0*
