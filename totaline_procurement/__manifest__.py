# -*- coding: utf-8 -*-
{
    'name': 'Totalien Procurement',
    'version': '18.0.3.0.0',
    'category': 'Purchase',
    'summary': 'Price Comparison Tool for Procurement',
    'description': """
        Totaline Procurement System
        ===========================

        Compare prices across Turkish e-commerce stores.

        Features:
        ---------
        * Upload Excel file with product list
        * Direct SerpAPI integration (no n8n dependency)
        * Smart filtering with Turkish product matching
        * Compare up to 3 prices per product
        * Create Purchase Orders from results
        * Price history tracking with analytics
        * Scheduled automatic price searches
    """,
    'author': 'Totaline',
    'license': 'LGPL-3',
    'depends': [
        'base',
        'mail',
        'purchase',
        'product',
    ],
    'data': [
        'security/ir.model.access.csv',
        'views/price_search_views.xml',
        'views/menu_views.xml',
        'data/cron_data.xml',
    ],
    'assets': {
        'web.assets_backend': [
            'totaline_procurement/static/src/css/totaline.css',
        ],
    },
    'installable': True,
    'application': True,
    'auto_install': False,
}
