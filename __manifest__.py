# -*- coding: utf-8 -*-
{
    'name': 'AP Reconciliation',
    'version': '17.0.1.0.0',
    'category': 'Accounting',
    'summary': 'Reconcile Afghan Post invoices with HesabPay bank records',
    'description': """
Afghan Post Reconciliation
=======================
- Filter invoices: Today / This Week / This Month / This Year / Custom date range
- Filter by journal (multi-select)
- Load Odoo invoices into a clean table with one click
- Upload HesabPay .xlsx file (Number | Customer | Total | Invoice Date)
- Auto-match by invoice number — red rows for mismatches, yellow for not found
- Export colour-coded .xlsx report
    """,
    'author': 'Custom',
    'depends': ['account'],
    'data': [
        'views/reconciliation_views.xml',
        'views/manual_reconciliation_views.xml',
        'views/menu_views.xml',
        'security/ir.model.access.csv',
    ],
    'assets': {
        'web.assets_backend': [
            'ap_reconciliation/static/src/css/reconciliation.css',
            'ap_reconciliation/static/src/js/reconciliation.js',
        ],
    },
    'installable': True,
    'application': True,
    'license': 'LGPL-3',
}
