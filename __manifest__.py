# -*- coding: utf-8 -*-
{
    'name': 'HesabPay Reconciliation',
    'version': '17.0.1.0.0',
    'category': 'Accounting',
    'summary': 'Reconcile Odoo invoices with HesabPay Excel bank records',
    'description': """
HesabPay Reconciliation
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
        'security/ir.model.access.csv',
        'views/reconciliation_views.xml',
        'views/menu_views.xml',
    ],
    'assets': {
        'web.assets_backend': [
            'hesabpay_reconciliation/static/src/css/reconciliation.css',
            'hesabpay_reconciliation/static/src/js/reconciliation.js',
        ],
    },
    'installable': True,
    'application': True,
    'license': 'LGPL-3',
}
