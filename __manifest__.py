# -*- coding: utf-8 -*-
{
    'name': 'Financial Reconciliation',
    'version': '17.0.2.0.0',
    'category': 'Afghan Post',
    'summary': 'Financial Reconciliation — compare invoices with HesabPay records',
    'description': """
Financial Reconciliation
========================
- AP Reconciliation: load invoices and compare with HesabPay Excel
- Manual Reconciliation: compare two Excel files directly
- Custom security groups (User / Manager)
- User sees only own records, cannot delete
- Manager has full access to all records
- High-performance loading with batch create and DB indexing
- Standalone menu — independent from Accounting module
    """,
    'author': 'Custom',
    'depends': ['base', 'web', 'account'],
    'data': [
        'security/groups.xml',
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
