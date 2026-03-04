# -*- coding: utf-8 -*-
{
    'name': 'AP Reconciliation',
    'version': '17.0.1.0.0',
    'category': 'Afghan Post',
    'summary': 'Afghan Post Reconciliation — compare invoices with HesabPay records',
    'description': """
Afghan Post Reconciliation
==========================
- AP Reconciliation: load Odoo invoices and compare with HesabPay Excel
- Manual Reconciliation: compare two Excel files directly
- Custom security groups (User / Manager)
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
