# -*- coding: utf-8 -*-
from odoo import models, fields, api
from odoo.exceptions import UserError
import base64
import io
import logging

_logger = logging.getLogger(__name__)

try:
    import openpyxl
except ImportError:
    openpyxl = None


class HesabpayReconciliation(models.TransientModel):
    _name        = 'hesabpay.reconciliation'
    _description = 'HesabPay Reconciliation'

    date_filter = fields.Selection([
        ('today',      'Today'),
        ('this_week',  'This Week'),
        ('this_month', 'This Month'),
        ('this_year',  'This Year'),
        ('custom',     'Custom Range'),
    ], string='Date Filter', default='this_month', required=True)

    date_from   = fields.Date(string='Date From')
    date_to     = fields.Date(string='Date To')

    journal_ids = fields.Many2many(
        'account.journal',
        string='Journals',
        domain=[('type', 'in', ['sale', 'purchase', 'bank', 'cash'])],
    )

    excel_file     = fields.Binary(string='HesabPay Excel File', attachment=False)
    excel_filename = fields.Char(string='Filename')

    # All lines
    line_ids = fields.One2many(
        'hesabpay.reconciliation.line',
        'reconciliation_id',
        string='All Lines',
    )

    # Mismatch lines only — separate One2many with domain in Python
    mismatch_line_ids = fields.One2many(
        'hesabpay.reconciliation.line',
        'reconciliation_id',
        string='Mismatches',
        domain=lambda self: [
            '|', ('name_mismatch', '=', True),
            '|', ('amount_mismatch', '=', True),
                 ('hp_found', '=', False),
        ],
    )

    state = fields.Selection([
        ('draft',      'Draft'),
        ('loaded',     'Odoo Data Loaded'),
        ('reconciled', 'Reconciled'),
    ], default='draft')

    # ── Actions ───────────────────────────────────────────────────────────────

    def action_load_odoo_data(self):
        self.ensure_one()
        invoices = self.env['account.move'].search(
            self._build_domain(), order='name asc', limit=5000
        )
        self.line_ids.unlink()

        vals_list = []
        for inv in invoices:
            vals_list.append({
                'reconciliation_id': self.id,
                'invoice_number':    inv.name,
                'partner_name':      inv.partner_id.name or '',
                'invoice_date':      inv.invoice_date,
                'total_amount':      inv.amount_total,
                'currency_id':       inv.currency_id.id,
                'move_id':           inv.id,
                'status':            inv.state,
                'payment_state':     inv.payment_state or '',
            })
        if vals_list:
            self.env['hesabpay.reconciliation.line'].create(vals_list)

        self.state = 'loaded'
        return self._reload()

    def action_import_hesabpay(self):
        self.ensure_one()
        if not self.excel_file:
            raise UserError('Please upload a HesabPay Excel file first.')
        if not openpyxl:
            raise UserError('openpyxl is not installed. Run: pip3 install openpyxl')

        file_data = base64.b64decode(self.excel_file)
        wb = openpyxl.load_workbook(io.BytesIO(file_data), data_only=True)
        ws = wb.active

        hp_data = {}
        header_skipped = False
        for row in ws.iter_rows(values_only=True):
            if not header_skipped:
                header_skipped = True
                continue
            if not any(row):
                continue
            try:
                inv_no   = str(row[0]).strip() if row[0] is not None else ''
                customer = str(row[1]).strip() if row[1] is not None else ''
                total    = float(row[2])        if row[2] is not None else 0.0
                inv_date = row[3]
                if hasattr(inv_date, 'date'):
                    inv_date = inv_date.date()
                elif isinstance(inv_date, str):
                    from datetime import datetime
                    for fmt in ('%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y'):
                        try:
                            inv_date = datetime.strptime(inv_date, fmt).date()
                            break
                        except ValueError:
                            inv_date = None
                if inv_no:
                    hp_data[inv_no] = {'customer': customer, 'total': total, 'date': inv_date}
            except Exception as e:
                _logger.warning('Skipping row: %s', e)

        for line in self.line_ids:
            hp = hp_data.get(line.invoice_number)
            if hp:
                diff = line.total_amount - hp['total']
                line.write({
                    'hp_partner_name':   hp['customer'],
                    'hp_total_amount':   hp['total'],
                    'hp_invoice_date':   hp['date'],
                    'hp_found':          True,
                    'amount_difference': diff,
                    'name_mismatch':     line.partner_name.strip().lower() != hp['customer'].strip().lower(),
                    'amount_mismatch':   abs(diff) > 0.01,
                })
            else:
                line.write({
                    'hp_found':          False,
                    'hp_partner_name':   '',
                    'hp_total_amount':   0.0,
                    'amount_difference': 0.0,
                    'name_mismatch':     False,
                    'amount_mismatch':   False,
                })

        self.state = 'reconciled'
        return self._reload()

    def action_export_report(self):
        self.ensure_one()
        if not openpyxl:
            raise UserError('openpyxl is not installed.')

        from openpyxl.styles import PatternFill, Font, Alignment
        wb  = openpyxl.Workbook()
        ws  = wb.active
        ws.title = 'Reconciliation'

        hdr_fill      = PatternFill('solid', fgColor='4472C4')
        mismatch_fill = PatternFill('solid', fgColor='FFCCCC')
        ok_fill       = PatternFill('solid', fgColor='CCFFCC')
        nf_fill       = PatternFill('solid', fgColor='FFF2CC')
        hdr_font      = Font(bold=True, color='FFFFFF')

        headers = ['Invoice #', 'Odoo Customer', 'Odoo Date', 'Odoo Total',
                   'HP Customer', 'HP Date', 'HP Total',
                   'Difference', 'Name Match?', 'Amount Match?', 'Result']
        ws.append(headers)
        for i, _ in enumerate(headers, 1):
            c = ws.cell(row=1, column=i)
            c.fill = hdr_fill
            c.font = hdr_font
            c.alignment = Alignment(horizontal='center')

        for line in self.line_ids:
            if not line.hp_found:
                result, fill = 'NOT FOUND', nf_fill
            elif line.name_mismatch or line.amount_mismatch:
                result, fill = 'MISMATCH', mismatch_fill
            else:
                result, fill = 'OK', ok_fill

            ws.append([
                line.invoice_number, line.partner_name,
                str(line.invoice_date or ''), line.total_amount,
                line.hp_partner_name, str(line.hp_invoice_date or ''),
                line.hp_total_amount, line.amount_difference,
                'NO' if line.name_mismatch  else 'YES',
                'NO' if line.amount_mismatch else 'YES',
                result,
            ])
            r = ws.max_row
            for col in range(1, len(headers) + 1):
                ws.cell(row=r, column=col).fill = fill

        for col in ws.columns:
            w = max((len(str(c.value or '')) for c in col), default=10)
            ws.column_dimensions[col[0].column_letter].width = min(w + 4, 45)

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)

        attach = self.env['ir.attachment'].create({
            'name':     'hesabpay_reconciliation_report.xlsx',
            'type':     'binary',
            'datas':    base64.b64encode(buf.read()),
            'mimetype': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        })
        return {
            'type':   'ir.actions.act_url',
            'url':    f'/web/content/{attach.id}?download=true',
            'target': 'self',
        }

    def _reload(self):
        return {
            'type':      'ir.actions.act_window',
            'res_model': 'hesabpay.reconciliation',
            'res_id':    self.id,
            'view_mode': 'form',
            'target':    'current',
        }

    def _build_domain(self):
        domain = [
            ('move_type', 'in', ['out_invoice', 'out_refund', 'in_invoice', 'in_refund']),
            ('state', '!=', 'cancel'),
        ]
        today = fields.Date.today()
        if self.date_filter == 'today':
            domain += [('invoice_date', '=', today)]
        elif self.date_filter == 'this_week':
            from datetime import timedelta
            week_start = today - timedelta(days=today.weekday())
            domain += [('invoice_date', '>=', week_start),
                       ('invoice_date', '<=', week_start + timedelta(days=6))]
        elif self.date_filter == 'this_month':
            domain += [('invoice_date', '>=', today.replace(day=1)),
                       ('invoice_date', '<=', today)]
        elif self.date_filter == 'this_year':
            domain += [('invoice_date', '>=', today.replace(month=1, day=1)),
                       ('invoice_date', '<=', today)]
        elif self.date_filter == 'custom':
            if self.date_from: domain += [('invoice_date', '>=', self.date_from)]
            if self.date_to:   domain += [('invoice_date', '<=', self.date_to)]
        if self.journal_ids:
            domain += [('journal_id', 'in', self.journal_ids.ids)]
        return domain


class HesabpayReconciliationLine(models.TransientModel):
    _name        = 'hesabpay.reconciliation.line'
    _description = 'Reconciliation Line'
    _order       = 'invoice_number asc'

    reconciliation_id = fields.Many2one(
        'hesabpay.reconciliation', ondelete='cascade', required=True
    )

    # Odoo side
    move_id        = fields.Many2one('account.move', ondelete='set null')
    invoice_number = fields.Char(string='Invoice #')
    partner_name   = fields.Char(string='Odoo Customer')
    invoice_date   = fields.Date(string='Odoo Date')
    total_amount   = fields.Float(string='Odoo Total',  digits=(16, 2))
    currency_id    = fields.Many2one('res.currency')
    status         = fields.Char(string='Status')
    payment_state  = fields.Char(string='Payment')

    # HesabPay side
    hp_found        = fields.Boolean(string='In HesabPay', default=False)
    hp_partner_name = fields.Char(string='HP Customer')
    hp_invoice_date = fields.Date(string='HP Date')
    hp_total_amount = fields.Float(string='HP Total', digits=(16, 2))

    # Difference
    amount_difference = fields.Float(string='Difference',     digits=(16, 2))
    name_mismatch     = fields.Boolean(string='Name Mismatch',   default=False)
    amount_mismatch   = fields.Boolean(string='Amount Mismatch', default=False)
