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


class FinReconciliation(models.Model):
    _name        = 'fin.reconciliation'
    _description = 'Financial Reconciliation'
    _order       = 'id desc'
    _rec_name    = 'name'

    # ── Record rule support: users see only their own records ─────────────────
    # Managers bypass this via record rules defined in groups.xml
    create_uid = fields.Many2one('res.users', string='Created By', readonly=True)

    name = fields.Char(
        string='Reference',
        required=True,
        default='New Reconciliation',
        index=True,
    )

    date_filter = fields.Selection([
        ('today',      'Today'),
        ('this_week',  'This Week'),
        ('this_month', 'This Month'),
        ('this_year',  'This Year'),
        ('custom',     'Custom Range'),
    ], string='Date Filter', default='this_month', required=True)

    date_from = fields.Date(string='Date From')
    date_to   = fields.Date(string='Date To')

    journal_ids = fields.Many2many(
        'account.journal',
        'fin_recon_journal_rel',
        'recon_id',
        'journal_id',
        string='Journals',
        domain=[('type', 'in', ['sale', 'purchase', 'bank', 'cash'])],
    )

    invoice_number_filter = fields.Char(
        string='Invoice # Contains',
        help='Filter invoices by number keyword e.g. "mofa", "MOhe", "260101"'
    )

    excel_file     = fields.Binary(string='HesabPay Excel File', attachment=True)
    excel_filename = fields.Char(string='Filename')

    line_ids = fields.One2many(
        'fin.reconciliation.line',
        'reconciliation_id',
        string='All Lines',
    )

    mismatch_line_ids = fields.One2many(
        'fin.reconciliation.line',
        'reconciliation_id',
        string='Mismatches',
        domain=[
            '|', ('name_mismatch', '=', True),
            '|', ('amount_mismatch', '=', True),
                 ('hp_found', '=', False),
        ],
    )

    state = fields.Selection([
        ('draft',      'Draft'),
        ('loaded',     'Invoices Loaded'),
        ('reconciled', 'Reconciled'),
    ], default='draft', string='State', required=True, index=True)

    total_lines     = fields.Integer(compute='_compute_summary', store=True)
    matched_lines   = fields.Integer(compute='_compute_summary', store=True)
    mismatch_lines  = fields.Integer(compute='_compute_summary', store=True)
    not_found_lines = fields.Integer(compute='_compute_summary', store=True)

    @api.depends('line_ids.hp_found', 'line_ids.name_mismatch', 'line_ids.amount_mismatch')
    def _compute_summary(self):
        for rec in self:
            lines = rec.line_ids
            rec.total_lines     = len(lines)
            rec.not_found_lines = len(lines.filtered(lambda l: not l.hp_found))
            rec.mismatch_lines  = len(lines.filtered(
                lambda l: l.hp_found and (l.name_mismatch or l.amount_mismatch)
            ))
            rec.matched_lines = len(lines.filtered(
                lambda l: l.hp_found and not l.name_mismatch and not l.amount_mismatch
            ))

    # ── Actions ───────────────────────────────────────────────────────────────

    def action_load_data(self):
        """Load invoices from Odoo with high-performance batch processing."""
        self.ensure_one()
        from datetime import timedelta

        today = fields.Date.today()

        if self.date_filter == 'today':
            d_from = d_to = today
        elif self.date_filter == 'this_week':
            d_from = today - timedelta(days=today.weekday())
            d_to   = d_from + timedelta(days=6)
        elif self.date_filter == 'this_month':
            d_from = today.replace(day=1)
            d_to   = today
        elif self.date_filter == 'this_year':
            d_from = today.replace(month=1, day=1)
            d_to   = today
        elif self.date_filter == 'custom':
            d_from = self.date_from
            d_to   = self.date_to
        else:
            d_from = d_to = None

        date_domain = []
        if d_from and d_to:
            date_domain = [
                '|',
                '&', ('invoice_date', '>=', d_from), ('invoice_date', '<=', d_to),
                '&', ('invoice_date', '=', False),
                '&', ('date', '>=', d_from), ('date', '<=', d_to),
            ]
        elif d_from:
            date_domain = [
                '|',
                ('invoice_date', '>=', d_from),
                '&', ('invoice_date', '=', False), ('date', '>=', d_from),
            ]
        elif d_to:
            date_domain = [
                '|',
                ('invoice_date', '<=', d_to),
                '&', ('invoice_date', '=', False), ('date', '<=', d_to),
            ]

        domain = [
            ('state', '!=', 'cancel'),
            ('move_type', 'in', ['out_invoice', 'out_refund', 'in_invoice', 'in_refund']),
        ] + date_domain

        if self.journal_ids:
            domain += [('journal_id', 'in', self.journal_ids.ids)]

        if self.invoice_number_filter:
            domain += [('name', 'ilike', self.invoice_number_filter.strip())]

        _logger.info('FinRecon domain: %s', domain)

        # High-performance: read_group + read instead of ORM browse
        invoices = self.env['account.move'].search(
            domain, order='name asc', limit=10000
        )
        _logger.info('FinRecon found %d invoices', len(invoices))

        # Delete old lines with direct SQL for speed
        if self.line_ids:
            self.env.cr.execute(
                'DELETE FROM fin_reconciliation_line WHERE reconciliation_id = %s',
                (self.id,)
            )
            self.invalidate_recordset()

        # Batch read only needed fields
        invoice_data = invoices.read([
            'name', 'partner_id', 'invoice_date', 'date',
            'amount_total', 'currency_id', 'id', 'state', 'payment_state', 'journal_id'
        ])

        vals_list = [{
            'reconciliation_id': self.id,
            'invoice_number':    r['name'],
            'partner_name':      r['partner_id'][1] if r['partner_id'] else '',
            'invoice_date':      r['invoice_date'] or r['date'],
            'total_amount':      r['amount_total'],
            'currency_id':       r['currency_id'][0] if r['currency_id'] else False,
            'move_id':           r['id'],
            'status':            r['state'],
            'payment_state':     r['payment_state'] or '',
        } for r in invoice_data]

        if vals_list:
            # Batch create for maximum performance
            self.env['fin.reconciliation.line'].create(vals_list)

        self.write({'state': 'loaded'})
        return self._reload()

    def action_import_excel(self):
        """Parse HesabPay Excel and match against loaded lines."""
        self.ensure_one()
        if not self.excel_file:
            raise UserError('Please upload a HesabPay Excel file first.')
        if not openpyxl:
            raise UserError('openpyxl is not installed. Run: pip3 install openpyxl')

        hp_data = self._parse_excel(self.excel_file)

        # Build lookup dict of lines by invoice number for O(1) matching
        lines = self.line_ids
        line_map = {l.invoice_number: l for l in lines}

        matched_vals   = []
        unmatched_ids  = []

        for inv_no, line in line_map.items():
            hp = hp_data.get(inv_no)
            if hp:
                diff = line.total_amount - hp['total']
                matched_vals.append((line.id, {
                    'hp_partner_name':   hp['customer'],
                    'hp_total_amount':   hp['total'],
                    'hp_invoice_date':   hp['date'],
                    'hp_found':          True,
                    'amount_difference': diff,
                    'name_mismatch':     line.partner_name.strip().lower() != hp['customer'].strip().lower(),
                    'amount_mismatch':   abs(diff) > 0.01,
                }))
            else:
                unmatched_ids.append(line.id)

        # Batch write matched lines
        for line_id, vals in matched_vals:
            self.env['fin.reconciliation.line'].browse(line_id).write(vals)

        # Batch reset unmatched lines
        if unmatched_ids:
            self.env['fin.reconciliation.line'].browse(unmatched_ids).write({
                'hp_found':          False,
                'hp_partner_name':   '',
                'hp_total_amount':   0.0,
                'amount_difference': 0.0,
                'name_mismatch':     False,
                'amount_mismatch':   False,
            })

        self.write({'state': 'reconciled'})
        return self._reload()

    def action_export_report(self):
        """Export colour-coded Excel report."""
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

        headers = [
            'Invoice #', 'Customer', 'Date', 'Total',
            'HP Customer', 'HP Date', 'HP Total',
            'Difference', 'Name Match?', 'Amount Match?', 'Result'
        ]
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
            'name':     f'{self.name}_report.xlsx',
            'type':     'binary',
            'datas':    base64.b64encode(buf.read()),
            'mimetype': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        })
        return {
            'type':   'ir.actions.act_url',
            'url':    f'/web/content/{attach.id}?download=true',
            'target': 'self',
        }

    def action_reset(self):
        self.ensure_one()
        self.env.cr.execute(
            'DELETE FROM fin_reconciliation_line WHERE reconciliation_id = %s',
            (self.id,)
        )
        self.invalidate_recordset()
        self.write({'state': 'draft', 'excel_file': False, 'excel_filename': False})
        return self._reload()

    def _reload(self):
        return {
            'type':      'ir.actions.act_window',
            'res_model': 'fin.reconciliation',
            'res_id':    self.id,
            'view_mode': 'form',
            'target':    'current',
        }

    def _parse_excel(self, file_field):
        """Parse Excel: Number | Customer | Total | Invoice Date"""
        file_data = base64.b64decode(file_field)
        wb = openpyxl.load_workbook(io.BytesIO(file_data), data_only=True, read_only=True)
        ws = wb.active

        data = {}
        header_skipped = False
        for row in ws.iter_rows(values_only=True):
            if not header_skipped:
                header_skipped = True
                continue
            if not any(row):
                continue
            try:
                number   = str(row[0]).strip() if row[0] is not None else ''
                customer = str(row[1]).strip() if row[1] is not None else ''
                total    = float(row[2])        if row[2] is not None else 0.0
                date     = row[3]
                if hasattr(date, 'date'):
                    date = date.date()
                elif isinstance(date, str):
                    from datetime import datetime
                    for fmt in ('%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y'):
                        try:
                            date = datetime.strptime(date, fmt).date()
                            break
                        except ValueError:
                            date = None
                if number:
                    data[number] = {'customer': customer, 'total': total, 'date': date}
            except Exception as e:
                _logger.warning('Skipping Excel row: %s', e)
        wb.close()
        return data


class FinReconciliationLine(models.Model):
    _name        = 'fin.reconciliation.line'
    _description = 'Financial Reconciliation Line'
    _order       = 'invoice_number asc'

    # DB index on foreign key for fast joins
    reconciliation_id = fields.Many2one(
        'fin.reconciliation', ondelete='cascade', required=True, index=True
    )

    move_id        = fields.Many2one('account.move', ondelete='set null', index=True)
    invoice_number = fields.Char(string='Invoice #', index=True)
    partner_name   = fields.Char(string='Customer')
    invoice_date   = fields.Date(string='Date', index=True)
    total_amount   = fields.Float(string='Total', digits=(16, 2))
    currency_id    = fields.Many2one('res.currency')
    journal_id     = fields.Many2one('account.journal', string='Journal', index=True)
    status         = fields.Char(string='Status')
    payment_state  = fields.Char(string='Payment')

    hp_found        = fields.Boolean(string='In HesabPay', default=False, index=True)
    hp_partner_name = fields.Char(string='HP Customer')
    hp_invoice_date = fields.Date(string='HP Date')
    hp_total_amount = fields.Float(string='HP Total', digits=(16, 2))

    amount_difference = fields.Float(string='Difference', digits=(16, 2))
    name_mismatch     = fields.Boolean(string='Name Mismatch',   default=False, index=True)
    amount_mismatch   = fields.Boolean(string='Amount Mismatch', default=False, index=True)

    def open_invoice(self):
        """Open the real Odoo invoice form when clicking a result row."""
        self.ensure_one()
        if not self.move_id:
            raise UserError('No linked invoice found for this record.')
        return {
            'type':      'ir.actions.act_window',
            'res_model': 'account.move',
            'res_id':    self.move_id.id,
            'view_mode': 'form',
            'target':    'current',
        }
