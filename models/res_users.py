# -*- coding: utf-8 -*-
from datetime import timedelta

from odoo import fields, models, api
from datetime import datetime


class ResUsers(models.Model):
    _inherit = 'res.users'

    allowed_journals = fields.Many2many(
        'account.journal',
        'user_journal_rel',  # Relation table
        'user_id',  # Field for users in the relation table
        'journal_id',  # Field for journals in the relation table
        string="Allowed Journals"
    )


