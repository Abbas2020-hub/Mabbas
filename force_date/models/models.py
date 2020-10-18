# -*- coding: utf-8 -*-

from odoo import models, fields, api

class stockinventory(models.Model):
    
    _inherit = "stock.inventory"
    
    
    date = fields.Datetime(
        'Inventory Date', required=True,
        default=fields.Datetime.now,
        help="If the inventory adjustment is not validated, date at which the theoritical quantities have been checked.\n"
             "If the inventory adjustment is validated, date at which the inventory adjustment has been validated.")
