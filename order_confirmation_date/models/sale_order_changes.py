# -*- coding: utf-8 -*-
from odoo import models, fields, api, _
from odoo.exceptions import Warning
# Ahmed Salama Code Start ---->


class SaleOrderInherit(models.Model):
	_inherit = 'sale.order'
	
	date_order = fields.Datetime(readonly=False)

# Ahmed Salama Code End.