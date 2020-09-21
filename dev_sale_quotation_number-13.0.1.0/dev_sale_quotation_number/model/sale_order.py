# -*- coding: utf-8 -*-
##############################################################################
#
#    OpenERP, Open Source Management Solution
#    Copyright (C) 2015 DevIntelle Consulting Service Pvt.Ltd (<http://www.devintellecs.com>).
#
#    For Module Support : devintelle@gmail.com  or Skype : devintelle 
#
##############################################################################
#
from odoo import models, api, fields, _

class sale_order(models.Model):
    _inherit = "sale.order"
    
    @api.model
    def _get_quotation_sequence(self):
        seq_id = self.env['ir.sequence'].search([('code','=','sale.order')],limit=1)
        if seq_id:
            return seq_id.id
        else:
            return False

    sale_number = fields.Many2one('ir.sequence', string='Sequence',required=True,
                                  domain=[('code', 'in', ['sale.order', 'sale.quotation'])],
                                  default=_get_quotation_sequence)


    @api.model
    def create(self, vals):
        if vals.get('name', _('New')) == _('New'):
            if vals.get('sale_number'):
                obj_seq = self.env['ir.sequence'].browse(vals.get('sale_number'))
                if obj_seq.code == 'sale.order':
                    vals['name'] = self.env['ir.sequence'].next_by_code(
                        'sale.order') or '/'
                else:
                    vals['name'] = self.env['ir.sequence'].next_by_code(
                        'sale.quotation') or '/'
        return super(sale_order, self).create(vals)


    def _action_confirm(self):
        super(sale_order, self)._action_confirm()
        sale_order_sequence_id = self.env['ir.sequence']\
            .search([('code', '=', 'sale.order')])
        if sale_order_sequence_id:
            if self.sale_number.id != sale_order_sequence_id.id:
                self.sale_number = sale_order_sequence_id.id
                self.name = self.env['ir.sequence'].next_by_code('sale.order')
    
    
        

# vim:expandtab:smartindent:tabstop=4:softtabstop=4:shiftwidth=4:
