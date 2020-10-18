# -*- coding: utf-8 -*-
from odoo import http

# class ForceDate(http.Controller):
#     @http.route('/force_date/force_date/', auth='public')
#     def index(self, **kw):
#         return "Hello, world"

#     @http.route('/force_date/force_date/objects/', auth='public')
#     def list(self, **kw):
#         return http.request.render('force_date.listing', {
#             'root': '/force_date/force_date',
#             'objects': http.request.env['force_date.force_date'].search([]),
#         })

#     @http.route('/force_date/force_date/objects/<model("force_date.force_date"):obj>/', auth='public')
#     def object(self, obj, **kw):
#         return http.request.render('force_date.object', {
#             'object': obj
#         })