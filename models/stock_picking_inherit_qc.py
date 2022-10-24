
from odoo import models, api, fields

class StockPicking(models.Model):
    _inherit = "stock.picking"

    def print_report_xml(self):
        #redirect to /account/account_accpunt_extra_sales_report to generate the excel file 
        return {
            'type':'ir.actions.act_url',
            'url':'/account/account_extra_sales_report/%s' %(self.id),
            'target':'new',
        }
