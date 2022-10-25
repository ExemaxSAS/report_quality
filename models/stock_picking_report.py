from odoo import models
import re
from datetime import timedelta, date

class StockPicking(models.AbstractModel):
    _name = 'report.report_quality.stock_picking_quality'
    _inherit = 'report.report_xlsx.abstract'

    def generate_xlsx_report(self, workbook, data, pickings):
        report_name = 'Reporte Calidad'
        sheet = workbook.add_worksheet(report_name[:31])
        h = "#"
        money_format = workbook.add_format({'num_format': "$ 0" + h + h + '.' + h + h + ',' + h + h})
        date_style = workbook.add_format({'text_wrap': True, 'num_format': 'dd-mm-yyyy'})
        bold = workbook.add_format({'bold': True})
        titles = ['Fecha de Recepción', 'Responsable de Recepción', 'Proveedor', 'Producto', 'Marca', 'Nº de Remito', 'Cantidad', 'Fecha de Vencimiento', 'Lote']

        row = 1
        index = 0

        for i,title in enumerate(titles):
            sheet.write(0, i, title, bold)

        for rec in pickings:
            for product in rec.move_ids_without_package:
               date_exp = rec.date_done + timedelta(days=product.product_id.expiration_time)

               for i,lot in enumerate(product.move_line_nosuggest_ids):
                   sheet.write(row + index, 0, rec.date_done, date_style) # Fecha de Recepcion
                   if rec.user_id:
                       sheet.write(row + index, 1, rec.user_id.name) # Responsable
                   sheet.write(row + index, 2, rec.partner_id.name) # Proveedor
                   sheet.write(row + index, 3, product.product_id.name) # Producto
                   if product.product_id.product_brand_id:
                       sheet.write(row + index, 4, product.product_id.product_brand_id.name) # Marca
                   if rec.num_remi:
                      sheet.write(row + index, 5, rec.num_remi) # N° Remito
                   sheet.write(row + index, 6, lot.qty_done) # Cantidad
                   sheet.write(row + index, 7, date_exp, date_style) # Fecha de Vencimiento
                   sheet.write(row + index, 8, lot.lot_name) # Lote

                   index = index + 1

