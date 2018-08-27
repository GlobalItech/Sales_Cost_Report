from openerp import models, fields, api
from reportlab.graphics.shapes import String


class CostReport(models.TransientModel):
    _name = "wizard.scr"
    _description = "Sales Cost Report"

    date_to= fields.Date("Date To")
    date_from= fields.Date("Date From")
    product_categ= fields.Many2many('product.category', string="Product Category")
    company= fields.Many2one('res.company', string="Company")

    @api.multi
    def export_xls(self):
        context = self._context
        datas = {'ids': context.get('active_ids', [])}
        datas['model'] = 'account.invoice'
        datas['form'] = self.read()[0]
        for field in datas['form'].keys():
            if isinstance(datas['form'][field], tuple):
                datas['form'][field] = datas['form'][field][0]
        if context.get('xls_export'):
            return {'type': 'ir.actions.report.xml',
                    'report_name': 'Sales Cost Report.scr_report_xls.xlsx',
                    'datas': datas,
                    'name': 'SCR'
                    }
          