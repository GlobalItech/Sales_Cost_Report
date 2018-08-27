from odoo.addons.report_xlsx.report.report_xlsx import ReportXlsx
from odoo import fields, models,api
from datetime import date, datetime

class CsrReportXls(ReportXlsx):
    @api.multi
    def get_lines(self,date_from,date_to,product_categ,company):
        
        lines = []

        sale_obj_ids = self.env['account.invoice'].search([('date_invoice', '>=',date_from),
                                                                ('date_invoice', '<=',date_to),('state','!=',['draft', 'cancel'])])       
        if sale_obj_ids:
            
            for sale_obj in sale_obj_ids:
                flag = False
                if sale_obj.origin:
                    if 'SO' in sale_obj.origin and sale_obj.company_id.id == company:
                        flag = True
                        a = 1
                    if 'INV' in sale_obj.origin and sale_obj.company_id.id == company:
                        flag = True
                        a = -1
                    
                if flag == True:
                    for order in sale_obj.invoice_line_ids:
                        if product_categ:
                            for pr in product_categ:
                                if order.product_id.categ_id.id == pr:
                            
                                    vals = {
                                        'code': order.product_id.default_code or ' ',
                                        'customer': sale_obj.partner_id.name or ' ',
                                        'date': sale_obj.date_invoice,
                                        'name': order.product_id.name + ' ' + str(order.product_id.attribute_value_ids.name or ' '),
                                        'qty': a*order.quantity,
                                        'cost': order.product_id.standard_price,
                                        'tcost': order.product_id.standard_price * order.quantity,
                                        'sale': a*(order.price_unit * order.quantity),
                                        'profit': a*((order.quantity * order.price_unit) - (order.quantity* order.product_id.standard_price)),
                    
                                    }
                                    lines.append(vals)
                                    
                                
                        else:
                            vals = {
                                'code': order.product_id.default_code or ' ',
                                'customer': sale_obj.partner_id.name or ' ',
                                'date': sale_obj.date_invoice,
                                'name': order.product_id.name + ' ' + str(order.product_id.attribute_value_ids.name or ' '),
                                'qty': a*order.quantity,
                                'cost': order.product_id.standard_price,
                                'tcost': order.product_id.standard_price * order.quantity,
                                'sale': a*(order.price_unit * order.quantity),
                                'profit': a*((order.quantity * order.price_unit) - (order.quantity* order.product_id.standard_price)),
                            }
                            lines.append(vals)
            return lines

    def generate_xlsx_report(self, workbook, data, lines):
        sheet = workbook.add_worksheet()
       
        
        format1 = workbook.add_format({'font_size': 14, 'bottom': True, 'right': True, 'left': True, 'top': True, 'align': 'center', 'bold': True})
        format11 = workbook.add_format({'font_size': 14, 'align': 'center', 'bold': True,})
#         format123 = workbook.set_column('A:A', 100)
        period_format= workbook.add_format({'font_size': 11, 'align': 'center', 'bold': True})

        format12 = workbook.add_format({'font_size': 11, 'align': 'center', 'bold': True,'right': True, 'left': True,'bottom': True, 'top': True})
        format21 = workbook.add_format({'font_size': 10, 'align': 'right', 'right': True, 'left': True,'bottom': True, 'top': True})
        format21.set_num_format('#,##0.00')
        qty_format = workbook.add_format({'font_size': 10, 'align': 'right', 'right': True, 'left': True,'bottom': True, 'top': True})
        qty_format.set_num_format('#,##0')
        Pname_format = workbook.add_format({'font_size': 10, 'align': 'left', 'right': True, 'left': True,'bottom': True, 'top': True})
        format_center = workbook.add_format({'font_size': 10, 'align': 'center', 'right': True, 'left': True,'bottom': True, 'top': True})
        format3 = workbook.add_format({'bottom': True, 'top': True, 'font_size': 12})
        font_size_8 = workbook.add_format({'bottom': True, 'top': True, 'right': True, 'left': True, 'font_size': 8})
        red_mark = workbook.add_format({'bottom': True, 'top': True, 'right': True, 'left': True, 'font_size': 8,
                                        'bg_color': 'red'})
        justify = workbook.add_format({'bottom': True, 'top': True, 'right': True, 'left': True, 'font_size': 12})
#         style = workbook.add_format('align: wrap yes; borders: top thin, bottom thin, left thin, right thin;')
#         style.num_format_str = '#,##0.00'
        format3.set_align('center')
        font_size_8.set_align('center')
        justify.set_align('justify')
        format1.set_align('center') 
        red_mark.set_align('center')
        
        date_from = datetime.strptime(data['form']['date_from'], '%Y-%m-%d').strftime('%d/%m/%y')
        date_to = datetime.strptime(data['form']['date_to'], '%Y-%m-%d').strftime('%d/%m/%y')
        sheet.merge_range(0, 0, 0, 8, 'Sales Cost Report', format11)
        sheet.merge_range(1, 0, 1, 8, 'Period from: ' + (date_from) +  ' to ' + (date_to), period_format)


        sheet.write(3, 0,'DATE', format12)
        sheet.write(3, 1,'CUSTOMER', format12)
        sheet.write(3, 2,'CODE', format12)
        sheet.write(3, 3, 'PRODUCT NAME', format12)
        sheet.write(3, 4, 'PRODUCT QUANTITY', format12)
        sheet.write(3, 5, 'COST', format12)
        sheet.write(3, 6, 'TOTAL COST', format12)
        sheet.write(3, 7, 'SALE VALUE', format12)
        sheet.write(3, 8, 'NET/ PROFIT', format12)
         
        # report start
        product_row = 4
        get_lines = self.get_lines(data['form']['date_from'],data['form']['date_to'],data['form']['product_categ'],data['form']['company'])
       
        
        sr_no =1
        for line in  get_lines:
           
            date =datetime.strptime(line['date'],'%Y-%m-%d').strftime('%d-%m-%y')

            sheet.write(product_row, 0, date, format_center)
            sheet.write(product_row, 1, line['customer'], Pname_format)
            sheet.write(product_row, 2, line['code'], format_center)
            sheet.write(product_row, 3, line['name'], Pname_format)
            sheet.write(product_row, 4, line['qty'], qty_format)
            sheet.write(product_row, 5, line['cost'], format_center)
            sheet.write(product_row, 6, line['tcost'], format21)
            sheet.write(product_row, 7, line['sale'], format21)
            sheet.write(product_row, 8, line['profit'], format21)
              
            sr_no +=1   
            product_row +=1

            
CsrReportXls('report.Sales Cost Report.scr_report_xls.xlsx','account.invoice')
