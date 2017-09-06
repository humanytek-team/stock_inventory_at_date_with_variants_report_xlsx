# -*- coding: utf-8 -*-
###############################################################################
#
#    Odoo, Open Source Management Solution
#    Copyright (C) 2017 Humanytek (<www.humanytek.com>).
#    Manuel Márquez <manuel@humanytek.com>
#
#    This program is free software: you can redistribute it and/or modify
#    it under the terms of the GNU Affero General Public License as
#    published by the Free Software Foundation, either version 3 of the
#    License, or (at your option) any later version.
#
#    This program is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU Affero General Public License for more details.
#
#    You should have received a copy of the GNU Affero General Public License
#    along with this program.  If not, see <http://www.gnu.org/licenses/>.
#
###############################################################################

from openerp.addons.report_xlsx.report.report_xlsx import ReportXlsx
from openerp.tools.translate import _


class InventoryAtDateReportXlsx(ReportXlsx):

    def generate_xlsx_report(self, workbook, data, stock_history):
        import logging
        _logger = logging.getLogger(__name__)        

        report_name = _('Inventory at Date')
        sheet = workbook.add_worksheet(report_name)
        bold = workbook.add_format({'bold': True})

        # Header
        col = 0
        header_prefix = [_('Company'), _('Product'), _('Category')]
        header_sufix = [_('Quantity Total'), _('Price'), _('Inventory Value')]
        for col_title in header_prefix:
            sheet.write(0, col, col_title, bold)
            col += 1

        attrs = ['talla']
        ProductAttribute = self.env['product.attribute']

        attrs_records = list()
        for attr in attrs:
            attr_data = ProductAttribute.search([
                ('name', '=', attr),
            ])
            if attr_data:
                attrs_records.append(attr_data[0])
                for value in attr_data[0].value_ids:
                    col_title = '{0} ({1})'.format(
                        attr_data[0].name, value.name)
                    sheet.write(0, col, col_title, bold)
                    col += 1

        for col_title in header_sufix:
            sheet.write(0, col, col_title, bold)
            col += 1

        data = list()
        product_ids = list()

        for line in stock_history:

            if line.product_template_id.id not in product_ids:
                data_product = dict()
                data_product[str(line.product_template_id.id)] = list()

                data_by_company = {
                    'company_id': line.company_id.id,
                    'company': line.company_id.name,
                    'product': line.product_template_id.name,
                    'category': line.product_template_id.categ_id.name,
                    'qty': line.quantity,
                    'price': line.product_template_id.list_price,
                    'inventory_value': line.inventory_value,
                }

                for attr in attrs_records:
                    for value in attr.value_ids:
                        dict_key = '{0}_{1}'.format(attr.id, value.id)
                        if value.id in \
                            line.product_id.attribute_value_ids.mapped('id'):
                            data_by_company.update({dict_key: line.quantity})
                        else:
                            data_by_company.update({dict_key: 0})

                data_product[str(line.product_template_id.id)].append(
                    data_by_company)
                data.append(data_product)

            else:
                product_id = str(line.product_template_id.id)
                data_product = False
                data_product = (item for item in data
                    if product_id in item).next()

                if data_product:
                    companies_ids = [
                        data_by_company['company_id']
                        for data_by_company in data_product[product_id]]

                    if line.company_id.id in companies_ids:
                        for data_by_company in data_product[product_id]:
                            if data_by_company['company_id'] == \
                                line.company_id.id:

                                data_by_company['qty'] += line.quantity
                                data_by_company['inventory_value'] += \
                                    line.inventory_value

                                for attr in attrs_records:
                                    for value in attr.value_ids:
                                        if value.id in \
                                            line.product_id.attribute_value_ids.mapped('id'):
                                            dict_key = '{0}_{1}'.format(attr.id, value.id)
                                            data_by_company[dict_key] += line.quantity
                    else:
                        data_by_company = {
                            'company_id': line.company_id.id,
                            'company': line.company_id.name,
                            'product': line.product_template_id.name,
                            'category': line.product_template_id.categ_id.name,
                            'qty': line.quantity,
                            'price': line.product_template_id.list_price,
                            'inventory_value': line.inventory_value,
                        }

                        for attr in attrs_records:
                            for value in attr.value_ids:
                                dict_key = '{0}_{1}'.format(attr.id, value.id)
                                if value.id in \
                                    line.product_id.attribute_value_ids.mapped('id'):
                                    data_by_company.update({dict_key: line.quantity})
                                else:
                                    data_by_company.update({dict_key: 0})

                        data_product[product_id].append(data_by_company)

            product_ids.append(line.product_template_id.id)

        row = 1
        cols_lines_prefix = ['company', 'product', 'category']
        cols_lines_sufix = ['qty', 'price', 'inventory_value']
        for data_product in data:
            col = 0
            data_by_company = data_product[data_product.keys()[0]]

            for data_stock_company in data_by_company:

                for col_value in cols_lines_prefix:
                    sheet.write(row, col, data_stock_company[col_value])
                    col += 1

                for attr in attrs_records:
                    for value in attr.value_ids:
                        dict_key = '{0}_{1}'.format(attr.id, value.id)
                        sheet.write(row, col, data_stock_company[dict_key])
                        col += 1

                for col_value in cols_lines_sufix:
                    sheet.write(row, col, data_stock_company[col_value])
                    col += 1

                row += 1

InventoryAtDateReportXlsx(
    'report.inventory.at.date.with.variants.report.xlsx', 'stock.history')
