from odoo import models, fields
from datetime import datetime, timedelta, date
import dateutil
import xlsxwriter
import io
import base64
import logging
import re

_logger = logging.getLogger(__name__)


class ProductTemplate(models.Model):
    _inherit = 'product.product'

    licence_length_months = fields.Integer(string='Licence Length (Months)')


class LicenseExpiryReport(models.Model):
    _inherit = 'account.move'  # ??

    # LIVE
    # TIME_LIMITS = [30, 60, 90]
    # TEST
    TIME_LIMITS = [3, 6, 9]
    HEADER_TEXT = 'Licence Expiration Report'
    HEADER_VALUES_LIST = ['Note', 'Product Code', 'Product Name', 'Invoice Number', 'Invoice Date',
                          'Licence Length (Months)', 'Expiration date', 'Sale Order', 'Product Variant ID']

    def get_sale_order_name(self, inv_line):
        so_name_list = []
        if not inv_line.sale_line_ids:
            return ''
        for so_line in inv_line.sale_line_ids:
            so_name_list.append(so_line.order_id.name)

        return ', '.join(so_name_list)

    def process_field(self, field_value):
        """ Process and format a field value for report. """
        return field_value or '/'

    def get_and_format_data(self):
        """
        Example of returned data:
        {product.product(8927,): {3: sale.order(32718, 23791), 6: sale.order(), 9: sale.order()},
        product.product(8828,): {3: sale.order(), 6: sale.order(60105,), 9: sale.order(59561, 59086)}}
        """
        today_date = date.today()
        report_data_dict = {}

        # Find all Products where licence_length_months > 0
        all_products = self.env['product.product'].search(
            [('licence_length_months', '>', 0), ('active', 'in', [True, False])])

        if not all_products:
            _logger.warning('WARNING: No products found')
            return {}

        for product in all_products:

            report_data_dict[product] = {}

            for days_until_expiry in self.TIME_LIMITS:

                time_boundary = today_date + timedelta(
                    days=days_until_expiry) - dateutil.relativedelta.relativedelta(
                    months=product.licence_length_months)

                invoices = self.env['account.move'].search([
                    ('invoice_line_ids.product_id', '=', product.id),
                    ('invoice_date', '=', time_boundary),
                    ('state', 'in', ['posted']),
                    ('move_type', '=', 'out_invoice'),
                ])

                if not invoices:
                    report_data_dict[product][days_until_expiry] = []
                    continue

                inv_lines_data_list = []
                for invoice in invoices:
                    inv_lines = [
                        line for line in invoice.invoice_line_ids if line.product_id.id == product.id]

                    if not inv_lines:
                        raise Warning('No INV lines found')
                        continue

                    for inv_line in inv_lines:
                        expiration_date = invoice.invoice_date + dateutil.relativedelta.relativedelta(
                            months=product.licence_length_months)

                        line_data = [
                            f'{days_until_expiry} days until expiration',
                            self.process_field(
                                inv_line.product_id.default_code),
                            self.process_field(inv_line.product_id.name),
                            self.process_field(invoice.name),
                            # f'{time_boundary.strftime('%Y-%m-%d')}',
                            self.process_field(
                                invoice.invoice_date.strftime('%Y%m%d')),
                            self.process_field(product.licence_length_months),
                            self.process_field(
                                expiration_date.strftime('%Y%m%d')),
                            self.process_field(
                                self.get_sale_order_name(inv_line)),
                            self.process_field(inv_line.product_id.id),
                        ]
                        inv_lines_data_list.append(line_data)

                    report_data_dict[product][days_until_expiry] = inv_lines_data_list

        return report_data_dict

    def generate_xlsx_file(self, data_dict):

        # header_values_list = ['Note', 'Product Code', 'Product Name', 'Invoice Number', 'Expiration date', 'Invoice Date', 'Licence Length (Months)']

        # Create a new workbook using XlsxWriter
        buffer = io.BytesIO()
        workbook = xlsxwriter.Workbook(buffer, {'in_memory': True})

        # Defining a bold format for the header
        bold_format = workbook.add_format({'bold': True})
        # Define the top border format
        top_border_format = workbook.add_format({'top': 1})

        worksheet = workbook.add_worksheet()

        # Seting the width of the columns
        # Headers are in the first row of data_matrix and their length determines the column width
        for col_num, header in enumerate(self.HEADER_VALUES_LIST):
            ### TO DO ###
            # column_width = len(header) + 30
            if col_num == 0:
                column_width = len(header) + 20
            elif col_num == 2:
                column_width = len(header) + 30
            else:
                # elif col_num in [1, 3, 4, 5, 6]:
                column_width = len(header)

            # Set the column width
            worksheet.set_column(col_num, col_num, column_width)

            # format_to_use = bold_format if row_num == 0 else None
            worksheet.write(0, col_num, header, bold_format)

        row_num = 1
        for product_dict in list(data_dict.values()):
            new_product_marker = True

            for day_number in self.TIME_LIMITS:
                matrix_of_invoice_lines = product_dict[day_number]

                for line_list in matrix_of_invoice_lines:

                    for col_num, cell_value in enumerate(line_list):

                        format_to_use = top_border_format if new_product_marker else None
                        worksheet.write(row_num, col_num,
                                        cell_value, format_to_use)

                    row_num += 1
                    new_product_marker = False

        # Close the workbook to save changes
        workbook.close()

        # Get the binary data from the BytesIO buffer
        binary_data = buffer.getvalue()
        return base64.b64encode(binary_data)

    def get_email_body(self):
        table_width = 600

        email_content = {
            'text_line_1': 'Hi,',
            'text_line_2': f'Please find attached a {self.HEADER_TEXT}.',
            'text_line_3': 'Kind regards,',
            'text_line_4': 'JTRS Odoo',
            'table_width': table_width
        }

        email_html = f"""
        <!--?xml version="1.0"?-->
        <div style="background:#F0F0F0;color:#515166;padding:10px 0px;font-family:Arial,Helvetica,sans-serif;font-size:12px;">
            <table style="background-color:transparent;width:{email_content['table_width']}px;margin:5px auto;">
                <tbody>
                    <tr>
                        <td style="padding:0px;">
                            <a href="/" style="text-decoration-skip:objects;color:rgb(33, 183, 153);">
                                <img src="/web/binary/company_logo" style="border:0px;vertical-align: baseline; max-width: 100px; width: auto; height: auto;" class="o_we_selected_image" data-original-title="" title="" aria-describedby="tooltip935335">
                            </a>
                        </td>
                        <td style="padding:0px;text-align:right;vertical-align:middle;">&nbsp;</td>
                    </tr>
                </tbody>
            </table>
            <table style="background-color:transparent;width:{email_content['table_width']}px;margin:0px auto;background:white;border:1px solid #e1e1e1;">
                <tbody>
                    <tr>
                        <td style="padding:15px 20px 10px 20px;">
                            <p>{email_content['text_line_1']}</p>
                            </br>
                            <p>{email_content['text_line_2']}</p>
                            </br>
                            <p style="padding-top:20px;">{email_content['text_line_3']}</p>
                            <p>{email_content['text_line_4']}</p>
                        </td>
                    </tr>
                    <tr>
                        <td style="padding:15px 20px 10px 20px;">
                            <!--% include_table %-->
                        </td>
                    </tr>
                </tbody>
            </table>
            <table style="background-color:transparent;width:{email_content['table_width']}px;margin:auto;text-align:center;font-size:12px;">
                <tbody>
                    <tr>
                        <td style="padding-top:10px;color:#afafaf;">
                            <!-- Additional content can go here -->
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>
        """
        return email_html

    def send_email(self, recipient_email, sender_email, cc_email, binary_data):
        # Define email parameters
        subject = f"{self.HEADER_TEXT} ({date.today().strftime('%d/%m/%y')})"
        body = self.get_email_body()

        # Using regular expression to replace '(', ')' and ' ' with '_'.
        attachment_name = re.sub(r'[() /]', '_', f"{subject}.xlsx")
        attachments = [(attachment_name, binary_data)]

        mail_mail = self.env['mail.mail'].create({
            'email_to': recipient_email,
            'email_from': sender_email,
            'email_cc': cc_email,
            'subject': subject,
            'body_html': body,
            'attachment_ids': [(0, 0, {'name': attachment[0], 'datas': attachment[1]}) for attachment in attachments],
        })
        mail_mail.send()
        return True

    def send_license_expiration_report(self, recipient_email, sender_email, cc_email):

        data_dictionary = self.get_and_format_data()

        # raise Warning(f'\n\ndata_dictionary:\n{data_dictionary}\n\n')

        if not data_dictionary:
            _logger.warning('WARNING: data_dictionary was not created.')
            return

        binary_data_report = self.generate_xlsx_file(data_dictionary)

        self.send_email(recipient_email, sender_email,
                        cc_email, binary_data_report)
        return True
