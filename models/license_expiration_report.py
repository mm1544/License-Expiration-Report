from odoo.exceptions import UserError
from odoo import models, fields
from datetime import date, timedelta
import dateutil.relativedelta
import xlsxwriter
import io
import base64
import logging
import re

_logger = logging.getLogger(__name__)


class ProductTemplate(models.Model):
    _inherit = 'product.product'

    # Field to store license length in months
    licence_length_months = fields.Integer(string='Licence Length (Months)')


class LicenseExpirationReport(models.Model):
    _inherit = 'account.move'

    TIME_LIMITS = [14, 30, 60, 90]
    HEADER_TEXT = 'Licence Expiration Report'
    HEADER_VALUES_LIST = [
        'Note', 'Product Code', 'Product Name', 'Invoice Number',
        'Invoice Date', 'Licence Length (Months)', 'Expiration date',
        'Sale Order', 'Customer', 'Salesperson', 'Product Variant ID'
    ]

    def get_sale_order_name(self, inv_line):
        """
        Extracts and concatenates sale order names from invoice lines.
        """
        so_name_list = [
            so_line.order_id.name for so_line in inv_line.sale_line_ids]
        return ', '.join(so_name_list) if so_name_list else ''

    def process_field(self, field_value):
        """
          Formats a field value for report.
        """
        return field_value or '/'

    def process_invoice_line(self, inv_line, invoice, product, days_until_expiry):
        if invoice.invoice_date and product.licence_length_months:
            expiration_date = invoice.invoice_date + \
                dateutil.relativedelta.relativedelta(
                    months=product.licence_length_months)
        else:
            expiration_date = None

        return [
            f'{days_until_expiry} days until expiration' if days_until_expiry else '/',
            self.process_field(inv_line.product_id.default_code),
            self.process_field(inv_line.product_id.name),
            self.process_field(invoice.name),
            self.process_field(invoice.invoice_date.strftime(
                '%Y-%m-%d')) if invoice.invoice_date else '/',
            self.process_field(product.licence_length_months),
            self.process_field(expiration_date.strftime(
                '%Y-%m-%d')) if expiration_date else '/',
            self.process_field(self.get_sale_order_name(inv_line)),
            self.process_field(invoice.partner_id.display_name),
            self.process_field(invoice.invoice_user_id.name),
            self.process_field(inv_line.product_id.id),
        ]

    def check_if_any_data_found(self, data_dict):
        """
        Checks if passed in dictionary contains data in nested lists. If no data found, returns False.

        Args:
            data_dict (dictionary)
            Example:
            {product.product(12846,): {30: [], 60: [], 90: []}, product.product(15643,): {30: [], 60: [], 90: []}, product.product(14851,): {30: [], 60: [], 90: []}}
        Returns: True/False
        """
        for inner_dict in data_dict.values():
            for value_list in inner_dict.values():
                if value_list:  # Checks if the list is non-empty
                    return True
        return False

    def get_and_format_data(self):
        """
        Generates a dictionary with report data structured as:
        {product: {days_until_expiry: [[invoice_line_data], ...]}, ...}
        """
        today_date = date.today()
        report_data_dict = {}

        # Searching for products with a defined license length
        all_products = self.env['product.product'].search([
            ('licence_length_months', '>', 0),
            ('active', 'in', [True, False])
        ])

        if not all_products:
            _logger.warning('WARNING: No products found')
            return {}

        # Looping through each product to populate report data
        for product in all_products:
            # OLD
            # report_data_dict[product] = {}
            # NEW / TEST
            report_data_dict[product] = {days: [] for days in self.TIME_LIMITS}
            for days_until_expiry in self.TIME_LIMITS:
                time_boundary = today_date + timedelta(days=days_until_expiry) - \
                    dateutil.relativedelta.relativedelta(
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
                        _logger.warning('WARNING: No INV lines found')
                        continue

                    for inv_line in inv_lines:
                        line_data = self.process_invoice_line(
                            inv_line, invoice, product, days_until_expiry)

                        inv_lines_data_list.append(line_data)

                    report_data_dict[product][days_until_expiry] = inv_lines_data_list

        if not self.check_if_any_data_found(report_data_dict):
            self.log('No data found', 'get_and_format_data')
            return {}

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
            if col_num in [0]:
                column_width = len(header) + 20
            elif col_num in [9]:
                column_width = len(header) + 10
            elif col_num in [2, 8]:
                column_width = len(header) + 30
            else:
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
        self.log('Email was sent', 'send_email')
        return True

    def send_license_expiration_report(self):
        # Temporary
        recipient_email = 'laura.stockton@jtrs.co.uk'
        # recipient_email = 'martynas.minskis@jtrs.co.uk'
        sender_email = 'OdooBot <odoobot@jtrs.co.uk>'
        cc_email = 'martynas.minskis@jtrs.co.uk'

        data_dictionary = self.get_and_format_data()
        # raise UserError('bp1')

        if not data_dictionary:
            _logger.warning('WARNING: data_dictionary was not created.')
            return

        binary_data_report = self.generate_xlsx_file(data_dictionary)

        self.send_email(recipient_email, sender_email,
                        cc_email, binary_data_report)
        return True

    def log(self, message, function_name):

        # Create a log entry in ir.logging
        self.env['ir.logging'].create({
            'name': 'Licence Expiration Report',  # Name of the log
            'type': 'server',  # Indicates that this log is from the server-side
            'dbname': self.env.cr.dbname,  # Current database name
            'level': 'info',  # Log level (info, warning, error)
            'message': message,  # The main log message
            'path': 'models.account.move',  # Path indicates the module/class path
            # Method name or line number
            'line': 'LicenseExpirationReport.log',
            'func': f'__{function_name}__',  # Function name
        })
