from odoo.exceptions import UserError
from odoo import models, fields
from datetime import date, timedelta, datetime
import dateutil.relativedelta
import xlsxwriter
import io
import base64
import logging
import re

_logger = logging.getLogger(__name__)


class LicenseExpirationReport(models.Model):
    _inherit = 'account.move'

    HEADER_TEXT = 'Licence Expiration Report'
    HEADER_VALUES_LIST = [
        'Note', 'Product Code', 'Product Name', 'Invoice Number',
        'Invoice Date', 'Licence Length (Months)', 'Expiration date',
        'Sale Order', 'Delivery Address', 'Salesperson', 'Product Variant ID'
    ]

    def create_scheduled_activity(self, inv_line, info_str, days_until_expiration):
        activity_summary = f'Licence Expiration (#{inv_line.id})'
        sale_order_model_id = 280
        date_deadline = datetime.now() + timedelta(days=days_until_expiration)
        date_deadline_formated = date_deadline.strftime('%d/%m/%Y')
        if not inv_line.sale_line_ids:
            _logger.error(f"No sale_line_ids on inv_line #{inv_line.id}")
            return None
        if not inv_line.sale_line_ids[0].order_id:
            _logger.error(f"No SO associated with inv_line #{inv_line.id}")
            return None

        sale_order = inv_line.sale_line_ids[0].order_id

        existing_activity = self.env['mail.activity'].search([
            ('res_model_id', '=', sale_order_model_id),
            ('res_id', '=', sale_order.id),
            ('date_deadline', '=', date_deadline),
            ('summary', '=', activity_summary)])

        if existing_activity:
            return None

        self.env['mail.activity'].create({
            'date_deadline': date_deadline,
            'res_id': sale_order.id,
            # 'sale.order' ir.model #280
            'res_model_id': 280,
            'user_id': sale_order.user_id.id,
            'note': f'<div style="margin-top: 5px;">Licence expires on  <strong>{date_deadline_formated}</strong>.</div><div style="margin-top: 5px;">Product: {inv_line.product_id.display_name}.</div>',
            'display_name': f'{inv_line.id}',
            'summary': activity_summary,
        })

    def is_integer(self, string):
        return bool(re.match(r"-?\d+$", string))

    def get_config_param(self, key):
        try:
            return self.env['ir.config_parameter'].get_param(key) or ''
        except Exception as e:
            _logger.error(f"Error getting configuration parameter {key}: {e}")
            return ''

    def get_time_checkpoints(self):
        try:
            time_string = self.env['ir.config_parameter'].get_param(
                'licence_expiration_report.time_checkpoints')
            if not time_string:
                return []
            time_string_list = time_string.replace(', ', ',').split(',')

            return [int(time_str) for time_str in time_string_list if self.is_integer(time_str)]
        except Exception as e:
            _logger.error(f"Error in retrieving time checkpoints: {e}")
            return []

    def log_message(self, message, function_name):
        self.env['ir.logging'].create({
            'name': 'Licence Expiration Report',
            'type': 'server',
            'dbname': self.env.cr.dbname,
            'level': 'info',
            'message': message,
            'path': 'models.account.move',
            'line': f'LicenseExpirationReport.{function_name}',
            'func': f'__{function_name}__',
        })

    def get_sale_order_name(self, inv_line):
        """
        Extracts and concatenates sale order names from invoice lines.
        """
        so_name_list = [
            so_line.order_id.name for so_line in inv_line.sale_line_ids]
        return ', '.join(so_name_list) if so_name_list else ''

    def get_sale_order_obj(self, inv_line):
        """
        Extracts the first sale order object from the invoice line.
        """
        try:
            # Attempt to return the first sale order object if it exists
            return inv_line.sale_line_ids and inv_line.sale_line_ids[0].order_id or None
        except Exception as e:
            _logger.error('Failed to extract sale order object: %s', e)
            return None

    def process_field(self, field_value):
        """
          Formats a field value for report.
        """
        return field_value or '/'

    def get_note_text(self, days_until_expiration):
        if days_until_expiration == None:
            return '/'
        if days_until_expiration > 0:
            return f'{days_until_expiration} days until expiration'
        elif days_until_expiration < 0:
            days_until_expiration *= -1
            return f'Expired {days_until_expiration} days ago'
        else:
            return 'Expires today'

    def get_salesperson_from_so_partner(self, inv_line):
        sale_order = self.get_sale_order_obj(inv_line)
        if not sale_order:
            return None
        if sale_order.partner_id.user_id:
            return sale_order.partner_id.user_id.name
        else:
            return None

    def process_invoice_line(self, inv_line, invoice, product, days_until_expiry):
        if invoice.invoice_date and product.x_licence_length_months:
            expiration_date = invoice.invoice_date + \
                dateutil.relativedelta.relativedelta(
                    months=product.x_licence_length_months)
        else:
            expiration_date = None

        return [
            self.get_note_text(days_until_expiry),
            self.process_field(inv_line.product_id.default_code),
            self.process_field(inv_line.product_id.name),
            self.process_field(invoice.name),
            self.process_field(invoice.invoice_date.strftime(
                '%Y-%m-%d')) if invoice.invoice_date else '/',
            self.process_field(product.x_licence_length_months),
            self.process_field(expiration_date.strftime(
                '%Y-%m-%d')) if expiration_date else '/',
            self.process_field(self.get_sale_order_name(inv_line)),
            self.process_field(invoice.partner_shipping_id.display_name),
            self.process_field(self.get_salesperson_from_so_partner(inv_line)),
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

    def switch_on_so_line_is_on(self, inv_line):
        # Check if the invoice line has associated sale line ids
        if not inv_line.sale_line_ids:
            _logger.warning('WARNING: No inv_line.sale_line_ids')
            return False

        # Return True if any of the sale lines are marked to omit from license expiration report
        # This checks each sale line's custom field 'x_omit_from_licence_expiration_report'
        return any(so_line.x_omit_from_licence_expiration_report for so_line in inv_line.sale_line_ids)

    def get_and_format_data(self):
        """
        Generates a dictionary with report data structured as:
        {product: {days_until_expiry: [[invoice_line_data], ...]}, ...}
        """
        try:
            today_date = date.today()
            report_data_dict = {}

            # Searching for products with a defined license length
            all_products = self.env['product.product'].search([
                ('x_licence_length_months', '>', 0),
                ('active', 'in', [True, False])
            ])

            if not all_products:
                _logger.warning('WARNING: No products found')
                return {}

            # Looping through each product to populate report data
            for product in all_products:
                time_checkpoints = self.get_time_checkpoints()

                report_data_dict[product] = {days: []
                                             for days in time_checkpoints}
                for days_until_expiry in time_checkpoints:
                    time_boundary = today_date + timedelta(days=days_until_expiry) - \
                        dateutil.relativedelta.relativedelta(
                            months=product.x_licence_length_months)

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
                            if self.switch_on_so_line_is_on(inv_line):
                                _logger.warning(
                                    f'WARNING: SO Line #{inv_line.id} is omitted because the switch button is on.')
                                continue
                            line_data = self.process_invoice_line(
                                inv_line, invoice, product, days_until_expiry)

                            inv_lines_data_list.append(line_data)

                            self.create_scheduled_activity(
                                inv_line, line_data[0], days_until_expiry)

                        report_data_dict[product][days_until_expiry] = inv_lines_data_list

            if not self.check_if_any_data_found(report_data_dict):
                self.log_message('No data found', 'get_and_format_data')
                return {}

            # raise UserError(f'report_data_dict:\n{report_data_dict}')
            return report_data_dict
        except Exception as e:
            _logger.error(f"Error in formatting data: {e}")
            return {}

    def apply_cell_formating(self, col_num, day_number, new_product_marker):
        format_dict = {}
        if day_number < 0 and col_num == 0:
            # format_dict['bg_color'] = 'red'
            format_dict['bg_color'] = '#c47772'
        if day_number >= 0 and day_number < 30 and col_num == 0:
            # format_dict['bg_color'] = 'green'
            # format_dict['bg_color'] = '#5d917e'
            format_dict['bg_color'] = '#FFBF00'
        if new_product_marker:
            format_dict['top'] = 1
        return format_dict

    def generate_xlsx_file(self, data_dict):
        try:
            # Create a new workbook using XlsxWriter
            buffer = io.BytesIO()
            workbook = xlsxwriter.Workbook(buffer, {'in_memory': True})

            # Defining a bold format for the header
            bold_format = workbook.add_format({'bold': True})

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

                worksheet.write(0, col_num, header, bold_format)

            row_num = 1
            # raise UserError(f'list(data_dict.values()):\n{list(data_dict.values())}')

            for product_dict in list(data_dict.values()):
                new_product_marker = True

                for day_number in self.get_time_checkpoints():
                    matrix_of_invoice_lines = product_dict[day_number]

                    # Example of line_list: ['Expires today', 'LBX-V3-ONB20UKB', 'LoxBox 20 Cart - (Black/Orange) for Tablets & Notebooks - UK Power', 'INV/2023/0736', '2023-02-23', 12, '2024-02-23', 'SO55291', '### School, Julie Chandler', 'M### M###', 13969]
                    for line_list in matrix_of_invoice_lines:

                        for col_num, cell_value in enumerate(line_list):

                            format_to_use = workbook.add_format(
                                self.apply_cell_formating(col_num, day_number, new_product_marker))

                            worksheet.write(row_num, col_num,
                                            cell_value, format_to_use)

                        row_num += 1
                        new_product_marker = False

            # Close the workbook to save changes
            workbook.close()

            # Get the binary data from the BytesIO buffer
            binary_data = buffer.getvalue()
            return base64.b64encode(binary_data)

        except Exception as e:
            _logger.error(f"Error in generating XLSX file: {e}")
            return None

    def send_email_with_attachment(self, subject, body, attachment):
        try:
            mail_mail = self.env['mail.mail'].create({
                'email_to': self.get_config_param('licence_expiration_report.recipient_email'),
                'email_from': self.get_config_param('licence_expiration_report.sender_email'),
                'email_cc': self.get_config_param('licence_expiration_report.cc_email'),
                'reply_to': self.get_config_param('licence_expiration_report.reply_to_email'),
                'subject': subject,
                'body_html': body,
                'attachment_ids': [(0, 0, {'name': attachment[0], 'datas': attachment[1]})],
            })
            mail_mail.send()
            self.log_message('Email sent', 'send_email_with_attachment')

        except Exception as e:
            _logger.error(f"Error in sending email: {e}")

    def prepare_email_content(self):
        return {
            'text_line_1': 'Hi,',
            'text_line_2': f'Please find attached the {self.HEADER_TEXT}.',
            'text_line_3': 'Kind regards,',
            'text_line_4': self.get_config_param('licence_expiration_report.email_company_name'),
            'table_width': 600
        }

    def generate_email_html(self, email_content):
        return f"""
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
                        </td>
                    </tr>
                </tbody>
            </table>
            <table style="background-color:transparent;width:{email_content['table_width']}px;margin:auto;text-align:center;font-size:12px;">
                <tbody>
                    <tr>
                        <td style="padding-top:10px;color:#afafaf;">
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>
        """
        return email_html

    def create_email_attachment(self, binary_data, subject):
        attachment_name = re.sub(r'[() /]', '_', f"{subject}.xlsx")
        return (attachment_name, binary_data)

    def send_licence_expiration_report(self):
        data_dict = self.get_and_format_data()

        if not data_dict:
            _logger.warning('No data to report.')
            return

        binary_data = self.generate_xlsx_file(data_dict)
        subject = f"{self.HEADER_TEXT} ({date.today().strftime('%d/%m/%y')})"
        email_body = self.generate_email_html(self.prepare_email_content())
        attachment = self.create_email_attachment(binary_data, subject)

        self.send_email_with_attachment(subject, email_body, attachment)
