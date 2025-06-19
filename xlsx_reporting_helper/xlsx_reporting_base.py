from odoo import fields, models
import base64, xlsxwriter
from io import BytesIO
from string import ascii_uppercase as char
from odoo.exceptions import ValidationError
from datetime import datetime, date, timedelta
import re
from decimal import Decimal
import pytz


class XlsxReportingBase(models.AbstractModel):
    """
    Abstract base model to generate XLSX reports in Odoo.
    This class provides methods for creating and formatting Excel files.
    """
    _name = "xlsx.reporting.base"
    _description = "XLSX Reporting Base"

    start_date = fields.Datetime("Start Date", required=True, default=lambda x: x._get_start_date())
    end_date = fields.Datetime("End Date", required=True, default=lambda x: x._get_end_date())
    module_name = fields.Char("Module", default=lambda self: self._name, help="Report Module Name")
    report_file = fields.Binary(string="Reporting", required=True)

    def local_time(self, dt, tz):
        """
        convert utc time to tz time str
        dt - datetime object
        tz - timezone string (e.g. 'Asia/Yangon')
        """
        # set naive datetime to utc. this isn't needed in python3.6. but server is python3.5
        datetime_obj = pytz.utc.localize(dt)
        local_tz = pytz.timezone(tz)
        local_datetime = datetime_obj.astimezone(local_tz)
        return local_datetime

    def _get_start_date(self):
        dt = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        return dt - timedelta(hours=6, minutes=30)

    def _get_end_date(self):
        dt = datetime.now().replace(hour=23, minute=59, second=59, microsecond=59)
        return dt - timedelta(hours=6, minutes=30)

    def title_list(self) -> list:
        """
        Generate a formatted title for the report based on the model's name.
         - Remove 'wizard' from the title if present.
         - Replaces dots with space and capitalizes the first letter of each word.
        :return: list of Formatted title string
        """
        title_list = []
        title = self._name
        if 'wizard' in title.lower():
            title = title.replace('wizard', '')
        title_list.append(title.replace('.', ' ').title())
        return title_list

    @staticmethod
    def ss_styling() -> dict:
        """
        Defines the styling for the Excel spreadsheet file.
        :return: Dict styles of title, header, data columns
        """
        return {
            'title': {
                'bold': True, 'align': 'center', 'valign': 'vcenter',
                'font_size': 15, 'font_name': 'Arial', 'underline': True
            },
            'header': {
                'bold': True, 'align': 'center', 'valign': 'vcenter',
                'font_size': 12, 'font_name': 'Arial', 'border': 1
            },
            'data_column': {
                'font_size': 11, 'align': 'center', 'valign': 'vcenter',
                'border': 1, 'font_name': 'Arial', 'num_format': '#,##0.00'
            }
        }

    @staticmethod
    def headers() -> list:
        """
        Placeholder method for defining report headers.
        This method should be overridden by subclasses to provide specific headers.
        :return: List of header strings.
        """
        return []

    def data_keys(self) -> list:
        """
        Generates a list of keys for data entries based on the headers.
        - Uses regular expressions to replace spaces with underscores and () and convert text to lowercase.
        :return: List of formatted keys.
        """
        def replacement(match):
            ch = match.group(0)
            return '' if ch in '}])' else '_'
        return [re.sub('_+', '_',re.sub(r'[\s(){}/\-\[\]]', replacement, key)).lower() for key in self.headers()]

    def _get_data(self) -> list[dict]:
        """
        Generates a list of keys for data entries based on the headers.
        - Uses regular expressions to replace spaces with underscores and convert text to lowercase.
        :return: List of formatted keys.
        """
        return []

    def generate_report_xlsx(self):
        """
        Generates an XLSX report and prepares it for download.
        - Creates a filename based on the start date, end date, and report title.
        - Writes data to the spreadsheet and encodes it for download.
        :return: Dictionary containing the download URL for the report.
        """
        local_start = self.local_time(self.start_date, self.env.user.tz).strftime('%d-%m-%Y')
        local_end = self.local_time(self.end_date, self.env.user.tz).strftime('%d-%m-%Y')
        filename = f"{local_start}_{local_end}_{self._name.replace('.wizard', '')}.xlsx"
        stream = BytesIO()
        wb = xlsxwriter.Workbook(stream, {'in_memory': True})
        ss = wb.add_worksheet()
        num_column = 0 if 'no' in self.data_keys() else None
        if self._get_data():
            self.write_spreadsheet(wb, ss, self.title_list(), self._get_data(), num_column)
            wb.close()
            return self._prepare_download(stream, filename)
        raise ValidationError('There is no data to generate file.')

    def write_spreadsheet(self, wb, ss, titles, data, row_number=None) -> None:
        """
        Writes data to the Excel spreadsheet.
        - Sets column widths and default row height.
        - Adds titles and headers to the spreadsheet.
        - Writes data rows with appropriate formatting.
        :param wb: Workbook object.
        :param ss: Worksheet object.
        :param titles: Titles of the report.
        :param data: List of dictionaries containing data for the report.
        :param row_number: Optional starting row number for numbering rows.
        :return: None
        """
        try:
            ss.set_column(f'{char[0]}:{char[len(self.headers()) - 1]}', 15)
            r = 1
            ss.set_default_row(20)
            for title in titles:
                ss.merge_range(f'{char[0] + str(r)}:{char[len(self.headers()) - 1] + str(r)}', title,
                               wb.add_format(self.ss_styling()['title']))
                r += 1
            if self.headers():
                ss.write_row(f'{char[0] + str(r)}', self.headers(), wb.add_format(self.ss_styling()['header']))
            for row, entry in enumerate(data, r):
                if row_number is not None:
                    row_number += 1
                    entry.update({'no': row_number})
                for col, key in enumerate(self.data_keys()):
                    value = entry.get(key, "")
                    if isinstance(value, Decimal):
                        value = round(float(value), 2)
                    # if isinstance(value, float):
                    #     value = f"{value:,}"
                    elif isinstance(value, dict):
                        value = value.get('en_US', "")
                    elif isinstance(value, date):
                        value = value.strftime("%d-%m-%y")
                    elif isinstance(value, datetime):
                        value = value.strftime("%d-%m-%y %H:%M:%S")
                    ss.write(row, col, value, wb.add_format(self.ss_styling()['data_column']))
        except Exception as e:
            raise ValidationError(e)

    def _prepare_download(self, output, filename):
        """
        Prepares the generated report for download.
        - Encodes the report file in base64.
        - Returns a dictionary with the download URL.
        :param output: BytesIO object containing the report data.
        :param filename: Name of the report file.
        :return: Dictionary containing the download URL.
        """
        output.seek(0)
        self.report_file = base64.b64encode(output.read()).decode('utf-8')
        output.close()
        return {
            'type': 'ir.actions.act_url',
            'name': 'Report Download',
            'target': 'self',
            'url': f'/web/content/{self._name}/{self.id}/report_file/{filename}?download=true',
        }

