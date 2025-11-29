from odoo.models import AbstractModel
import io, datetime
from odoo.tools.misc import xlsxwriter, file_path
from collections import defaultdict
from PIL import ImageFont

XLSX_GRAY_200 = '#EEEEEE'
XLSX_BORDER_COLOR = '#B4B4B4'
XLSX_FONT_SIZE_DEFAULT = 8
XLSX_FONT_SIZE_HEADING = 11

class AccountReport(AbstractModel):
    _inherit = "account.journal.report.handler"

    def export_to_xlsx(self, options, response=None):
        # self.ensure_one()
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {
            'in_memory': True,
            'strings_to_formulas': False,
        })
        report = self.env['account.report'].search([('id', '=', options['report_id'])], limit=1)
        print_options = report.get_options(previous_options={**options, 'export_mode': 'print'})
        document_data = self._generate_document_data_for_export(report, print_options, 'xlsx')

        # We need to use fonts to calculate column width otherwise column width would be ugly
        # Using Lato as reference font is a hack and is not recommended. Customer computers don't have this font by default and so
        # the generated xlsx wouldn't have this font. Since it is not by default, we preferred using Arial font as default and keep
        # Lato as reference for columns width calculations.
        fonts = {}
        for font_size in (XLSX_FONT_SIZE_HEADING, XLSX_FONT_SIZE_DEFAULT):
            fonts[font_size] = defaultdict()
            for font_type in ('Reg', 'Bol', 'RegIta', 'BolIta'):
                try:
                    lato_path = f'web/static/fonts/lato/Lato-{font_type}-webfont.ttf'
                    fonts[font_size][font_type] = ImageFont.truetype(file_path(lato_path), font_size)
                except (OSError, FileNotFoundError):
                    # This won't give great result, but it will work.
                    fonts[font_size][font_type] = ImageFont.load_default()

        for journal_vals in document_data['journals_vals']:
            cursor_x = 0
            cursor_y = 0

            # Default sheet properties
            sheet = workbook.add_worksheet(journal_vals['name'][:31])
            columns = journal_vals['columns']

            for column in columns:
                align = 'left'
                if 'o_right_alignment' in column.get('class', ''):
                    align = 'right'
                self._write_cell(cursor_x, cursor_y, column['name'], 1, False, report, fonts, workbook, sheet, XLSX_FONT_SIZE_HEADING,
                                 True, XLSX_GRAY_200, align, 2, 2)
                cursor_x = cursor_x + 1

            # Set cursor coordinates for the table generation
            cursor_y += 1
            cursor_x = 0
            for line in journal_vals['lines'][:-1]:
                is_first_aml_line = False
                for column in columns:
                    border_top = 0 if not is_first_aml_line else 1
                    align = 'left'

                    if line.get(column['label'], {}).get('data'):
                        data = line[column['label']]['data']
                        is_date = isinstance(data, datetime.date)
                        bold = False

                        if 'o_right_alignment' in column.get('class', ''):
                            align = 'right'

                        if line[column['label']].get('class') and 'o_bold' in line[column['label']]['class']:
                            # if the cell has bold styling, should only be on the first line of each aml
                            is_first_aml_line = True
                            border_top = 1
                            bold = True

                        self._write_cell(cursor_x, cursor_y, data, 1, is_date, report, fonts, workbook, sheet, XLSX_FONT_SIZE_DEFAULT,
                                         bold, 'white', align, 0, border_top, XLSX_BORDER_COLOR)

                    else:
                        # Empty value
                        self._write_cell(cursor_x, cursor_y, '', 1, False, report, fonts, workbook, sheet, XLSX_FONT_SIZE_DEFAULT, False,
                                         'white', align, 0, border_top, XLSX_BORDER_COLOR)

                    cursor_x += 1
                cursor_x = 0
                cursor_y += 1

            # Draw total line
            total_line = journal_vals['lines'][-1]
            for column in columns:
                data = ''
                align = 'left'

                if total_line.get(column['label'], {}).get('data'):
                    data = total_line[column['label']]['data']

                if 'o_right_alignment' in column.get('class', ''):
                    align = 'right'

                self._write_cell(cursor_x, cursor_y, data, 1, False, report, fonts, workbook, sheet, XLSX_FONT_SIZE_DEFAULT, True,
                                 XLSX_GRAY_200, align, 2, 2)
                cursor_x += 1

            cursor_x = 0

            sheet.set_default_row(20)
            sheet.set_row(0, 30)

            # Tax tables drawing
            if journal_vals.get('tax_summary'):
                self._write_tax_summaries_to_sheet(report, workbook, sheet, fonts, len(columns) + 1, 1, journal_vals['tax_summary'])

        if document_data.get('global_tax_summary'):
            self._write_tax_summaries_to_sheet(
                report,
                workbook,
                workbook.add_worksheet(('Global Tax Summary')[:31]),
                fonts,
                0,
                0,
                document_data['global_tax_summary']
            )

        workbook.close()
        output.seek(0)
        generated_file = output.read()
        output.close()

        return {
            'file_name': 'test.xlsx',
            'file_content': generated_file,
            'file_type': 'xlsx',
        }
