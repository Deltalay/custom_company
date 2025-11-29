from odoo.models import Model
import io
from odoo.tools.misc import xlsxwriter

class AccountReport(Model):
    _inherit = "account.journal.report.handler"

    def export_to_xlsx(self, options, response=None):
        self.ensure_one()
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {
            'in_memory': True,
            'strings_to_formulas': False,
        })

        print_options = self.get_options(previous_options={**options, 'export_mode': 'print'})
        if print_options['sections']:
            reports_to_print = self.env['account.report'].browse([section['id'] for section in print_options['sections']])
        else:
            reports_to_print = self

        reports_options = []
        for report in reports_to_print:
            report_options = report.get_options(previous_options={**print_options, 'selected_section_id': report.id})
            reports_options.append(report_options)
            report._inject_report_into_xlsx_sheet(report_options, workbook, workbook.add_worksheet(report.name[:31]))

        self._add_options_xlsx_sheet(workbook, reports_options)

        workbook.close()
        output.seek(0)
        generated_file = output.read()
        output.close()

        return {
            'file_name': "hello_world.xlsx",
            'file_content': generated_file,
            'file_type': 'xlsx',
        }
