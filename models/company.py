from odoo import api, fields
from odoo.models import Model


class CustomCompany(Model):
    _inherit = "res.company"
    name_translate = fields.Char("Company Name", translate=True)
    street_translate = fields.Char("Address", translate=True)

    @api.onchange("name_translate")
    def _onchange_name_translate(self):
        lang = "en_US"
        for company in self:
            value = company.with_context(lang=lang).name_translate
            company.name = value or company.name

    @api.onchange("street_translate")
    def _onchange_street_translate(self):
        lang = "en_US"
        for company in self:
            value = company.with_context(lang=lang).street_translate
            company.street = value or company.street

            def write(self, vals):
                lang = "en_US"
                for company in self:
                    vals_copy = vals.copy()
                    if "name_translate" in vals_copy:
                        en_value = company.with_context(lang=lang).name_translate
                        if en_value:
                            vals_copy["name"] = en_value
                    if "street_translate" in vals_copy:
                        en_value = company.with_context(lang=lang).street_translate
                        if en_value:
                            vals_copy["street"] = en_value
                    super(CustomCompany, company).write(vals_copy)
                return True
