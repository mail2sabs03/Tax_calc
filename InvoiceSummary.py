class InvoiceSummary:

    def __init__(self, company_name, city, state, zip_code, country):
        self.company_name = company_name
        self.city = city
        self.state = state
        self.zip_code = zip_code
        self.country = country
        self.sub_total = 0
        self.tax = 0
        self.sub_total_with_tax = 0

    def set_sub_total(self, sub_total):
        self.sub_total = sub_total

    def get_sub_total(self):
        return self.sub_total

    def set_tax(self, tax):
        self.tax = tax

    def get_tax(self):
        return self.tax
