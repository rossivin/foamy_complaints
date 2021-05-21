import math
import openpyxl as xl
import datetime
import production_code as pc

brewery_dictionary = {
		"Creston BC Canada": "Creston",
		"Edmonton, Alberta, Canada (Labatt)": "Edmonton",
		"Halifax, Nova Scotia, Canada (Labatt)": "Halifax",
		"London, Ontario, Canada (Labatt)": "London",
		"Montreal, Quebec, Canada (Labatt)": "Montreal",
		"St Johns, Newfoundland, Canada (Labatt)": "St. John's",
		"Turning Point (Stanley Park)": "Turning Point"
	}

class Complaint():

	def __init__(self, sheet, r, one_brewery_brands, one_country_brands):

		self.customer_code = self.getCustomerCode(sheet, r)
		self.brand = self.getBrand(sheet, r)
		self.brewery = self.getBrewery(sheet, r, one_brewery_brands)
		self.case_number = self.getCaseNumber(sheet, r)
		self.complaint_date = self.getComplaintDate(sheet, r)
		self.country = self.getCountry(sheet, r, one_country_brands)
		self.incident_date = self.getIncidentDate(sheet, r)
		self.production_date = self.getProductionDate(sheet, r)

	def getCustomerCode(self, sheet, r):
		return pc.ProductionCode(sheet.cell(row = r, column = 12).value)
	
	def getBrand(self, sheet, r):
		return sheet.cell(row = r, column = 4).value

	def getBrewery(self, sheet, r, one_brewery_brands):
		
		if self.brand in one_brewery_brands.keys():
			return one_brewery_brands[self.brand]
		elif self.customer_code.validateProductionCode():
			return self.customer_code.brewery
		else:
			facility = sheet.cell(row = r, column = 6).value
			if facility in brewery_dictionary.keys():
				return brewery_dictionary[facility]
			elif facility == "":
				return "Unconfirmed"
			else:
				return facility

	def getCaseNumber(self, sheet, r):
		return sheet.cell(row = r, column = 2).value

	def getComplaintDate(self, sheet, r):
		return sheet.cell(row = r, column = 3).value

	def getCountry(self, sheet, r, one_country_brands):
		if self.customer_code.validateProductionCode():
			return "CANADA"
		elif self.brand in one_country_brands.keys():
			return one_country_brands[self.brand]
		else:
			data_country = sheet.cell(row = r, column = 14).value
			if data_country == "UNITED STATES":
				return "USA"
			else:
				return data_country

	def getIncidentDate(self, sheet, r):
		return sheet.cell(row = r, column = 31).value

	def getProductionDate(self, sheet, r):
		if self.customer_code.validateProductionCode():
			return self.customer_code.date
		else:
			return sheet.cell(row = r, column = 29).value

def headingPrinter(sheet):
	sheet.cell(row = 1, column = 1).value = "Case Number"
	sheet.cell(row = 1, column = 2).value = "Brand"
	sheet.cell(row = 1, column = 3).value = "Brewery"
	sheet.cell(row = 1, column = 4).value = "Country"
	sheet.cell(row = 1, column = 5).value = "Production Date"
	sheet.cell(row = 1, column = 6).value = "Complaint Date"
	sheet.cell(row = 1, column = 7).value = "Incident Date"

def complaintPrinter(complaint_data, sheet, r):

	sheet.cell(row = r, column = 1).value = complaint_data.case_number
	sheet.cell(row = r, column = 2).value = complaint_data.brand
	sheet.cell(row = r, column = 3).value = complaint_data.brewery
	sheet.cell(row = r, column = 4).value = complaint_data.country
	sheet.cell(row = r, column = 5).value = complaint_data.production_date
	sheet.cell(row = r, column = 6).value = complaint_data.complaint_date
	sheet.cell(row = r, column = 7).value = complaint_data.incident_date

def singleBreweryBrands(sheet):
	#[brand]: [facility]
	single_brewery_brands = {}
	for i in range(3, sheet.max_row + 1):
		single_brewery_brands[sheet.cell(column = 1, row = i).value] = sheet.cell(column = 2, row = i).value

	return single_brewery_brands

def singleCountryBrands(sheet):
	single_country_brands = {}
	i = 3
	while sheet.cell(column = 4, row = i).value:
		single_country_brands[sheet.cell(column = 4, row = i).value] = sheet.cell(column = 5, row = i).value
		i += 1

	return single_country_brands


def main():

	wb = xl.load_workbook("data.xlsx")
	ws = wb["Data"]
	ws_setup = wb["Setup"]
	ws_destiny = wb["Polished Data"]
	one_brewery_brands_dic = singleBreweryBrands(ws_setup)
	one_country_brands_dic = singleCountryBrands(ws_setup)

	for i in range(4, ws.max_row+1):
		print(i)
		current_complaint = Complaint(ws, i, one_brewery_brands_dic, one_country_brands_dic)
		headingPrinter(ws_destiny)
		complaintPrinter(current_complaint, ws_destiny, i-2)

	wb.save("data.xlsx")

main()
