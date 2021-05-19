import math
import openpyxl

class Complaint():

	def __init__(self, sheet, r):

		self.brand = self.getBrand(sheet, r)
		self.brewery = self.getBrewery(sheet, r)
		self.case_number = self.getCaseNumber(sheet, r)
		self.complaint_date = self.getComplaintDate(sheet, r)
		self.country = self.getCountry(sheet, r)
		self.incident_date = self.getIncidentDate(sheet, r)
		self.production_date = self.getProductionDate(sheet, r)
		
		def getBrand(self, sheet, r):
			return sheet.cell(row = r, column = 4).value

		def getBrewery(self, sheet, r):
			return sheet.cell(row = r, column = 6).value

		def getCaseNumber(self, sheet, r):
			return sheet.cell(row = r, column = 1).value

		def getComplaintDate(self, sheet, r):
			return sheet.cell(row = r, column = 3).value

		def getCountry(self, sheet, r):
			return sheet.cell(row = r, column = 14).value
			#Add code to look at brand (to see brands only produced here), and production code (to check for Canadian breweries)

		def getIncidentDate(self, sheet, r):
			return sheet.cell(row = r, column = 31).value

		def getProductionDate(self, sheet, r):
			return sheet.cell(row = r, column = 29).value

def convertProductionDate(production_code):

	#1 - Month
	#2,3 - Day
	#4 - Plant
	#5 - Year

	valid_month = "ABCDEFGHJKLM"

	production_code = production_code.strip()

	month_code = {
		"A": "Jan",
		"B": "Feb",
		"C": "Mar",
		"D": "Apr",
		"E": "May",
		"F": "Jun",
		"G": "Jul",
		"H": "Aug",
		"J": "Sep",
		"K": "Oct",
		"L": "Nov",
		"M": "Dec",
	}

	plant_code = {
		"M": "Montreal",
		"C": "Creston",
		"L": "London",
		"E": "Edmonton",
		"H": "Halifax",
		"N": "St. John's",
		"Q": "Archibald",
		"T": "Mill Street"
	}

def overwriteComplaints(overwrite_data_sheet, destination_sheet):






