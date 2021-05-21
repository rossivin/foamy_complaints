import datetime

plant_dic = {
	"FR": "Turning Point",
	"T": "Mill Street",
	"Q": "Archibald",
	"C": "Creston",
	"L": "London",
	"N": "St. John's",
	"H": "Halifax",
	"M": "Montreal",
	"E": "Edmonton",
}

year_dic = {
	"2": 2012,
	"3": 2013,
	"4": 2014,
	"5": 2015,
	"6": 2016,
	"7": 2017,
	"8": 2018,
	"9": 2019,
	"0": 2020,
	"1": 2021
}

month_dic = {
	"A": 1,
	"B": 2,
	"C": 3,
	"D": 4,
	"E": 5,
	"F": 6,
	"G": 7,
	"H": 8,
	"J": 9,
	"K": 10,
	"L": 11,
	"M": 12,
}


class ProductionCode():

	def __init__(self, code):
		self.original_code = code
		self.code = self.trimProductionCode()
		if self.validateProductionCode():
			self.brewery = self.getBrewery()
			self.date = self.translateProductionCode()
		else:
			self.brewery = "No Data"
			self.date = "No Data"

	def trimProductionCode(self):
		"""Removes all white spaces from production codes"""
		if self.original_code == None:
			return ""
		else:
			return self.original_code.replace(" ","")[:7]

	def validateProductionCode(self):
		"""Validates if the production code is a Labatt code"""
		if len(self.code) > 6:
			if self.code[5] + self.code[6] == "FR":
				if self.code[0] in month_dic.keys() and (self.code[1] + self.code[2] + self.code[3] + self.code[4]).isnumeric():
					return True
				else:
					return False
		if len(self.code) > 4:
			if self.code[3] in plant_dic.keys():
				if self.code[0] in month_dic.keys() and (self.code[1] + self.code[2] + self.code[4]).isnumeric():
					return True
				else:
					return False
			else:
				return False 

	def getBrewery(self):

		if self.code[3].isnumeric():
			if self.code[5] + self.code[6] == "FR":
				return plant_dic["FR"]
			else:
				return False
		else:
			if self.code[3] in plant_dic.keys():
				return plant_dic[self.code[3]]
			else:
				return False

	def translateProductionCode(self):

		day = int(self.code[1] + self.code[2])
		month = month_dic[self.code[0]]
		year = year_dic[self.code[4]]

		return datetime.datetime(year, month, day)


#my_code = ProductionCode("A22H011950")
#print(my_code.brewery)
