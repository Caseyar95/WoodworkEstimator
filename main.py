import xlrd
loc = ("woodspecies.xlsx")

class Client:
    def __init__(self,name):
        self.name = name

    def getName(self):
        return self.name

class Wood:
    def __init__(self,woodType,bf):
        self.wood = woodType
        self.bf = bf
        self.bfCost = 0

    def findBfCost(self, location):
        wb = xlrd.open_workbook(location)
        sheet = wb.sheet_by_index(0)
        for rows in range(sheet.nrows):
            row_value = sheet.row_values(rows)
            if row_value[0] == self.wood:
                self.bfCost = row_value[1]
        if self.bfCost == 0:
            print("wood type not found")

    def getBfCost(self):
        return self.bfCost

    def getWoodType(self):
        return self.wood

    def getBf(self):
        return self.bf

    def costForWood(self, waste):
        return (self.bf * self.bfCost * waste)


class Project:
    def __init__(self,hours,projId):
        self.hours = hours
        self.projId = projId
        self.cost = 0
    
    def getProj(self):
        return self.projId

    def getHours(self):
        return self.hours
    
    def getCost(self, woodCost, miscCost, laborCost):
        self.cost = woodCost + miscCost + (laborCost * self.hours)
        return self.cost



