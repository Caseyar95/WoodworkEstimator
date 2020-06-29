import xlrd
loc = ("woodspecies.xlsx")

class Client:
    def __init__(self,name):
        self.name = name

    # def getName(self):
    #     return self.name

class Wood:
    def __init__(self,woodType,bf):
        self.wood = woodType
        self.bf = bf
        self.bfCost = 0

    def setBfCost(self, location):
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

    # def getWoodType(self):
    #     return self.wood

    # def getBf(self):
    #     return self.bf

    def costForWood(self, waste):
        return (self.bf * self.bfCost + (waste * self.bf))


class Project:
    def __init__(self,projId,hours):
        self.hours = hours
        self.projId = projId
        self.cost = 0
        self.hourlyRate = 13
        self.miscRate = .1
    
    # def getProj(self):
    #     return self.projId

    # def getHours(self):
    #     return self.hours

    def setRate(self, rate):
        self.hourlyRate = rate

    def setMiscRate(self, misc):
        self.miscRate = misc
    
    def getCost(self, woodCost):
        self.cost = (woodCost * 3) + (woodCost * self.miscRate) + (self.hourlyRate * self.hours)
        return self.cost

client1 = Client("Casey Robertson")

wood1 = Wood("Walnut", 10)
wood1.setBfCost("woodspecies.xlsx")
woodC = wood1.costForWood(.2)

sideTable = Project("side table", 10)
sideTable.setRate(10)
sideTable.setMiscRate(.05)
projectCost = sideTable.getCost(woodC)


#test
print(client1.name)
print(wood1.wood, wood1.bf, round(wood1.getBfCost(), 2))
print(sideTable.projId, sideTable.hours, round(projectCost, 2))

workTup = (wood1.wood, wood1.bf, sideTable.projId, sideTable.hours, sideTable.cost)
workQueue = {client1.name : workTup}

print(workQueue[client1.name])

