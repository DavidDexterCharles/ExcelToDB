from openpyxl import load_workbook
import re
# import json
import simplejson as json

#https://www.datacamp.com/community/tutorials/python-excel-tutorial#gs.RwPnpzg

# Load in the workbook
wb = load_workbook('test.xlsx')

# Get sheet names
# print wb.get_sheet_names()

sheet = wb.get_sheet_by_name('sheet1')

# Database aspect
# http://www.python-course.eu/sql_python.php
# https://stackoverflow.com/questions/23110383/how-to-dynamically-build-a-json-object-with-python

print sheet['B2'].value

# print len(sheet['B4:J7'][0])
# print str (sheet['B4:J7'][0][0].value) +"  "+ str (sheet['B4:J7'][0][1].value)
# print str (sheet['B4:J7'][1][0].value) +"  "+ str (sheet['B4:J7'][1][1].value)
# print str (sheet['B4:J7'][2][0].value) +"  "+ str (sheet['B4:J7'][2][1].value)

# print "\n\n"


# print len(sheet['A5:J7'])
# print str (sheet['A5:J7'][0][0].value) +"  "+  str (sheet['A5:J7'][1][0].value)
# print str (sheet['A5:J7'][0][1].value) +"  "+  str (sheet['A5:J7'][1][1].value)
# print str (sheet['A5:J7'][0][2].value) +"  "+ str (sheet['A5:J7'][1][2].value)


class Tools(object):

   def getjsondata(self,path):
        with open(path) as data_file:    
            data = json.load(data_file)
        return data


class ExceltoDBAPI(Tools):

    def __init__(self):
        self.mapdb = {}

    def gettables(self, path):
        
        tables= self.getjsondata(path)["tables"]
        print tables
        for x in range(0,len(tables)):
            t={}
            t["tablename"]=tables[x]["name"]
            t["tablerange"]=tables[x]["table"]
            t["tableorient"]=tables[x]["orient"]
            
            self.get_one_table(x,tables[x]["table"],tables[x]["orient"],t)
        print self.mapdb
        # print  json.dumps(self.mapdb)

    def get_one_table(self,tindex,table,orient,t):
        
        tablerange = table.replace("-", ":")
        if orient=="vertical" or orient=="v":
            headercount = len(sheet[tablerange])
        else:
            headercount = len(sheet[tablerange][0])
            vcount = len(sheet[tablerange])
            t["headercount"]=str(headercount)
            t["vcount"]=str(vcount)
            t["headers"]={}
            headers={}
            h={}
            for i in range(0,headercount):
                attrval = sheet[tablerange][0][i].value
                header={}
                header["h"]=attrval
                header["data"]={}
                data={}
              

                print attrval
                for j in range(1,vcount):
                    
                    print sheet[tablerange][j][i].value
                    d={}
                    d["d"]=str(sheet[tablerange][j][i].value)
                    data[j]= d
                header["data"]=data
                headers[i]=header
            t["headers"]=headers   
            self.mapdb[tindex] = t
            
        print tablerange
        print headercount

        







api = ExceltoDBAPI()
api.gettables("tables.json")
'''
# Retrieve cell value 
sheet.cell(row=1, column=2).value

# Print out values in column 2 
for i in range(1, 30):
    print i
    for j in range(1, 24):
        val = sheet.cell(row=i, column=j).value
        # if re.search("None" , str(val)) and r:
        #     c=j
        #     r=i
        
        if not re.search("None" , str(val)):
            print val

'''