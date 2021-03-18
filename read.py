import xlrd
import datetime
import json
import math

loc = ("hello_xlsx_.xls")

wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

def extractColumnsHandler(columns,extractColumns):
    finalExtractColumns = []
    for (key,value) in enumerate(columns):
        try:
            elmIndexFound = extractColumns.index(value.strip())
            finalExtractColumns.append({"columnIndex":key,"columnName":value})
        except ValueError:
            print("No element found:"+value)
    return finalExtractColumns;

extractColumns = ["id","Test"];
#extractColumns = ["Reg.No","Patient Name","Date of Registration","Contact No","email Address","Test Name"];
columns = sheet.row_values(0); 

finalExtractColumns_ = extractColumnsHandler(columns,extractColumns)
#print(finalExtractColumns_);

def xldate_to_datetime(xldatetime): #something like 43705.6158241088

    tempDate = datetime.datetime(1899, 12, 31)
    (days, portion) = math.modf(xldatetime)

    deltaDays = datetime.timedelta(days=days)
    #changing the variable name in the edit
    secs = int(24 * 60 * 60 * portion)
    detlaSeconds = datetime.timedelta(seconds=secs)
    TheTime = (tempDate + deltaDays + detlaSeconds )
    return TheTime.strftime("%Y-%m-%d %H:%M:%S")


def extractAllRowHandler(finalExtractColumns_):
    extractAllRow = [];
    lastRecord = {};
    for row in range(sheet.nrows):
        # skip first row - they are columns
        if row == 0:
            continue;
        rowValue = sheet.row_values(row);
        rowTempObj={};
        dataExist = True;
        for dict_ in finalExtractColumns_:
            dictColumnIndex = dict_["columnIndex"];
            columnName = dict_["columnName"];

            rowTempObj[columnName] = rowValue[dictColumnIndex]

            if (dictColumnIndex == 3):
                #print("===========");
                if len(str(rowValue[dictColumnIndex]))!=0:
                   rowTempObj[columnName] = xldate_to_datetime(rowValue[dictColumnIndex]);
            if (dictColumnIndex == 6 and len(rowValue[dictColumnIndex].strip())==0):
                #print("===========");
                dataExist = False;
                break;
            elif len(str(rowValue[dictColumnIndex]))==0:
                #print("======<<");
                rowTempObj = { **lastRecord};
                #print(lastRecord);
                continue;
        
        if dataExist == True:        
            lastRecord = rowTempObj;
            extractAllRow.append(rowTempObj);
"""            
        if row == 9:
            #print(extractAllRow);
            print(json.dumps(extractAllRow))
            exit();
"""
    return extractAllRow;


allRows = extractAllRowHandler(finalExtractColumns_)

with open("output.json", "w") as file:
    json.dump(allRows, file)
