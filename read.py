import xlrd
import json

loc = ("hello_xlsx_.xls")

wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

def extractColumnsHandler(columns,extractColumns):
    finalExtractColumns = []
    for (key,value) in enumerate(columns):
        try:
            elmIndexFound = extractColumns.index(value)
            finalExtractColumns.append({"columnIndex":key,"columnName":value})
        except ValueError:
            print("No element found:"+value)
    return finalExtractColumns;

extractColumns = ["id","Test"];
columns = sheet.row_values(0); 

finalExtractColumns_ = extractColumnsHandler(columns,extractColumns)

def extractAllRowHandler(finalExtractColumns_):
    extractAllRow = [];
    for row in range(sheet.nrows):
        # skip first row - they are columns
        if row == 0:
            continue;
        rowValue = sheet.row_values(row);
        rowTempObj={};
        for dict_ in finalExtractColumns_:
            dictColumnIndex = dict_["columnIndex"];
            columnName = dict_["columnName"];
            rowTempObj[columnName] = rowValue[dictColumnIndex]
        extractAllRow.append(rowTempObj);
    return extractAllRow;


allRows = extractAllRowHandler(finalExtractColumns_)

with open("output.json", "w") as file:
    json.dump(allRows, file)
