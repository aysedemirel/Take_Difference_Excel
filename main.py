import xlrd

# ----------------------------------------------------------------------
def open_file(path):
    """
    Open and read an Excel file
    """
    book = xlrd.open_workbook(path)

    # print number of sheets
    #print(book.nsheets)

    # print sheet names
   # print(book.sheet_names())

    # get the first worksheet
    first_sheet = book.sheet_by_index(0)

    # read a row
    #print(first_sheet.row_values(0))
    #print(first_sheet.col_values(0))
    # read a cell
    #cell = first_sheet.cell(0, 0)
    #print(cell)
    #print(cell.value)
    #col = first_sheet.col(0,0)
    #print(col)
    # read a row slice
    """print(first_sheet.row_slice(rowx=0,
                          start_colx=0,
                          end_colx=2))
    print(first_sheet.row_slice(rowx=1,
                          start_colx=0,
                          end_colx=1))

    print(first_sheet.col_slice(colx=0,
                                start_rowx=0))"""

    return first_sheet.col_values(0)

#----------------------------------------------------------
def take_differences(file1,file2):
   diff = []
   for i in range(len(file1)):
       for j in range(len(file2)):
           if file1[i] == file2[j]:
               break
           elif j == (len(file2)-1):
               print(file1[i])
               diff.append(file1[i])

   for i in range(len(file2)):
       for j in range(len(file1)):
           if file2[i] == file1[j]:
               break
           elif j == (len(file1) - 1):
               print(file2[i])

# ----------------------------------------------------------------------
if __name__ == "__main__":
    path = "excel1.xlsx"
    file1 = open_file(path)

    path2 = "excel2.xlsx"
    file2 = open_file(path2)

    take_differences(file1,file2)
