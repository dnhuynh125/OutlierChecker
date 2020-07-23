import wx
import xlrd
import xlsxwriter
import os


def CreateCharts(tempWorksheet, filePaths, chartTitle, chartSeriesTitle, numTests, row, col, chartLocX, chartLocY):
    """
        creates line charts for each column of test data

        :params:
            tempWorksheet: object variable of an excel worksheet
            filePaths: string list of file path directories used only for length of list which is number of files
            chartTitle: string value of chart title 
            chartSeriesTitle: string value of chart series title
            numTests: integer value of number of test columns
            row: integer value of current row location for writing
            col: integer value of current col location for writing
            chartLocX: character value of current excel col letter coordinate
            chartLocY: integer value of current excel row coordinate
            
        :returns:
            row: integer value of updated row coordinate
    """
    
    chart = workbook.add_chart({'type': 'line'})
    seriesCounter = 1
    for i in range(len(filePaths)):
        chart.add_series({
            'name': chartSeriesTitle + '_' + str(seriesCounter),                        # name series with string passed in
            'values': ['Sheet1', row, col, row, numTests + 1]                           # [SheetName, startRow, startCol, endRow, endCol] cell ranges
            })
        seriesCounter += 1
        row += 1
    chart.set_title({'name': chartTitle})
    chart.set_legend({'position': 'bottom'})
    worksheet.insert_chart(str(chartLocX) + str(chartLocY), chart)
    
    return row

    
def MergeCells(sampFilePaths, prodFilePaths):
    """
        merges cells over a specific range to create a title for a block of data

        :params:
            sampFilePaths: string list of file path directories of sample files
            prodFilePaths: string list of file path directories of production files
            
        :returns:
            N/A
    """
    
    mergeFormat = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
    
    # merge cells from start to end value to create title for samp avg values
    start = 1
    end = len(sampFilePaths)
    worksheet.merge_range("A" + str(start) + ":B" + str(end), "SAMP AVG", mergeFormat)
    
    # merge cells from start to end value to create title for samp stdev values
    start = len(sampFilePaths) + 2
    end = len(sampFilePaths) * 2 + 1
    worksheet.merge_range("A" + str(start) + ":B" + str(end), "SAMP STDEV", mergeFormat)

    # merge cells from start to end value to create title for prod avg values
    start = len(sampFilePaths) * 2 + 5
    end = len(sampFilePaths) * 2 + 4 + len(prodFilePaths)
    worksheet.merge_range("A" + str(start) + ":B" + str(end), "PROD AVG", mergeFormat)

    # merge cells from start to end value to create title for prod stdev values
    start = len(sampFilePaths) * 2 + 6 + len(prodFilePaths)
    end = len(sampFilePaths) * 2 + 5 + len(prodFilePaths) * 2
    worksheet.merge_range("A" + str(start) + ":B" + str(end), "PROD STDEV", mergeFormat)

    # merge cells from start to end value to create title for avg of stdev values
    start = len(sampFilePaths) * 2 + 9 + len(prodFilePaths) * 2
    end = len(sampFilePaths) * 2 + 9 + len(prodFilePaths) * 2
    worksheet.merge_range("A" + str(start) + ":B" + str(end), "AVG STDEV", mergeFormat)

    # merge cells from start to end value to create title for stdev of stdev values
    start = len(sampFilePaths) * 2 + 10 + len(prodFilePaths) * 2
    end = len(sampFilePaths) * 2 + 10 + len(prodFilePaths) * 2
    worksheet.merge_range("A" + str(start) + ":B" + str(end), "STDEV STDEV", mergeFormat)


def CarryOverStatInfo(tempWorksheet, filePaths, row, col):
    """
        carries over statisical information such as average and standard deviation from
        tempWorksheet over to final worksheet

        :params:
            tempWorksheet: object variable of an excel worksheet
            filePaths: string list of file path directories
            row: integer value of current row location for writing
            col: integer value of current col location for writing
        :returns:
            row: integer value of updated row coordinate
            col: integer value of updated col coordinate
    """
    
    for i in filePaths:                                                             # for each file
        for val in tempWorksheet.row_values(row, start_colx = 2, end_colx = None):      # for each value in specified row
            worksheet.write(row, col, val)                                                  # write value to temp excel file
            col += 1                                                                        # inc col to write adjacently to right of cell
        # reset to start writing next row of data
        row += 1
        col = 2
        
    return row, col

    
def GetColLetter(number):
    """
        converts column integer value into character value and returns as a string

        :params:
            number: integer value of column coordinate in excel file starting at 0 as column A

        :returns:
            string: string value of column lettering
    """
    
    string = ""
    while number > 0:
        number, remainder = divmod(number - 1, 26)
        string = chr(65 + remainder) + string                   # adds character by character in reverse order to get column letter
        
    return string

        
def WriteRowFromFiles(tempWorksheet, filePaths, row, col, rowLoc):
    """
        writes row splice from every file in the filePath list passed in and
        returns updated row and col coordinates

        :params:
            tempWorksheet: object variable of an excel worksheet
            filePaths: string list of file path directories
            row: integer value of current row location for writing
            col: integer value of current col location for writing
            rowLoc: integer value of row being referenced for row slice
            
        :returns:
            row: integer value of updated row coordinate
            col: integer value of updated col coordinate
    """
    
    for i in filePaths:
        workbook = xlrd.open_workbook(i)
        worksheet = workbook.sheet_by_index(0)
        
        for val in worksheet.row_values(rowLoc, start_colx = 2, end_colx = None):
            tempWorksheet.write(row, col, val)
            col += 1
        row += 1
        col = 2

    return row, col    


def ReceiveFiles(prompt):
    """
        returns multiple excel file paths in list format by using file dialog box

        :params:
            prompt: string value to display in title bar of file dialog
            
        :returns:
            filePaths: string list of file path directories
    """
    
    openFileDialog = wx.FileDialog(None, prompt, "", "", "Excel Files (*.xls,*.xlsx)|*.xls;*.xlsx", wx.FD_MULTIPLE)
    openFileDialog.ShowModal()
    filePaths = openFileDialog.GetPaths()
    openFileDialog.Destroy()
    
    return filePaths                                                              


if __name__ == "__main__":
    app = wx.App()

    sampFilePaths = ReceiveFiles("Open Sample Files")
    prodFilePaths = ReceiveFiles("Open Production Files")
    
    if sampFilePaths and prodFilePaths:
        tempWorkbook = xlsxwriter.Workbook('temp.xlsx')
        tempWorksheet = tempWorkbook.add_worksheet()

        # initial writing location of temp excel sheet
        row = 0
        col = 2
        
        # write avg values of all samp files into tempWorksheet which is always on row 16
        row, col = WriteRowFromFiles(tempWorksheet, sampFilePaths, row, col, 16)
        row += 1
        sampStartSTDEV = row                                                            # save row coordinate of sample stdev for later
        # write stdev values of all samp files into tempWorksheet which is always on row 17
        row, col = WriteRowFromFiles(tempWorksheet, sampFilePaths, row, col, 17)      
        row += 3
        prodStartAVG = row                                                              # save row coordinate of prod avg for later
        # write avg values of all prod files into tempWorksheet which is always on row 16
        row, col = WriteRowFromFiles(tempWorksheet, prodFilePaths, row, col, 16)
        row += 1
        prodStartSTDEV = row                                                            # save row coordinate of production stdev for later
        # write stdev values of all prod files into tempWorksheet which is always on row 17
        row, col = WriteRowFromFiles(tempWorksheet, prodFilePaths, row, col, 17)
        row += 3
        tempWorkbook.close()

        tempWorkbook = xlrd.open_workbook('temp.xlsx')
        tempWorksheet = tempWorkbook.sheet_by_index(0)
        finalPath = os.path.dirname(prodFilePaths[0])                       # get directory of final data folder by referencing first prod file
        workbook = xlsxwriter.Workbook(finalPath + '\Final Review.xlsx')
        worksheet = workbook.add_worksheet()

        numTests = len(tempWorksheet.row_values(0, start_colx = 2, end_colx = None))
        for i in range(numTests):
            currCol = GetColLetter(col + 1)                                             # get current column letter
            sampRange = currCol + "1:" + currCol + str(len(sampFilePaths))              # range of samp avg values          
            prodRange = currCol + str(prodStartAVG + 1) + ":" + currCol + str(prodStartAVG + len(prodFilePaths))    # range of prod avg values
            stdevFormula = "=STDEV(" + sampRange + "," + prodRange + ")"                # create stdev formula to include all avg values
            worksheet.write(row, col, stdevFormula)
            col += 1

        row += 1
        col = 2
        for i in range(numTests):
            currCol = GetColLetter(col + 1)                                             # get current column letter
            sampRange = currCol + str(sampStartSTDEV + 1) + ":" + currCol + str(sampStartSTDEV + len(sampFilePaths))    # range of samp stdev values          
            prodRange = currCol + str(prodStartSTDEV + 1) + ":" + currCol + str(prodStartSTDEV + len(prodFilePaths))    # range of prod stdev values
            stdevFormula = "=STDEV(" + sampRange + "," + prodRange + ")"                # create stdev formula to include all stdev values
            worksheet.write(row, col, stdevFormula)
            col += 1
            
        # save starting point for chart locations
        chartLocY = row + 3

        # initial writing location of final excel sheet
        row = 0
        col = 2
        for i in range(2):
            row, col = CarryOverStatInfo(tempWorksheet, sampFilePaths, row, col)
            row += 1
        row += 2
        for i in range(2):
            row, col = CarryOverStatInfo(tempWorksheet, prodFilePaths, row, col)
            row += 1

        MergeCells(sampFilePaths, prodFilePaths)
        
        # initial writing location of temp excel sheet
        row = 0
        col = 2     
        chartLocX = ['A', 'I']
        
        chartSampTitle = ['Sample Averages', 'Sample Standard Deviations']
        chartSampSeriesTitle = ['SampAVG', 'SampSTDEV']    
        for i in range(2):
            row = CreateCharts(tempWorksheet, sampFilePaths, chartSampTitle[i], chartSampSeriesTitle[i], numTests, row, col, chartLocX[i], chartLocY)
            row += 1
        row += 2
        chartLocY += 15

        chartProdTitle = ['Production Averages', 'Production Standard Deviations']
        chartProdSeriesTitle = ['ProdAVG', 'ProdSTDEV']
        for i in range(2):
            row = CreateCharts(tempWorksheet, prodFilePaths, chartProdTitle[i], chartProdSeriesTitle[i], numTests, row, col, chartLocX[i], chartLocY)
            row += 1
        row += 2
        chartLocY -= 8

        chartSTDEVTitle = ['STDEV_AVG', 'STDEV_STDEV']
        chartAvgSTDEVStdevSTDEV = workbook.add_chart({'type': 'line'})
        for i in range(2):
            chartAvgSTDEVStdevSTDEV.add_series({
                'name': chartSTDEVTitle[i],                                                 # name series with string passed in
                'values': ['Sheet1', row, col, row, numTests + 1]                           # use values of row & col coordinates
                })
            row += 1
        chartAvgSTDEVStdevSTDEV.set_title({'name': 'STDEV AVG/STDEV STDEV'})
        chartAvgSTDEVStdevSTDEV.set_legend({'position': 'bottom'})
        worksheet.insert_chart('Q' + str(chartLocY), chartAvgSTDEVStdevSTDEV)

        workbook.close()
        os.remove('temp.xlsx')
    else:
        if not sampFilePaths and prodFilePaths:
            wx.MessageBox("You did not select any sample files. Quitting program...", "ERROR", wx.OK|wx.ICON_ERROR)
        elif sampFilePaths and not prodFilePaths:
            wx.MessageBox("You did not select any lot files. Quitting program...", "ERROR", wx.OK|wx.ICON_ERROR)
        else:
            wx.MessageBox("You did not select any sample or lot files. Quitting program...", "ERROR", wx.OK|wx.ICON_ERROR)
            
    app.MainLoop()




