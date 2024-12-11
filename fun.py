import win32com
from win32com.client import Constants as catEnum
import os
from pyxll import xl_app
from win32com.client import constants


class navisComApi(object):
    @staticmethod
    def getNavisDoc():
        """Navisworks 2025 """
        doc = win32com.client.dynamic.Dispatch("Navisworks.document.22") # 22 for navis2025, 18 for navis 2021
        doc.visible=True
        doc.stayopen()
        return doc
    @staticmethod
    def getOpenedFileName(state):
        try:
            return state.GetCurrentFilename()
        except:
            return None
    @staticmethod
    def getClashTestName(clashtest) -> str:
        return clashtest.Name
    @staticmethod
    def getClashResults_underClashTest(clashtest):
        crs=clashtest.results()
        return crs
    @staticmethod
    def getState(doc):
        state=doc.state
        return state
    @staticmethod
    def openFile(doc,filePath):
        doc.OpenFile(filePath)
    @staticmethod
    def appendFile(doc,filePath):
        doc.AppendFile(filePath)
    @staticmethod
    def saveAsFile(doc,filePath):
        doc.SaveAs(filePath)
    @staticmethod
    def selectAll(osel):
        osel.SelectAll()
    @staticmethod
    def selectInvert(osel):
        osel.Invert()
    @staticmethod
    def get_ClashResult_Status(clashResult) -> str:
        '''
            status=0 -> new
            status=1 -> Active
            status=2 -> Approved
            status=3 -> Resolved
            status=4 -> Reviewed
        '''
        status=""
        if clashResult.status==0:
            status="New"
        elif clashResult.status==1:
            status="Active"
        elif clashResult.status==2:
            status="Approved"
        elif clashResult.status==3:
            status="Resolved"
        elif clashResult.status==4:
            status="Reviewed"                        
        
        return status
            
        
    @staticmethod
    def get_ClashResultName(clashResult):
        return clashResult.Name
    @staticmethod
    def get_ClashResult_ClashCenter(clashResult) ->list:
        return [clashResult.GetClashCenter().data1,clashResult.GetClashCenter().data2,clashResult.GetClashCenter().data3]
    @staticmethod
    def createViewPointWithHiddenElement(state, viewpointName="temp Viewpoint"):
        currentview = state.CurrentView.Copy()
        view = state.ObjectFactory(11)
        view.name = viewpointName
        view.anonview = currentview
        view.ApplyHideAttribs = True
        view.ApplyMaterialAttribs = True
        state.SavedViews().Add(view)
        return view
    @staticmethod
    def getClashTests(state):
        myclash = None
        for item in state.Plugins():
            myclash = item
            if myclash != None:
                break
        if myclash != None:
            clashtestcollection = myclash.Tests()
            return clashtestcollection
        else:
            return None
    @staticmethod
    def exportClashTestsDataToExcel():
        doc = navisComApi.getNavisDoc()
        state = navisComApi.getState(doc)
        clashtests = navisComApi.getClashTests(state)

        excelapp = my_excel.getExcelApp()
        myworkbook = my_excel.getWorkbookByFilePath()
        mysheet = myworkbook.Sheets[1]
        excelapp.Visible = True
        mysheet.UsedRange.Delete()
        row = 2

        # region excel first line
        mysheet.Range("a1").Value = "ClashTest_Name";
        mysheet.Range("b1").Value = "New";
        my_excel.changeRangeColor(mysheet.Range("b1").EntireColumn, 255, 0, 0)
        mysheet.Range("c1").Value = "Active";
        my_excel.changeRangeColor(mysheet.Range("c1").EntireColumn, 255, 192, 0)
        mysheet.Range("d1").Value = "Reviewed";
        my_excel.changeRangeColor(mysheet.Range("d1").EntireColumn, 0, 176, 240)
        mysheet.Range("e1").Value = "Approved";
        my_excel.changeRangeColor(mysheet.Range("e1").EntireColumn, 0, 255, 0)
        mysheet.Range("f1").Value = "Resolved";
        my_excel.changeRangeColor(mysheet.Range("f1").EntireColumn, 255, 255, 0)
        mysheet.Range("g1").Value = "New+Active";
        mysheet.Range("h1").Value = "New+Active+Reviewed";
        mysheet.Range("i1").Value = "Total";
        mysheet.Range("j1").Value = "Ungrouped";
        mysheet.Range("k1").Value = "SelectionA";
        mysheet.Range("l1").Value = "SelectionB";
        mysheet.Range("m1").Value = "ClashTestGuid";
        mysheet.Range("aa1").Value = "SelectionATreeOrModelItemInstanceGuid";
        mysheet.Range("ab1").Value = "SelectionBTreeOrModelItemInstanceGuid";
        # endregion

        for clashtest in clashtests:
            countNew = 0
            countActive = 0
            countReviewed = 0
            countApproved = 0
            countResolved = 0
            countTotal = 0
            countUngrouped = 0
            groupPathList = []
            clashresultlist = []
            for clashresult in clashtest.results():
                clashresultlist.append(clashresult)
                if len(clashresult.GroupPath.split("\n")) > 1:
                    groupPathList.append(clashresult.GroupPath.split('\n')[0])
            groupPathList = my_list.remove_DuplcatedString_InList(groupPathList)

            # status=0 -> new
            # status=1 -> Active
            # status=4 -> Reviewed
            # status=2 -> Approved
            # status=3 -> Resolved

            for i in range(len(groupPathList)):
                tempGroupedResultList = [x for x in clashresultlist if x.GroupPath.split('\n')[0] == groupPathList[i]]
                if len([x for x in tempGroupedResultList if x.status == 0]) > 0:
                    countNew += 1
                elif len([x for x in tempGroupedResultList if x.status == 1]) > 0:
                    countActive += 1
                elif len([x for x in tempGroupedResultList if x.status == 4]) > 0:
                    countReviewed += 1
                elif len([x for x in tempGroupedResultList if x.status == 2]) > 0:
                    countApproved += 1
                elif len([x for x in tempGroupedResultList if x.status == 3]) > 0:
                    countResolved += 1

            countNew = countNew + len(
                [x for x in clashresultlist if len(x.GroupPath.split('\n')) == 1 and x.status == 0])
            countActive = countActive + len(
                [x for x in clashresultlist if len(x.GroupPath.split('\n')) == 1 and x.status == 1])
            countReviewed = countReviewed + len(
                [x for x in clashresultlist if len(x.GroupPath.split('\n')) == 1 and x.status == 4])
            countApproved = countApproved + len(
                [x for x in clashresultlist if len(x.GroupPath.split('\n')) == 1 and x.status == 2])
            countResolved = countResolved + len(
                [x for x in clashresultlist if len(x.GroupPath.split('\n')) == 1 and x.status == 3])
            countTotal = countNew + countActive + countReviewed + countApproved + countResolved
            countUngrouped = countTotal - len(groupPathList)

            mysheet.Range("a" + str(row)).Value = clashtest.Name
            mysheet.Range("b" + str(row)).Value = countNew
            mysheet.Range("c" + str(row)).Value = countActive
            mysheet.Range("d" + str(row)).Value = countReviewed
            mysheet.Range("e" + str(row)).Value = countApproved
            mysheet.Range("f" + str(row)).Value = countResolved
            mysheet.Range("g" + str(row)).Value = countNew + countActive
            mysheet.Range("h" + str(row)).Value = countNew + countActive + countReviewed
            mysheet.Range("i" + str(row)).Value = countTotal
            mysheet.Range("j" + str(row)).Value = countUngrouped

            row += 1

        mysheet.Range("a1").EntireColumn.AutoFit()
        print("finished")


"""excel"""
class my_excel:
    @staticmethod
    def getExcelApp():
        firstopen =  win32com.client.dynamic.Dispatch("Excel.Application")
        excel = xl_app()
        return excel
    @staticmethod
    def getOpenedWorkbook(openedWorkBook_Name):
        try:
            return my_excel.getExcelApp().Workbooks[openedWorkBook_Name]
        except:
            print("please check workbook name")
            return None
    @staticmethod
    def getWorkbookByFilePath(excelFileFullPath=r"C:\Template\tempExcel.xlsx"):
        if os.path.isfile(excelFileFullPath):
            try:
                myworkbook=my_excel.getExcelApp().Workbooks.Open(excelFileFullPath)
                return myworkbook
            except:
                print("open file error, please check if it is excel file.")
                return None
        else:
            print("file path error")
            return None
    @staticmethod
    def runExcelMacro(workbookName,sheetName,macroName):
        tempstring=workbookName+"!"+sheetName+"."+macroName
        my_excel.getExcelApp().Run(tempstring)
    @staticmethod
    def getCurrentCell(excelapp):
        currentCell=excelapp.ActiveCell
        return currentCell
    @staticmethod
    def changeRangeColor(range,r,g,b):
        range.Interior.Color=my_color.colorRGB2ExcelColor(r,g,b)
    @staticmethod
    def addRangeComment(range,string_comments):
        range.AddComment(string_comments)
    @staticmethod
    def getColumnLastRange(mysheet,columnString="a"):
        range1 = mysheet.Range(columnString+ str(mysheet.Rows.Count)).End(Direction=constants.xlUp)
        return range1
    @staticmethod
    def getColumnLastRange_By_ColumnNumber(mysheet,colnum=1):
        range1=mysheet.Cells(mysheet.Rows.Count,colnum).End(Direction=constants.xlUp)
        # range1 = mysheet.Range(columnString+ str(mysheet.Rows.Count)).End(Direction=constants.xlUp)
        return range1
    @staticmethod
    def getRowLastRange(mysheet,rowNumber=1):
        # Excel.Range
        # range1 = myWorkSheet.Cells[RowNumberString, myWorkSheet.Columns.Count].End[Excel.XlDirection.xlToLeft];

        range1 = mysheet.Cells(rowNumber,mysheet.Columns.Count).End(Direction=constants.xlToLeft)
        return range1
    @staticmethod
    def searchLastRangeAppeared(myRange,searchString):
        '''这里的myRange是一个区域'''
        lastrange = myRange.Find(
            What=searchString,
            After=myRange.Cells(myRange.Rows.Count, 1),#rangeq区域中最后一行的第一个Range
            LookIn=constants.xlValues,
            SearchOrder=constants.xlByRows,
            SearchDirection=constants.xlPrevious,
            MatchCase=False)
        return lastrange
    @staticmethod
    def searchFirstRangeAppeared(myRange,searchString):
        '''这里的myRange是一个区域'''
        firstrange = myRange.Find(
            What=searchString,
            After=myRange.Cells(1, 1),#rangeq区域中第一行的第一个Range
            LookIn=constants.xlValues,
            SearchOrder=constants.xlByRows,
            SearchDirection=constants.xlNext,
            MatchCase=False)
        return firstrange
    @staticmethod
    def insertEntireRowBelowRange(mysheet,myRange):
        mysheet.Cells(myRange.Row+1,1).EntireRow.Insert()
    @staticmethod
    def insertEntireRowAboveRange(mysheet,myRange):
        mysheet.Cells(myRange.Row,1).EntireRow.Insert()
    @staticmethod
    def getWorkbookActiveSheetName(myWorkbook):
        sheetName= myWorkbook.ActiveSheet.Name
        return sheetName
    @staticmethod
    def getWorkbookActiveSheet(myWorkbook):
        mysheet=myWorkbook.ActiveSheet
        return mysheet
    @staticmethod
    def rotateRangeValue(myRange,degree=90):
        myRange.Orientation=degree
    @staticmethod
    def activateSheet(myWorkbook,myWorksheet):
        myWorkbook.Activate()
        myWorksheet.Activate()
    @staticmethod
    def isEmpty(cell):
        if  cell.Value==None:
            return True
        else:
            return False
    @staticmethod
    def createWorkbook(excelapp):
        workbook=excelapp.Workbooks.Add()
        return workbook
    
    @staticmethod
    def set_range_font_bold(myRange):
        myRange.Font.Bold=True
    
    @staticmethod
    def set_Column_width(myRange,columnWidth = 12):
        myRange.EntireColumn.ColumnWidth=columnWidth
        # sheet.Range("a1").EntireColumn.ColumnWidth =12.14
        
    
    @staticmethod
    def saveAs(workbook,full_path):
        workbook.SaveAs(full_path)
    @staticmethod
    def close(workbook):
        workbook.Close()
    @staticmethod
    def autoFit(ranges):
        ranges.AutoFit()
    @staticmethod
    def autoFilterByUsedRange(worksheet):
        worksheet.UsedRange.Columns.AutoFilter(1)
    @staticmethod
    def existStrikethrough_Bool(myrange):
        if myrange.Font.Strikethrough:
            return True
        else:
            return False
    @staticmethod
    def getRangeCommentText(myRange):
        comments= myRange.Comment.Text()
        return comments
    @staticmethod
    def colorRGB2ExcelColor(r,g,b):
        colorResult=r+(g<<8)+(b<<16)
        return colorResult
    @staticmethod
    def excelColorToRGB_ListStringValue(excelColor) -> list :
        '''
        example:
        color=int(temprange.Interior.Color)
        excelColorToRGB(color)
        '''
        red =str(excelColor & 255)
        green =str( (excelColor >> 8) & 255 )
        blue =str( (excelColor >> 16) & 255)
        return [red,green,blue]
    @staticmethod
    def set_displayAlert(excelapp, bool_TF):
        """set_displayAlert(excelapp, true / false)"""
        excelapp.DisplayAlerts=bool_TF
    @staticmethod
    def set_alignmentCenter(range):
        range.HorizontalAlignment=-4108
    @staticmethod
    def mergeRange(myRange):
        myRange.Merge()
    
      
    pass

'''list'''
class my_list(object):
    def filter_list_more_than_number(mylist, mynumber):
        return list(filter(lambda x: x >= mynumber, mylist))

    def filter_list_less_than_number(mylist, mynumber):
        return list(filter(lambda x: x <= mynumber, mylist))

    def get_list_obj_name_list(mylist):
        return [x.name for x in mylist]

    def remove_DuplcatedString_InList(mylist):
        templist = []
        for word in mylist:
            if word not in templist:
                templist.append(word)
        return templist


"""color transfer"""
class my_color:
    @staticmethod
    def colorRGB2ExcelColor(r,g,b):
        colorResult=r+(g<<8)+(b<<16)
        return colorResult
    
    
    
    
    

if __name__ == "__main__":

    print("test")