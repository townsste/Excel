VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Search a specific cell from workbooks in the specified directory
Sub Search()
    Dim sPath As String, sName As String, cellLoc As String
    Dim openedWB As Workbook, currSheet As Worksheet
    
    Set currSheet = ActiveSheet 'Set current Worksheet to Search Tool WB
    rowCounter = 1 'The row that will be written to
    
    clearPrevious   'Clear anything from previous run
    
    'Set Titles
    currSheet.Cells(rowCounter, "A") = "Workbook Name"
    currSheet.Cells(rowCounter, "B") = "Cell Data"

    sPath = Worksheets("Data").Range("E2") + "\"    'hold the location path
    sName = Dir(sPath & "*.xlsx") 'workbook name from current directory
    cellLoc = Worksheets("Data").Range("E4") 'The cell location this will grab information from in workbook
    
    Do While sName <> ""
        rowCounter = rowCounter + 1   'Move to next row
        Set openedWB = Workbooks.Open(sPath & sName)    'Open Workbook with dir and name
        currSheet.Cells(rowCounter, "A") = openedWB.Name  'Save to row # in Column A
        currSheet.Cells(rowCounter, "B") = openedWB.Worksheets(1).Range(cellLoc) 'Save to row # in Column B
        openedWB.Close SaveChanges:=False   'Close opened workbook wihtout saving changes
        sName = Dir()   'Get next workbook name
    Loop
End Sub

'Function to clear existing data
Private Sub clearPrevious()
    Range("A:B").ClearContents     'Clear A and B Columns
End Sub
