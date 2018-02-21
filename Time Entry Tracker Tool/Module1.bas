Attribute VB_Name = "Module1"
Sub ShowUF()
    'Display UserForm.
    Add_to_Table.Show Modal
End Sub

Sub ClearMySlicers()
    Dim Slcr As SlicerCache
    For Each Slcr In ActiveWorkbook.SlicerCaches
    Slcr.ClearManualFilter
    Next
End Sub

Sub ManualRefresh()
    Dim pt As PivotTable
    Set pt = ActiveSheet.PivotTables("SummaryTable")
    pt.RefreshTable
End Sub