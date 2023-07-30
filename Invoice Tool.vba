Attribute VB_Name = "Module1"
Sub NextInvoice()
    Range("D2").Value = Range("D2").Value + 1
    Range("A4:F34,B35,B36,B37,H7:P7,K6:Q6,S6,I8:K8,O8:S8,O9:S9,I9:K9,H13:S34,H37:O37,M37:O37").ClearContents
End Sub


Sub SaveInvWithNewName()
    Dim NewFN As Variant
    ' Copy Invoice to a new workbook
    ActiveSheet.Copy
    NewFN = "Inv" & Range("D2") & Range("K6").Value & ".xlsx"
    ActiveWorkbook.SaveAs CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\" & NewFN, FileFormat:=xlOpenXMLWorkbook
    ActiveWorkbook.Close
    NextInvoice
End Sub

