'Time Tracking Database
'Developed by Stephen Townsend
'All rights reserved

'PLEASE NOTE:  The reference of override is used to change the employees defaulted Title.

Public useOverride As Boolean
Public disabledOverride As Integer
Public overrideStatus As Boolean
Public overrideNumber As Integer 'Used to hold the override number for the employee.  Currently set to use 1-4
Public defaultName As String 'Holds the current employees name
Public defaultStaffNumber As Integer 'Holds the current selected employee staff #
Public defaultTitle As String 'Holds the current selected employee Title #

''''''''''''''''''''''''''ARRAYS'''''''''''''''''
    'SEARCH BY
Public Function empSearchArr() As Variant
    empSearchArr = Worksheets("Summary").Range("A12:A112")
End Function

    'SAP NAME
Public Function empArr() As Variant
    empArr = Worksheets("Summary").Range("B12:B112")
End Function

    'STAFF #
Public Function empStaffArr() As Variant
    empStaffArr = Worksheets("Summary").Range("C12:C112")
End Function

    'Title
Public Function empTitleArr() As Variant
    empTitleArr = Worksheets("Summary").Range("D12:D112")
End Function

    'Override
Public Function empOverrideArr() As Variant
    empOverrideArr = Worksheets("Summary").Range("E12:E112")
End Function

''''''''''''''''''''''''''END OF ARRAYS'''''''''''''''''


''''''''''''''''''''''''''MAIN'''''''''''''''''
'EMPLOYEE BOX CHNAGE
    'When the Employee box changes this will check the name and use that to determine the Title number to display in the statusBox
Private Sub employeeBox_Change()
    useOverride = True
    employeeFound = False 'Used to exit nested For
    
    If useOverride = False And disabledOverride = 0 Then
        disableOverrideCompletely
        disabledOverride = 1
    End If
    
    'CHANGED EMPLOYEE - RESET SOME INFORMATION
        resets
'''GET EMPLOYEE'''
    For R = 1 To UBound(empSearchArr, 1) 'First array dimension is rows.
    For C = 1 To UBound(empSearchArr, 2) 'Second array dimension is columns.
        If empSearchArr(R, C) = employeeBox.Text Or empArr(R, C) = employeeBox.Text Then 'Checks if the search Array or emp Array equal employeeBox text
            If useOverride = True Then
                If empOverrideArr(R, C) = "1" Or empOverrideArr(R, C) = "2" Or empOverrideArr(R, C) = "3" Or empOverrideArr(R, C) = "4" Then
                override = empOverrideArr(R, C)
                'Enable Override Status
                    overrideStatus = True
         
            'NOTE: This can be used to specify if an override is availiable based
            'on the current override.  For example:  If someone is override 1
            'then the lesser overrides may not apply.
                overrideEnableOrDisable (override)
                End If
            End If
'''END OF GET EMPLOYEE'''
            
'''SET CAPTIONS'''
        'If NO Title
            If empTitleArr(R, C) = "" Then
                Me.statusBox.Caption = ""   'Set caption to blank

        'If Overide
            ElseIf overrideStatus = True Then
                    overrideOverride  'Get new caption for operator status
                'If no Radio has been selected yet.  Get default caption
                If Me.statusBox.Caption = "" Then
                    Me.statusBox.Caption = empTitleArr(R, C) 'Default to Title number
                End If
                
        'Everyone Else
            Else
                Me.statusBox.Caption = empTitleArr(R, C) 'Set the caption to Title number
            End If
'''END OF SET CAPTIONS'''

'''INFORMATION STORAGE'''
            'Store Name, Title Number, and Staff Number for current selected employee
            defaultName = empArr(R, C)
            defaultTitle = empTitleArr(R, C)
            defaultStaffNumber = empStaffArr(R, C)
'''END OF INFORMATION STORAGE'''
            'Found Employee
            employeeFound = True     'Exit Loop Early
            Exit For        'Exit First Loop
        End If
    Next C
    If employeeFound Then Exit For   'Exit Second Loop
    Next R
End Sub

''''''''''''''''''''''''''END OF MAIN'''''''''''''''''


'HELPER FUNCTIONS BELOW

''''''''''''''''''''''''''BUTTONS'''''''''''''''''

'SAVE BUTTON
    'This is for the current employee button to clear specific values
Private Sub addCurrentForm_Click()
    If overrideStatus = True And Me.op1Box = False And Me.op2Box = False And Me.op3Box = False And Me.op4Box = False Then
        'If Me.op1Box = False Or Me.op2Box = False Or Me.op3Box = False Or Me.op4Box = False Then
        errorOverrideHandle
    Else
    'Call to add values to Database Tab
       setValues
        
    'Clear input controls.
        Me.equipoverride = False
        'Me.woBox.Value = ""
        Me.timeTypeBox.Value = ""
        Me.timeQtyBox.Value = ""
        
    'Call Refresh
        pivotRefresh
    End If
End Sub


'SAVE & NEXT BUTTON
    'This is for the next employee button to clear specific values
Private Sub nextEmployeeForm_Click()
    If overrideStatus = True And Me.op1Box = False And Me.op2Box = False And Me.op3Box = False And Me.op4Box = False Then
        'If Me.op1Box = False Or Me.op2Box = False Or Me.op3Box = False Or Me.op4Box = False Then
            errorOverrideHandle
    Else
    'Call to add values to Database Tab
        setValues
    
    'Clear input controls.
        Me.override1Box = False
        Me.override2Box = False
        Me.override3Box = False
        Me.override4Box = False
        Me.equipoverride = False
        Me.woBox.Value = ""
        Me.dateBox.Value = ""
        Me.timeTypeBox.Value = ""
        Me.timeQtyBox.Value = ""
        Me.employeeBox.Value = ""
    
    'Call Refresh
        pivotRefresh
    End If
End Sub


'QUIT BUTTON
    'Close UserForm.
Private Sub closeForm_Click()
'Call Refresh
    pivotRefresh
    
'Close Form Command
    Unload Me
End Sub

''''''''''''''''''''''''''END OF BUTTONS'''''''''''''''''


''''''''''''''''''''''''''OVERRIDE FUNCTIONS'''''''''''''''''

Public Sub disableOverrideCompletely()
    'Turn Off
    Me.override1Box.Enabled = False
    Me.override2Box.Enabled = False
    Me.override3Box.Enabled = False
    Me.override4Box.Enabled = False
    
    'Hide the Frame
    Me.Frame1.Visible = False
    
End Sub

    'OVERRIDE STATUS
Private Sub override1Box_Click()
    overrideOverride 'Call override function below
End Sub

Private Sub override2Box_Click()
    overrideOverride 'Call override function below
End Sub

Private Sub override3Box_Click()
    overrideOverride 'Call override function below
End Sub

Private Sub override4Box_Click()
    overrideOverride 'Call override function below
End Sub

    'TOGGLE override STATUS
'This is an exception and will change the job Title number to Operator Status.
Private Sub overrideOverride()
    If override1Box = True Then
        overrideNumber = 1
        Me.statusBox.Caption = Me.override1Box.Caption
    ElseIf override2Box = True Then
        overrideNumber = 2
        Me.statusBox.Caption = Me.override2Box.Caption
    ElseIf override3Box = True Then
        overrideNumber = 3
        Me.statusBox.Caption = Me.override3Box.Caption
    ElseIf override4Box = True Then
        overrideNumber = 4
        Me.statusBox.Caption = Me.override4Box.Caption
    End If
End Sub

'Function to control the override raido boxes.  Enable or disable as needed
Public Sub overrideEnableOrDisable(override)
'Enable override status
    Select Case override
        Case 1
            Me.override1Box.Enabled = True
            'Me.override2Box.Enabled = True
            'Me.override3Box.Enabled = True
            'Me.override4Box.Enabled = True
        Case 2
            Me.override1Box.Enabled = True
            Me.override2Box.Enabled = True
            'Me.overridep3Box.Enabled = True
            'Me.override4Box.Enabled = True
        Case 3
            Me.override1Box.Enabled = True
            Me.override2Box.Enabled = True
            Me.override3Box.Enabled = True
            'Me.override4Box.Enabled = True
        Case 4
            Me.override1Box.Enabled = True
            Me.override2Box.Enabled = True
            Me.override3Box.Enabled = True
            Me.override4Box.Enabled = True
    End Select
End Sub

''''''''''''''''''''''''''END OF OVERRIDE FUNCTIONS'''''''''''''''''


'Reset data due to new employee being selected
Public Sub resets()
'RESETS
'Clear current caption to blank
        Me.statusBox.Caption = ""
    'Reset the Operator Radios
        Me.override1Box = False
        Me.override2Box = False
        Me.override3Box = False
        Me.override4Box = False
    'Disable Operator Status
        overrideStatus = False
        overrideNumber = 0
    'Disable Operator Radios
        Me.override1Box.Enabled = False
        Me.override2Box.Enabled = False
        Me.override3Box.Enabled = False
        Me.override4Box.Enabled = False
'END OF RESETS
End Sub



'OUTPUT VALUES TO DATABASE TAB
    'This is used to copy input values for the form to the Database Tab
Public Sub setValues()
    Dim lRow As String
    Dim ws As Worksheet
    Set ws = Worksheets("Database")
    lRow = ws.Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row
    With ws

    'Place data in specified location
        'Column A
        .Cells(lRow, 1).Value = Me.woBox.Value '+ "0001"  'Optional: Append to end of WO
        'Column B
        .Cells(lRow, 2).Value = defaultStaffNumber 'Output employee staff number
        'Column C
        .Cells(lRow, 3).Value = Me.statusBox.Caption 'Output what is stored in the caption
        'Column D
        .Cells(lRow, 4).Value = Me.dateBox.Value     'Output Date
        'Column E
        .Cells(lRow, 5).Value = Me.timeTypeBox.Value 'Output ST/OT/DT
        'Column F
        .Cells(lRow, 6).Value = Me.timeQtyBox.Value
        'Column G
        .Cells(lRow, 7).Value = defaultName
        
        '9 & 10 will allow for a slicer to sort each entry correctly by the week of date.
        'PLEASE NOTE: This is based on the date being in Column D (#4).  If date location is changed
        'then replace "D" with the new column letter.
        .Cells(lRow, 9).Value = "=WEEKDAY(D" + lRow + ",2)"
        
        'The fomrula below will calculate the end of the week date based on the date in
            'its row.  Set up as Mon to Sunday.
        'Formula: =TEXT(IF(M{Cell#} = 1,D{Cell#}+6, IF(M{Cell#} = 2, D{Cell#}+5, IF(M{Cell#} = 3,D{Cell#}+4, IF(M{Cell#}= 4, D{Cell#}+3, IF(M{Cell#} = 5, D{Cell#}+2, IF(M{Cell#}= 6, D{Cell#}+1, IF(M{Cell#}=7,D{Cell#}+0))))))), "mm/dd")
                '{Cell#} is equal to the lRow string var
        .Cells(lRow, 10).Value = "=TEXT(IF(M" + lRow + "= 1,D" + lRow + "+6, IF(M" + lRow + "= 2, D" + lRow + "+5, IF(M" + lRow + "= 3,D" + lRow + "+4, IF(M" + lRow + "= 4, D" + lRow + "+3, IF(M" + lRow + "= 5, D" + lRow + "+2, IF(M" + lRow + "= 6, D" + lRow + "+1, IF(M" + lRow + "=7,D" + lRow + "+0)))))))," + Chr(34) + "mm/dd" + Chr(34) + ")"
        
        'This is used to check and output an override value if on exists
        If overrideStatus = True Then
            .Cells(lRow, 8).Value = overrideNumber
        End If
    End With
End Sub


'This is used to Refresh the Pivot Table in Summary Tab
Public Sub pivotRefresh()
'Summary pivottable for auto refresh
    Dim pt As PivotTable
    Set pt = Worksheets("Summary").PivotTables("SummaryTable")
    
    pt.RefreshTable
End Sub


'Inform the user to select Override Status
Public Sub errorOverrideHandle()
    MsgBox "Please Select an Operator Status"
End Sub







''''''''''''''''''''''''''REMOVED FUNCTIONS'''''''''''''''''

'No Longer Need This Sub.  Using dropdown RowSource instead of code
'This is done by using define name.  Define the range by name and then use that name
'under the RowSource option for the dropdown.
'----------------------------------------------------------------------------------
'EMPLOYEE BOX
    'To Populate the Dropdown for Employee List
'Private Sub employeeBox_DropButtonClick()
'    Dim R As Long
'    Dim C As Long
'    If empDropCounter = 0 Then        'Run this section once to prevent duplicates
'        For R = 1 To UBound(empSearchArr, 1) 'First array dimension is rows.
'        For C = 1 To UBound(empSearchArr, 2) 'Second array dimension is columns.
'            Me.employeeBox.AddTitle empSearchArr(R, C)
'        Next C
'        Next R
'    End If
'    empDropCounter = 1
'End Sub
'----------------------------------------------------------------------------------



'No Longer Need This Sub.  Using dropdown RowSource instead of code
'This is done by using define name.  Define the range by name and then use that name
'under the RowSource option for the dropdown.
'----------------------------------------------------------------------------------
'TIME TYPE BOX
    'Dropdown for Time Type List
'Private Sub timeTypeBox_DropButtonClick()
'    If Me.timeTypeBox.ListCount = 0 Then  'Run this section once to prevent duplicates
'       Me.timeTypeBox.AddTitle "Standard"
'       Me.timeTypeBox.AddTitle "Overtime"
'       Me.timeTypeBox.AddTitle "Doubletime"
'    End If
'End Sub
'----------------------------------------------------------------------------------

''''''''''''''''''''''''''END OF REMOVED FUNCTIONS'''''''''''''''''
