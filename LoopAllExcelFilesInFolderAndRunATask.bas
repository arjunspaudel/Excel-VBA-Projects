Attribute VB_Name = "TEST1"
Sub LoopAllExcelFilesInFolder()
'PURPOSE: To loop through all Excel files in a user specified folder and perform a set task on them
'SOURCE: www.TheSpreadsheetGuru.com

Dim WB As Workbook
Dim myPath As String
Dim myFile As String
Dim myExtension As String
Dim FldrPicker As FileDialog

'Optimize Macro Speed
  Application.ScreenUpdating = False
  Application.EnableEvents = False
  Application.Calculation = xlCalculationManual

'Retrieve Target Folder Path From User
  Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker)

    With FldrPicker
      .Title = "Select A Target Folder"
      .AllowMultiSelect = False
        If .Show <> -1 Then GoTo NextCode
        myPath = .SelectedItems(1) & "\"
    End With

'In Case of Cancel
NextCode:
  myPath = myPath
  If myPath = "" Then GoTo ResetSettings

'Target File Extension (must include wildcard "*")
  myExtension = "*.xlsx*"

'Target Path with Ending Extention
  myFile = Dir(myPath & myExtension)

'Loop through each Excel file in folder
  Do While myFile <> ""
    'Set variable equal to opened workbook
      Set WB = Workbooks.Open(Filename:=myPath & myFile)
    
    'Ensure Workbook has opened before moving on to next line of code
      DoEvents
    
    'Code to run in the File
      Columns("N:N").Select
            Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            Range("N2").Select
            ActiveCell.FormulaR1C1 = _
                "=VLOOKUP(RC[-9],'[Master_EI-Aims_export.xlsx]Sheet1'!C2:C31,30,0)"
            'Selection.AutoFill Destination:=Range("N2:N4124")
            'Range("N2:N4124").Select
            Selection.AutoFill Destination:=Range("N2:N" & Range("K" & Rows.Count).End(xlUp).Row)
            Range(Selection, Selection.End(xlDown)).Select
            Range("O2").Select
            ActiveCell.FormulaR1C1 = _
                "=IFERROR(IF(SEARCH(""*Final CAE*"",RC[-1],1),""YES""),"""")"
            Range("O2").Select
            'Selection.AutoFill Destination:=Range("O2:O4124")
            'Range("O2:O4124").Select
            Selection.AutoFill Destination:=Range("O2:O" & Range("K" & Rows.Count).End(xlUp).Row)
            Range(Selection, Selection.End(xlDown)).Select
            Range("P2").Select
            ActiveCell.FormulaR1C1 = _
                "=IFERROR(IF(SEARCH(""*NF no*"",RC[-2],1),""NO""),IF(SEARCH(""*NF yes*"",RC[-2],1),""YES""))"
            Range("P2").Select
            'Selection.AutoFill Destination:=Range("P2:P4124")
            'Range("P2:P4124").Select
            Selection.AutoFill Destination:=Range("P2:P" & Range("K" & Rows.Count).End(xlUp).Row)
            Range(Selection, Selection.End(xlDown)).Select
            Range("Q2").Select
            ActiveCell.FormulaR1C1 = _
                "=IFEROOR(IF(SEARCH(""*Clean copy*"",RC[-3],1),""YES""),IF(SEARCH(""*Redlining*"",RC[-3],1),""NO""))"
            Range("Q2").Select
            ActiveCell.FormulaR1C1 = _
                "=IFEROOR(IF(SEARCH(""*Clean copy*"",RC[-3],1),""YES""),IF(SEARCH(""*Redlining*"",RC[-3],1),""NO""))"
            Range("Q2").Select
            ActiveCell.FormulaR1C1 = _
                "=IFERROR(IF(SEARCH(""*Clean copy*"",RC[-3],1),""YES""),IF(SEARCH(""*Redlining*"",RC[-3],1),""NO""))"
            Range("Q2").Select
            'Selection.AutoFill Destination:=Range("Q2:Q4124")
            'Range("Q2:Q4124").Select
            Selection.AutoFill Destination:=Range("Q2:Q" & Range("K" & Rows.Count).End(xlUp).Row)
            Range(Selection, Selection.End(xlDown)).Select
            Range("R2").Select
            ActiveCell.FormulaR1C1 = _
                "=IFERROR(IF(SEARCH(""*Clean copy*"",RC[-4],1),""NO""),IF(SEARCH(""*Redlining*"",RC[-4],1),""YES""))"
            Range("R2").Select
            'Selection.AutoFill Destination:=Range("R2:R4124")
            'Range("R2:R4124").Select
            Selection.AutoFill Destination:=Range("R2:R" & Range("K" & Rows.Count).End(xlUp).Row)
            Range(Selection, Selection.End(xlDown)).Select
            'Columns("O:R").Select
            'Selection.Copy
            'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                ':=False, Transpose:=False
            'Range("O2").Select
            'Application.CutCopyMode = False
            Columns("O:R").Select
            Selection.Replace What:="#VALUE!", Replacement:="", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
            'Columns("N:N").Select
            'Selection.Delete Shift:=xlToLeft
            Range("N2").Select
            ActiveWorkbook.Save
    
    'Save and Close Workbook
      WB.Close SaveChanges:=True
      
    'Ensure Workbook has closed before moving on to next line of code
      DoEvents

    'Get next file name
      myFile = Dir
  Loop

'Message Box when tasks are completed
  MsgBox "Task Complete! Thanks ARJUN"

ResetSettings:
  'Reset Macro Optimization Settings
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub
