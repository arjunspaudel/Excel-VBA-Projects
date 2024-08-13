Attribute VB_Name = "Module6"
Sub AddSheets()
'Updateby Extendoffice
    Dim xRg As Excel.Range
    Dim wSh As Excel.Worksheet
    Dim wBk As Excel.Workbook
    Set wSh = ActiveSheet
    Set wBk = ActiveWorkbook
    Application.ScreenUpdating = False
    For Each xRg In wSh.Range("A1:A14")
        With wBk
            .Sheets.Add After:=.Sheets(.Sheets.Count)
            On Error Resume Next
            ActiveSheet.Name = xRg.Value
            If Err.Number = 1004 Then
              Debug.Print xRg.Value & " already used as a sheet name"
            End If
            On Error GoTo 0
        End With
    Next xRg
    Application.ScreenUpdating = True
End Sub

