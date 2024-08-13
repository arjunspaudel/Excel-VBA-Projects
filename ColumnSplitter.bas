Attribute VB_Name = "Module3"
Sub ColumnSplitter_v2()
' v2 aims to be faster by appending rows rather than inserting rows

Dim col As Integer, rw As Integer, activemacro As Boolean, lb_cr() As String, lb_cr2() As String, acell As Range

Set acell = Range(ActiveCell.Address)
Set icell = Range(acell.Address)

RowCount = acell.CurrentRegion.Rows.Count - acell.Row + 1
CheckTimer = Now()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False


' check at most one cell is active
If acell.Cells.Count > 1 Then MsgBox "Please only select one cell, the macro will work down from the selected cell"
activemacro = True

' start loop
While activemacro = True

' skip merged cells
If acell.MergeArea.Cells.Count > 1 Then
    acell.Offset(1, 0).Select
Else

' identify if the cell has a line break or carriage return
If InStr(acell.Value, vbLf) Or InStr(acell.Value, vbCr) Then

' split into an array according to line breaks and carriage returns
    lb_cr = Split(Replace(acell.Value, vbCr, vbLf), vbLf)

' remove gaps
    J = UBound(lb_cr)
    lb_cr2 = lb_cr
    ReDim lb_cr(0)
    For i = LBound(lb_cr2) To UBound(lb_cr2)
        If Len(Trim(lb_cr2(i))) > 0 Then
            If lb_cr(UBound(lb_cr)) = "" Then
                lb_cr(UBound(lb_cr)) = lb_cr2(i)
            Else
                ReDim Preserve lb_cr(UBound(lb_cr) + 1)
                lb_cr(UBound(lb_cr)) = lb_cr2(i)
            End If
        End If
    Next i

' copy the line
    acell.EntireRow.Copy

' go to the end of the column
    With acell.Offset(acell.CurrentRegion.SpecialCells(xlCellTypeLastCell).Row + 1 - acell.Row, 0)

' paste the required number of new rows
    .Resize(UBound(lb_cr)).EntireRow.PasteSpecial xlPasteAll
    Application.CutCopyMode = False

' loop through and insert the split text
        For i = LBound(lb_cr) + 1 To UBound(lb_cr)
            .Offset(i - 1, 0).Value = Trim(lb_cr(i))
        Next i

' leave the end of the column
    End With

' replace the text in the original line
    acell.Value = Trim(lb_cr(LBound(lb_cr)))

' clean lb_cr
    Erase lb_cr
    Erase lb_cr2

End If

' check if cell was empty
    If IsEmpty(acell) Then

' check if any data is below to end the macro
        If acell.End(xlDown).Row = acell.Parent.Rows.Count Then
            activemacro = False
        Else

' select the next block of values
            Set acell = acell.End(xlDown)
        End If
    Else

' select the next cell
        Set acell = acell.Offset(1, 0)
    End If
    
End If

Wend

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True

MsgBox "Processed " & RowCount & " rows in " & Format(Now - CheckTimer, "hh:mm:ss") & " hh:mm:ss with output of " & Format(acell.Row - icell.Row, "#,##0") & " rows"


End Sub
Sub columnsplitter()

End Sub
