Attribute VB_Name = "ConcatenateCellsIfSameValues"
Sub COEDMConcatenateCellsIfSameValues()
    Dim xCol As New Collection
    Dim xSrc As Variant
    Dim xRes() As Variant
    Dim i As Long
    Dim J As Long
    Dim xRg As Range
    xSrc = Range("A1", Cells(Rows.Count, "A").End(xlUp)).Resize(, 2)
    Set xRg = Range("D1")
    On Error Resume Next
    For i = 2 To UBound(xSrc)
        xCol.Add xSrc(i, 1), TypeName(xSrc(i, 1)) & CStr(xSrc(i, 1))
    Next i
    On Error GoTo 0
    ReDim xRes(1 To xCol.Count + 1, 1 To 2)
    xRes(1, 1) = "Doc Nbr"
    xRes(1, 2) = "KKS after merge"
    For i = 1 To xCol.Count
        xRes(i + 1, 1) = xCol(i)
        For J = 2 To UBound(xSrc)
            If xSrc(J, 1) = xRes(i + 1, 1) Then
                xRes(i + 1, 2) = xRes(i + 1, 2) & ";" & vbLf & xSrc(J, 2)
            End If
        Next J
        xRes(i + 1, 2) = Mid(xRes(i + 1, 2), 3)
    Next i
    Set xRg = xRg.Resize(UBound(xRes, 1), UBound(xRes, 2))
    xRg.NumberFormat = "@"
    xRg = xRes
End Sub

