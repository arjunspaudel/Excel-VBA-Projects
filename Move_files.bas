Attribute VB_Name = "Module1"
Sub MoveFiles()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim sourceFilePath As String
    Dim destinationFolderPath As String
    Dim fso As Object
    Dim i As Long
    Dim fileName As String

    ' Set worksheet
    Set ws = ThisWorkbook.Sheets(1) ' Change this to your sheet name if it's not the first sheet

    ' Find the last row with data in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Create FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Loop through each row to move files
    For i = 2 To lastRow ' Assuming the first row is header, start from row 2
        sourceFilePath = ws.Cells(i, 1).Value
        destinationFolderPath = ws.Cells(i, 2).Value

        ' Get the file name from the source path
        fileName = fso.GetFileName(sourceFilePath)

        ' Check if the source file exists
        If fso.FileExists(sourceFilePath) Then
            ' Move the file
            fso.MoveFile sourceFilePath, fso.BuildPath(destinationFolderPath, fileName)
            ws.Cells(i, 3).Value = "Moved" ' Optionally mark the status in column C
        Else
            ws.Cells(i, 3).Value = "File Not Found" ' Mark as not found in column C
        End If
    Next i

    ' Clean up
    Set fso = Nothing

    MsgBox "File transfer completed!", vbInformation

End Sub

