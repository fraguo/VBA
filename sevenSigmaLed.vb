'Houssam Frague, Sunday, January 17, 2021
Option Explicit
Dim objSheet As Worksheet
Dim ledRng As Range
Dim arrNumbers As Variant
#If VBA7 Then ' Excel 2010 or later
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)
#Else ' Excel 2007 or earlier
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As Long)
#End If
Sub main()

Set objSheet = ActiveSheet
Set ledRng = objSheet.Range("I5:K9")
PrepareWorkSheet
arrNumbers = Array( _
                    Array(1, 2, 3, 4, 6, 7, 9, 10, 12, 13, 14, 15), _
                    Array(3, 6, 9, 12, 15), _
                    Array(1, 2, 3, 6, 7, 8, 9, 10, 13, 14, 15), _
                    Array(1, 2, 3, 6, 7, 8, 9, 12, 13, 14, 15), _
                    Array(1, 3, 4, 6, 7, 8, 9, 12, 15), _
                    Array(1, 2, 3, 4, 7, 8, 9, 12, 13, 14, 15), _
                    Array(1, 2, 3, 4, 7, 8, 9, 10, 12, 13, 14, 15), _
                    Array(1, 2, 3, 6, 9, 12, 15), _
                    Array(1, 2, 3, 4, 6, 7, 8, 9, 10, 12, 13, 14, 15), _
                    Array(1, 2, 3, 4, 6, 7, 8, 9, 12, 15) _
                    )
Dim i As Integer
For i = 0 To UBound(arrNumbers)
    ClearDisplay
    DisplayChar arrNumbers(i)
    DoEvents
    Sleep 1000
Next i
ClearDisplay
End Sub

Private Sub PrepareWorkSheet()
With objSheet
    .Cells.ClearContents
    .Cells.ClearFormats
    .Cells.Borders.Color = RGB(255, 255, 255)
    .Cells.UseStandardHeight = True
    .Cells.UseStandardWidth = True
End With
With ledRng
    Union(.Cells(1), .Cells(3), .Cells(7), .Cells(13)).ColumnWidth = 0.7
    Union(.Cells(1), .Cells(3), .Cells(7), .Cells(13)).RowHeight = 6
    Union(.Cells(4), .Cells(10)).ColumnWidth = 0.7
    Union(.Cells(4), .Cells(10)).RowHeight = 44
    .Cells(2).ColumnWidth = 8
    .Cells(2).RowHeight = 6
End With
End Sub

Private Sub ClearDisplay()
ledRng.Cells.ClearFormats
ledRng.Cells.Borders.Color = RGB(255, 255, 255)
End Sub

Private Sub DisplayChar(varCmd As Variant)
Dim i As Integer
For i = LBound(varCmd) To UBound(varCmd)
    ledRng.Cells(varCmd(i)).Interior.Color = RGB(38, 38, 38)
    ledRng.Cells(varCmd(i)).Borders.Color = RGB(38, 38, 38)
Next i
End Sub
