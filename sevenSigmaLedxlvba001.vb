'Houssam Frague, Sunday, January 17, 2021
'This program displays 7 sigma led digits in excel worksheet


Option Explicit             'force explicit declaration of all variables
Dim objSheet As Worksheet   'declare a Worksheet object
Dim ledRng As Range         'declare a Range object
Dim arrNumbers As Variant   'declare a variant or array

'Import and Declare Sleep function from kernel32.dll library
#If VBA7 Then ' Excel 2010 or later
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)
#Else ' Excel 2007 or earlier
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As Long)
#End If

'Clean & Prepare the active worksheet
Private Sub PrepareWorkSheet()
    With objSheet           'clean worksheet
        .Cells.ClearContents
        .Cells.ClearFormats
        .Cells.Borders.Color = RGB(255, 255, 255)
        .Cells.UseStandardHeight = True
        .Cells.UseStandardWidth = True
    End With
    With ledRng             'setup sizes
        Union(.Cells(1), .Cells(3), .Cells(7), .Cells(13)).ColumnWidth = 0.7
        Union(.Cells(1), .Cells(3), .Cells(7), .Cells(13)).RowHeight = 6
        Union(.Cells(4), .Cells(10)).ColumnWidth = 0.7
        Union(.Cells(4), .Cells(10)).RowHeight = 44
        .Cells(2).ColumnWidth = 8
        .Cells(2).RowHeight = 6
    End With
End Sub

'Clear the display area
Private Sub ClearDisplay()
ledRng.Cells.ClearFormats
ledRng.Cells.Borders.Color = RGB(255, 255, 255)
End Sub

'Display the number by changing excel ranges colors according to the segments
Private Sub DisplayChar(varCmd As Variant)
Dim i As Integer
For i = LBound(varCmd) To UBound(varCmd)
    ledRng.Cells(varCmd(i)).Interior.Color = RGB(38, 38, 38)
    ledRng.Cells(varCmd(i)).Borders.Color = RGB(38, 38, 38)
Next i
End Sub

'Start the macro
Sub Start()
    Set objSheet = ActiveSheet              'setting the active worksheet as our Worksheet object
    Set ledRng = objSheet.Range("I5:K9")    'define display area within the worksheet
    PrepareWorkSheet                        'clean worksheet
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
                        )                   'initialize 2d array with segments data for 0 to 9
    Dim i As Integer
    For i = 0 To UBound(arrNumbers)         'loop through all digits in the array and display them
        ClearDisplay                        'clear the area to display the next number
        DisplayChar arrNumbers(i)           'display the digit
        DoEvents                            'yields execution so that the os can process other events
        Sleep 1000                          'sleep for one second
    Next i
    ClearDisplay
End Sub
