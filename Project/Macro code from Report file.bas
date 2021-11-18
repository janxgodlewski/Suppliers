Attribute VB_Name = "Module1"
Public path As String
Public maxColumns As Integer
Public maxRows As Integer

Public name As String
Public suppliers As Range
Public month As Integer

'Calls function filterSheet for every .xlsx file
Sub generate()

Application.DisplayAlerts = False
Application.ScreenUpdating = False

path = ThisWorkbook.sheets("Narzêdzie").[Q3]
maxRows = 150
maxColumns = 240

Set suppliers = ThisWorkbook.sheets("Narzêdzie").Range(ThisWorkbook.sheets("Narzêdzie").Cells(5, 1), ThisWorkbook.sheets("Narzêdzie").Cells(ThisWorkbook.sheets("Narzêdzie").Cells(5, 1).End(xlDown).Row, 1))
month = CInt(ThisWorkbook.sheets("Narzêdzie").[B3])


Dim r As Range, cell As Range
Set r = ThisWorkbook.sheets("Narzêdzie").Range("C2:O2")

Dim sheets() As String
Dim sheetsCount As Integer

For Each cell In r

    If (StrComp(cell.Value, "PRAWDA")) = 0 Or (StrComp(cell.Value, "Prawda")) = 0 Then

        sheetsCount = sheetsCount + 1

    End If

Next cell

ReDim sheets(sheetsCount)

Dim i As Integer: i = 0
For Each cell In r

    If (StrComp(cell, "True")) = 0 Then

        name = cell.Offset(-1, 0).Value
        
        Call filterSheet

        sheets(i) = name
    
    End If
    
Next cell

Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub

'Defines which rows/columns should be hidden in every .xlsx file
Public Function filterSheet()

Dim wb As Workbook
Set wb = Workbooks.Open(path & "\" & name)

Dim column
Dim codeColumn, codeRow, headerRow, sheetNumber As Integer

If (StrComp(name, "BUG.xlsx")) = 0 Then

sheetNumber = 1 'sheet number
wb.sheets(sheetNumber).Activate

headerRow = 4 'header row of the table
codeRow = 5 'number of row from which starts to filter suppliers
codeColumn = 1 'column with supplier numbers
column = Array(1, 2, 2 + month, 11, 11 + month, 20, 20 + month, 29, 30, 31, 32) 'columns which should stay visible

Call hide(wb, sheetNumber, headerRow, codeRow, codeColumn, column) 'hide rows/columns

End If

'-----------------------------------------------------------------------------------------------

If (StrComp(name, "CV_SUG.xlsx")) = 0 Then

sheetNumber = 1 'sheet number
wb.sheets(sheetNumber).Activate

headerRow = 1 'header row of the table
codeRow = 2 'number of row from which starts to filter suppliers
codeColumn = 2 'column with supplier numbers
column = Array(2, 2 + month, 15) 'columns which should stay visible

Call hide(wb, sheetNumber, headerRow, codeRow, codeColumn, column) 'hide rows/columns

sheetNumber = 2 'sheet number
wb.sheets(sheetNumber).Activate

headerRow = 1 'header row of the table
codeRow = 2 'number of row from which starts to filter suppliers
codeColumn = 2 'column with supplier numbers
column = Array(2, 2 + month, 15) 'columns which should stay visible

Call hide(wb, sheetNumber, headerRow, codeRow, codeColumn, column) 'hide rows/columns

End If

'-----------------------------------------------------------------------------------------------

If (StrComp(name, "D2D.xlsx")) = 0 Then

sheetNumber = 1 'sheet number
wb.sheets(sheetNumber).Activate

headerRow = 4 'header row of the table
codeRow = 5 'number of row from which starts to filter suppliers
codeColumn = 1 'column with supplier numbers
column = Array(1, 1 + month, 10) 'columns which should stay visible

Call hide(wb, sheetNumber, headerRow, codeRow, codeColumn, column) 'hide rows/columns

End If

'-----------------------------------------------------------------------------------------------

If (StrComp(name, "EXS MIX.xlsx")) = 0 Then

sheetNumber = 1 'sheet number
wb.sheets(sheetNumber).Activate

headerRow = 2 'header row of the table
codeRow = 5 'number of row from which starts to filter suppliers
codeColumn = 1 'column with supplier numbers

If month <= 3 Then
column = Array(1, 2, 2 + (month * 2) - 1, 2 + (month * 2), 9, 10, 35, 36) 'columns which should stay visible
ElseIf month <= 6 Then
column = Array(1, 2, 10 + ((month - 3) * 2) - 1, 10 + ((month - 3) * 2), 17, 18, 35, 36) 'columns which should stay visible
ElseIf month <= 9 Then
column = Array(1, 2, 18 + ((month - 6) * 2) - 1, 18 + ((month - 6) * 2), 25, 26, 35, 36) 'columns which should stay visible
ElseIf month <= 12 Then
column = Array(1, 2, 26 + ((month - 9) * 2) - 1, 25 + ((month - 9) * 2), 33, 34, 35, 36)
End If

Call hide(wb, sheetNumber, headerRow, codeRow, codeColumn, column) 'hide rows/columns

End If

'-----------------------------------------------------------------------------------------------

If (StrComp(name, "EXS-CV.xlsx")) = 0 Then

sheetNumber = 1 'sheet number
wb.sheets(sheetNumber).Activate

headerRow = 2 'header row of the table
codeRow = 5 'number of row from which starts to filter suppliers
codeColumn = 2 'column with supplier numbers

If month <= 3 Then
column = Array(2, 3, 3 + (month * 2) - 1, 3 + (month * 2), 10, 11, 36, 37) 'columns which should stay visible
ElseIf month <= 6 Then
column = Array(2, 3, 11 + ((month - 3) * 2) - 1, 11 + ((month - 3) * 2), 18, 19, 36, 37) 'columns which should stay visible
ElseIf month <= 9 Then
column = Array(2, 3, 19 + ((month - 6) * 2) - 1, 19 + ((month - 6) * 2), 26, 27, 36, 37) 'columns which should stay visible
ElseIf month <= 12 Then
column = Array(2, 3, 27 + ((month - 9) * 2) - 1, 27 + ((month - 9) * 2), 34, 35, 36, 37)
End If

Call hide(wb, sheetNumber, headerRow, codeRow, codeColumn, column) 'hide rows/columns

sheetNumber = 1 + month 'sheet number
wb.sheets(sheetNumber).Activate

headerRow = 2 'header row of the table
codeRow = 3 'number of row from which starts to filter suppliers
codeColumn = 2 'column with supplier numbers
column = Array(2, 3, 4, 5, 6, 7) 'columns which should stay visible

Call hide(wb, sheetNumber, headerRow, codeRow, codeColumn, column) 'hide rows/columns

End If

'-----------------------------------------------------------------------------------------------

If (StrComp(name, "BUG.xlsx")) = 0 Then

sheetNumber = 1 'sheet number
wb.sheets(sheetNumber).Activate

headerRow = 4 'header row of the table
codeRow = 5 'number of row from which starts to filter suppliers
codeColumn = 1 'column with supplier numbers
column = Array(1, 2, 2 + month, 11, 11 + month, 20, 20 + month, 29, 30, 31, 32) 'columns which should stay visible

Call hide(wb, sheetNumber, headerRow, codeRow, codeColumn, column) 'hide rows/columns

End If

'-----------------------------------------------------------------------------------------------

If (StrComp(name, "VC.xlsx")) = 0 Then

sheetNumber = month 'sheet number
wb.sheets(sheetNumber).Activate

headerRow = 4 'header row of the table
codeRow = 6 'number of row from which starts to filter suppliers
codeColumn = 1 'column with supplier numbers
column = Array(1, 2, 3, 4, 5, 6, 7, 8) 'columns which should stay visible

Call hide(wb, sheetNumber, headerRow, codeRow, codeColumn, column) 'hide rows/columns

End If

'-----------------------------------------------------------------------------------------------

If (StrComp(name, "Stock Parameters.xlsx")) = 0 Then

sheetNumber = 1 'sheet number
wb.sheets(sheetNumber).Activate

headerRow = 1 'header row of the table
codeRow = 2 'number of row from which starts to filter suppliers
codeColumn = 1 'column with supplier numbers
column = Array(1, 1 + month) 'columns which should stay visible

Call hide(wb, sheetNumber, headerRow, codeRow, codeColumn, column) 'hide rows/columns


sheetNumber = 2 'sheet number
wb.sheets(sheetNumber).Activate

headerRow = 1 'header row of the table
codeRow = 3 'number of row from which starts to filter suppliers
codeColumn = 1 'column with supplier numbers
column = Array(1, 1 + month) 'columns which should stay visible

Call hide(wb, sheetNumber, headerRow, codeRow, codeColumn, column) 'hide rows/columns

End If

'-----------------------------------------------------------------------------------------------

If (StrComp(name, "Sell_Out.xls")) = 0 Then

sheetNumber = 1 'sheet number
wb.sheets(sheetNumber).Activate

headerRow = 1 'header row of the table
codeRow = 6 'number of row from which starts to filter suppliers
codeColumn = 3 'column with supplier numbers
column = Array(3, 185, 186, 187, 188, 189, 190, 191, 192, 193, 194, 195, 196, 197, 198, 199, 200, 201, 202, 203, 204, 205, 206, 207, 208, 209, 210) 'columns which should stay visible

Call hide(wb, sheetNumber, headerRow, codeRow, codeColumn, column) 'hide rows/columns

End If

'-----------------------------------------------------------------------------------------------

If (StrComp(name, "Sell_In.xls")) = 0 Then

sheetNumber = 2 'sheet number
wb.sheets(sheetNumber).Activate

headerRow = 1 'header row of the table
codeRow = 3 'number of row from which starts to filter suppliers
codeColumn = 1 'column with supplier numbers
column = Array(1, 2, month + 2) 'columns which should stay visible

Call hide(wb, sheetNumber, headerRow, codeRow, codeColumn, column) 'hide rows/columns

End If

'-----------------------------------------------------------------------------------------------

If (StrComp(name, "SAWA.xlsx")) = 0 Then

sheetNumber = 1 'sheet number
wb.sheets(sheetNumber).Activate

headerRow = 1 'header row of the table
codeRow = 2 'number of row from which starts to filter suppliers
codeColumn = 1 'column with supplier numbers
column = Array(1, 2, month + 2) 'columns which should stay visible

Call hide(wb, sheetNumber, headerRow, codeRow, codeColumn, column) 'hide rows/columns

End If

'-----------------------------------------------------------------------------------------------

If (StrComp(name, "SARA.xlsx")) = 0 Then

sheetNumber = 1 'sheet number
wb.sheets(sheetNumber).Activate

headerRow = 1 'header row of the table
codeRow = 3 'number of row from which starts to filter suppliers
codeColumn = 2 'column with supplier numbers
column = Array(1, 2, month + 2) 'columns which should stay visible

Call hide(wb, sheetNumber, headerRow, codeRow, codeColumn, column) 'hide rows/columns

sheetNumber = 2 'sheet number
wb.sheets(sheetNumber).Activate

headerRow = 1 'header row of the table
codeRow = 3 'number of row from which starts to filter suppliers
codeColumn = 2 'column with supplier numbers
column = Array(1, 2, month + 2) 'columns which should stay visible

Call hide(wb, sheetNumber, headerRow, codeRow, codeColumn, column) 'hide rows/columns

End If

'-----------------------------------------------------------------------------------------------

If (StrComp(name, "OSB.xlsx")) = 0 Then

sheetNumber = 1 'sheet number
wb.sheets(sheetNumber).Activate

headerRow = 1 'header row of the table
codeRow = 2 'number of row from which starts to filter suppliers
codeColumn = 2 'column with supplier numbers
column = Array(1, 2, month + 2) 'columns which should stay visible

Call hide(wb, sheetNumber, headerRow, codeRow, codeColumn, column) 'hide rows/columns

End If

'-----------------------------------------------------------------------------------------------

If (StrComp(name, "Hurt.xlsx")) = 0 Then

sheetNumber = month 'sheet number
wb.sheets(sheetNumber).Activate

headerRow = 1 'header row of the table
codeRow = 4 'number of row from which starts to filter suppliers
codeColumn = 2 'column with supplier numbers
column = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11) 'columns which should stay visible

Call hide(wb, sheetNumber, headerRow, codeRow, codeColumn, column) 'hide rows/columns

End If

'-----------------------------------------------------------------------------------------------

wb.Close SaveChanges:=False

End Function

'hides rows/columns in choosen file and copies the table to main excel file called "Tool"
Public Function hide(ByVal wb As Workbook, ByVal sheetNumber As Integer, _
ByVal headerRow As Integer, ByVal codeRow As Integer, ByVal codeColumn As Integer, ByVal column As Variant)

wb.sheets(sheetNumber).Cells.UnMerge
wb.sheets(sheetNumber).Rows.EntireRow.Hidden = False
wb.sheets(sheetNumber).Columns.EntireColumn.Hidden = False

wb.sheets(sheetNumber).Cells(headerRow, codeColumn).Select

Dim cx, cy, lastX, lastY As Integer
lastX = headerRow

Dim contains As Boolean
'hide columns
For cy = 1 To maxColumns

    For c = 0 To UBound(column)
    
        If (cy = column(c)) Then
            contains = True
            Exit For
        Else
            contains = False
        End If

    Next c

    If (contains <> True) Then
        wb.sheets(sheetNumber).Columns(cy).EntireColumn.Hidden = True
    Else
        lastY = cy
    End If
    
Next cy

'hide rows
For cx = codeRow To maxRows

    For Each cell In suppliers
    
        If (InStr(wb.sheets(sheetNumber).Cells(cx, codeColumn).Value, cell.Value)) = 1 Then
         contains = True
    Exit For
    Else
    contains = False
    End If
           
            
    Next cell
    
     If (contains <> True) Then
         wb.sheets(sheetNumber).Cells(cx, codeColumn).EntireRow.Hidden = True
     Else
        lastX = cx
    End If

Next cx

wb.sheets(sheetNumber).Range(Cells(headerRow, 1), Cells(codeRow - 1, 1)).EntireRow.Hidden = False

wb.sheets(sheetNumber).Cells(headerRow, 1).Select

wb.sheets(sheetNumber).Range(wb.sheets(sheetNumber).Cells(headerRow, 1), wb.sheets(sheetNumber).Cells(lastX, lastY)).SpecialCells(xlCellTypeVisible).Copy

Dim exists As Boolean

For i = 1 To ThisWorkbook.Worksheets.Count
    If ThisWorkbook.Worksheets(i).name = name Then
        exists = True
    End If
Next i

If Not exists Then
    ThisWorkbook.sheets.Add.name = name
    ThisWorkbook.sheets(name).Activate
    ThisWorkbook.ActiveSheet.Range("A1").Select
    
    ActiveCell.PasteSpecial xlPasteValues
    ActiveCell.PasteSpecial xlPasteFormats
Else
    ThisWorkbook.sheets(name).Activate
    ThisWorkbook.sheets(name).Cells(ActiveCell.Row + Selection.Rows.Count + 1, 1).Select

    ActiveCell.PasteSpecial xlPasteValues
    ActiveCell.PasteSpecial xlPasteFormats
    

End If

wb.sheets(sheetNumber).Rows.EntireRow.Hidden = False
wb.sheets(sheetNumber).Columns.EntireColumn.Hidden = False

End Sub


