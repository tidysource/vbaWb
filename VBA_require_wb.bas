Attribute VB_Name = "VBA_require_wb"
Option Explicit

'To consider
'===========
'Add checks to make sure that total cell reference
'will not be beyond excel limits
'(row>1048576 and column>16384)

'Helper functions
'================
'Get length of an array
'----------------------
Private Function arrLenght(arr As Variant) As Long
    arrLenght = UBound(arr) - LBound(arr) + 1
End Function

'Limits of excel
'---------------
'https://support.office.com/en-us/article/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3#ID0EBABAAA=2016,_2013
Private Function excelMaxRow() As Long
    excelMaxRow = 1048576
End Function
Private Function excelMaxColumn() As Integer
    excelMaxColumn = 16384
End Function

'Localisation
'------------
Private Function UItext(keyStr As String) As String
    'TO-THINK: could add an optional byval paramarray to pass extra info for text
    Dim result As New Collection
    result.Add _
                Key:="newWorksheet", _
                Item:="Vnesite ime zavika"
    result.Add _
                Key:="renameWorksheet", _
                Item:="Vnesite novo ime zavika"
    UItext = result(keyStr)
End Function

'Workbook functions
'==================
    'New workbook
    '------------
    Function newWorkbook()
        Dim wb As Workbook
        Set wb = Workbooks.Add

        'Display the new workbook on top
        AppActivate wb.Name

        'Return workbook name
        newWorkbook = wb.Name
    End Function

'Worksheet functions
'===================
    'New worksheet
    '-------------
    Function newWorksheet( _
                            Optional ByVal sheetName As String = "", _
                            Optional ByVal wbName As String = "", _
                            Optional ByVal promptForName As Boolean = False _
                            ) As String
        If promptForName = True Then
            sheetName = alertInput(UItext("newWorksheet"))
        End If

        If wbName = "" Then
            wbName = Application.ActiveWorkbook.Name
        End If

        Dim ws As Worksheet
        Set ws = Application.Workbooks(wbName).Worksheets.Add()

        If sheetName = "" Then
            'Return worksheet name
            newWorksheet = ws.Name
        Else 'sheetName <> ""
            renameWorksheet ws.Name, sheetName, wbName
            'Return worksheet name
            newWorksheet = sheetName
        End If
    End Function

    'Rename worksheet
    '----------------
    Function renameWorksheet( _
                                oldName As String, _
                                newName As String, _
                                Optional ByVal wbName As String = "", _
                                Optional ByVal promptForName As Boolean = False _
                                ) As String
        If promptForName = True Then
            newName = alertInput(UItext("renameWorksheet"))
        End If
        If wbName = "" Then
            wbName = Application.ActiveWorkbook.Name
        End If

        Application.Workbooks(wbName).Worksheets(oldName).Name = newName

        'Return new name
        renameWorksheet = newName
    End Function

    'Delete worksheet
    '----------------
    Function rmWorksheet( _
                            sheetName As String, _
                            Optional ByVal wbName As String = "", _
                            Optional ByVal promptBeforeRemoving As Boolean = False _
                            ) As Boolean
        If wbName = "" Then
            wbName = Application.ActiveWorkbook.Name
        End If

        If promptBeforeRemoving = False Then
            Application.DisplayAlerts = False
        End If

        'Silently fail if already deleted
        With Application.Workbooks(wbName)
            Dim exists As Boolean
            exists = False
            Dim i As Integer
            For i = 1 To .Worksheets.Count
                If LCase(.Worksheets(i).Name) = LCase(sheetName) Then
                    exists = True
                    Exit For
                End If
            Next i
            If exists = True Then
                .Worksheets(sheetName).Delete
            End If
        End With

        If promptBeforeRemoving = False Then
            Application.DisplayAlerts = True
        End If

        'RmWorksheet was successful
        rmWorksheet = True
    End Function

'Cell functions
'==============
    'Get cell (value)
    '----------------
    Function getCell( _
                    row As Long, _
                    column As Integer, _
                    Optional ByVal sheetName As String = "", _
                    Optional ByVal wbName As String = "", _
                    Optional ByVal getFormula As Boolean = False _
                    ) As Variant
        'Set default values
        If sheetName = "" Then
            sheetName = Application.ActiveWorkbook.ActiveSheet.Name
            wbName = Application.ActiveWorkbook.Name
        ElseIf wbName = "" Then
            wbName = Application.ActiveWorkbook.Name
        End If

        With Application.Workbooks(wbName).Worksheets(sheetName).Cells(row, column)
            If getFormula = True Then
                getCell = .Formula
            Else
                getCell = .Value
            End If
        End With
    End Function

    'Get cell formula
    '----------------
    Function getCellFormula( _
                    row As Long, _
                    column As Integer, _
                    Optional ByVal sheetName As String = "", _
                    Optional ByVal wbName As String = "" _
                    ) As String
        getCellFormula = getCell(row, column, sheetName, wbName, True)
    End Function

    'Set cell value
    '--------------
    Function setCell(val As Variant, _
                    row As Long, _
                    column As Integer, _
                    Optional ByVal sheetName As String = "", _
                    Optional ByVal wbName As String = "" _
                    ) As Boolean
        'Set default values
        If sheetName = "" Then
            sheetName = Application.ActiveWorkbook.ActiveSheet.Name
            wbName = Application.ActiveWorkbook.Name
        ElseIf wbName = "" Then
            wbName = Application.ActiveWorkbook.Name
        End If

        Application.Workbooks(wbName).Worksheets(sheetName).Cells(row, column).Value = val

        'Set was successful
        setCell = True
    End Function

    'Insert a single or a range of cells
    '-----------------------------------
    Function insertCell( _
                    startRow As Long, _
                    startColumn As Integer, _
                    Optional ByVal endRow As Long = -1, _
                    Optional ByVal endColumn As Integer = -1, _
                    Optional ByVal sheetName As String = "", _
                    Optional ByVal wbName As String = "" _
                    ) As Boolean
        'Set default values
        If endRow = -1 Then
            endRow = startRow
        End If
        If endColumn = -1 Then
            endColumn = startColumn
        End If
        If sheetName = "" Then
            sheetName = Application.ActiveWorkbook.ActiveSheet.Name
            wbName = Application.ActiveWorkbook.Name
        ElseIf wbName = "" Then
            wbName = Application.ActiveWorkbook.Name
        End If

        'Insert cell
        Application _
            .Workbooks(wbName) _
            .Worksheets(sheetName) _
            .Range( _
                    Cells(startRow, startColumn), _
                    Cells(endRow, endColumn) _
                    ).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

        'insertCell success
        insertCell = True
    End Function

    'Remove a single or a range of cells
    '-----------------------------------
    Function rmCell( _
                    startRow As Long, _
                    startColumn As Integer, _
                    Optional ByVal endRow As Long = -1, _
                    Optional ByVal endColumn As Integer = -1, _
                    Optional ByVal sheetName As String = "", _
                    Optional ByVal wbName As String = "" _
                    ) As Boolean
        'Set default values
        If endRow = -1 Then
            endRow = startRow
        End If
        If endColumn = -1 Then
            endColumn = startColumn
        End If
        If sheetName = "" Then
            sheetName = Application.ActiveWorkbook.ActiveSheet.Name
            wbName = Application.ActiveWorkbook.Name
        ElseIf wbName = "" Then
            wbName = Application.ActiveWorkbook.Name
        End If

        'Remove cell
        Application _
            .Workbooks(wbName) _
            .Worksheets(sheetName) _
            .Range( _
                    Cells(startRow, startColumn), _
                    Cells(endRow, endColumn) _
                    ).Delete Shift:=xlUp

        'rmCell success
        rmCell = True
    End Function


    'Clear cell value and styling
    '----------------------------
    Function clearCell( _
                    startRow As Long, _
                    startColumn As Integer, _
                    Optional ByVal endRow As Long = -1, _
                    Optional ByVal endColumn As Integer = -1, _
                    Optional ByVal sheetName As String = "", _
                    Optional ByVal wbName As String = "" _
                    ) As Boolean
        'Set default values
        If endRow = -1 Then
            endRow = startRow
        End If
        If endColumn = -1 Then
            endColumn = startColumn
        End If
        If sheetName = "" Then
            sheetName = Application.ActiveWorkbook.ActiveSheet.Name
            wbName = Application.ActiveWorkbook.Name
        ElseIf wbName = "" Then
            wbName = Application.ActiveWorkbook.Name
        End If

        'Clear cell
        Application.ScreenUpdating = False
        With Application _
            .Workbooks(wbName) _
            .Worksheets(sheetName) _
            .Range( _
                    Cells(startRow, startColumn), _
                    Cells(endRow, endColumn) _
                    )
            'Empty value
            .Value = ""
            'Clear formatting
            .ClearFormats
        End With
        Application.ScreenUpdating = True

        'rmCell was successful
        clearCell = True
    End Function
    
    'Comment a cell
    '--------------
    Function commentCell(comment As String, _
                        row As Long, _
                        column As Integer, _
                        Optional ByVal sheetName As String = "", _
                        Optional ByVal wbName As String = "" _
                        ) As Boolean
        'Set default values
        If sheetName = "" Then
            sheetName = Application.ActiveWorkbook.ActiveSheet.Name
            wbName = Application.ActiveWorkbook.Name
        ElseIf wbName = "" Then
            wbName = Application.ActiveWorkbook.Name
        End If
        
        'Remove existing comment (otherwise Excel will throw an error)
        Application.Workbooks(wbName).Worksheets(sheetName).Cells(row, column).comment.Delete
        
        'Add new comment
        Application.Workbooks(wbName).Worksheets(sheetName).Cells(row, column).AddComment (comment)
    
        'Comment was successful
        commentCell = True
    End Function
    
    'Remove cell background fill
    '---------------------------
    Function rmBgCell(row As Long, _
                    column As Integer, _
                    Optional ByVal sheetName As String = "", _
                    Optional ByVal wbName As String = "" _
                    ) As Boolean
        'Set default values
        If sheetName = "" Then
            sheetName = Application.ActiveWorkbook.ActiveSheet.Name
            wbName = Application.ActiveWorkbook.Name
        ElseIf wbName = "" Then
            wbName = Application.ActiveWorkbook.Name
        End If
        
        'Add new comment
        Application.Workbooks(wbName).Worksheets(sheetName).Cells(row, column).Interior.color = xlNone
    
        'Comment was successful
        commentCell = True
    End Function
    
    'Change cell background fill
    '---------------------------
    'Use RGB(255, 0, 0) to define a color value
    Function bgCell(color As String, _
                    row As Long, _
                    column As Integer, _
                    Optional ByVal sheetName As String = "", _
                    Optional ByVal wbName As String = "" _
                    ) As Boolean
        'Set default values
        If sheetName = "" Then
            sheetName = Application.ActiveWorkbook.ActiveSheet.Name
            wbName = Application.ActiveWorkbook.Name
        ElseIf wbName = "" Then
            wbName = Application.ActiveWorkbook.Name
        End If
        
        'Add new comment
        Application.Workbooks(wbName).Worksheets(sheetName).Cells(row, column).Interior.color = color
    
        'Comment was successful
        commentCell = True
    End Function

'Row and column functions
'========================
    'Sets a row of values
    '--------------------
    Function setRow( _
                    valueArray As Variant, _
                    row As Long, _
                    Optional ByVal sheetName As String = "", _
                    Optional ByVal wbName As String = "", _
                    Optional ByVal startColumn As Integer = 1 _
                    ) As Boolean
        'Set default values
        If sheetName = "" Then
            sheetName = Application.ActiveWorkbook.ActiveSheet.Name
            wbName = Application.ActiveWorkbook.Name
        ElseIf wbName = "" Then
            wbName = Application.ActiveWorkbook.Name
        End If

        Dim i As Integer
        Dim j As Integer 'to not depend on valueArray bounds
        j = 0
        For i = LBound(valueArray) To UBound(valueArray)
            setCell _
                valueArray(i), _
                row, _
                startColumn + j, _
                sheetName, _
                wbName
            j = j + 1
        Next i

        'setRow success
        setRow = True
    End Function

    'Get a row of values
    '-------------------
    Function getRow( _
                    row As Long, _
                    Optional ByVal sheetName As String = "", _
                    Optional ByVal wbName As String = "", _
                    Optional ByVal startColumn As Integer = 1, _
                    Optional ByVal endColumn As Integer = -1, _
                    Optional ByVal getFormula As Boolean = False _
                    ) As Variant
        'Set default values
        If sheetName = "" Then
            sheetName = Application.ActiveWorkbook.ActiveSheet.Name
            wbName = Application.ActiveWorkbook.Name
        ElseIf wbName = "" Then
            wbName = Application.ActiveWorkbook.Name
        End If
        If endColumn = -1 Then
            endColumn = lastColumn(row, sheetName, wbName)
        End If

        'If row is empty
        If endColumn - startColumn + 1 = 0 Then
            getRow = Array()
        Else
            'Get row
            Dim result As Variant
            ReDim result(1 To endColumn - startColumn + 1) As Variant
    
            Dim i As Integer
            For i = 1 To UBound(result)
                result(i) = getCell(row, i - 1 + startColumn, sheetName, wbName, getFormula)
            Next i
    
            getRow = result
        End If
    End Function

    'Get a row of formulas
    '---------------------
    Function getRowFormula( _
                    row As Long, _
                    Optional ByVal sheetName As String = "", _
                    Optional ByVal wbName As String = "", _
                    Optional ByVal startColumn As Integer = 1, _
                    Optional ByVal endColumn As Integer = -1 _
                    ) As Variant
        getRowFormula = getRow(row, sheetName, wbName, startColumn, endColumn, True)
    End Function

    'Get a column of values
    '----------------------
    Function getColumn( _
                    column As Integer, _
                    Optional ByVal sheetName As String = "", _
                    Optional ByVal wbName As String = "", _
                    Optional ByVal startRow As Long = 1, _
                    Optional ByVal endRow As Long = -1, _
                    Optional ByVal getFormula As Boolean = False _
                    ) As Variant
        'Set default values
        If sheetName = "" Then
            sheetName = Application.ActiveWorkbook.ActiveSheet.Name
            wbName = Application.ActiveWorkbook.Name
        ElseIf wbName = "" Then
            wbName = Application.ActiveWorkbook.Name
        End If
        If endRow = -1 Then
            endRow = lastRow(column, sheetName, wbName)
        End If

        'If column is empty
        If endRow - startRow + 1 = 0 Then
            getColumn = Array()
        Else
            'Get column
            Dim result As Variant
            ReDim result(1 To endRow - startRow + 1) As Variant
    
            Dim i As Integer
            For i = 1 To UBound(result)
                result(i) = getCell(i - 1 + startRow, column, sheetName, wbName, getFormula)
            Next i
    
            getColumn = result
        End If
    End Function


    'Get a column of formulas
    '------------------------
    Function getColumnFormula( _
                column As Integer, _
                Optional ByVal sheetName As String = "", _
                Optional ByVal wbName As String = "", _
                Optional ByVal startRow As Long = 1, _
                Optional ByVal endRow As Long = -1 _
                ) As Variant
        getColumnFormula = getColumn(column, sheetName, wbName, startRow, endRow, True)
    End Function

    'Set a column of values
    '----------------------
    Function setColumn( _
                    valueArray As Variant, _
                    column As Integer, _
                    Optional ByVal sheetName As String = "", _
                    Optional ByVal wbName As String = "", _
                    Optional ByVal startRow As Long = 1 _
                    ) As Boolean
        'Set default values
        If sheetName = "" Then
            sheetName = Application.ActiveWorkbook.ActiveSheet.Name
            wbName = Application.ActiveWorkbook.Name
        ElseIf wbName = "" Then
            wbName = Application.ActiveWorkbook.Name
        End If

        Dim i As Integer
        Dim j As Integer 'to not depend on valueArray bounds
        j = 0
        For i = LBound(valueArray) To UBound(valueArray)
            setCell _
                valueArray(i), _
                startRow + j, _
                column, _
                sheetName, _
                wbName
            j = j + 1
        Next i

        'setRow success
        setColumn = True
    End Function

    'Returns last row
    '----------------
    'Minimum return value = 1
    Function lastRow( _
                    column As Integer, _
                    Optional ByVal sheetName As String = "", _
                    Optional ByVal wbName As String = "" _
                    ) As Long
        'Set default values
        If sheetName = "" Then
            sheetName = Application.ActiveWorkbook.ActiveSheet.Name
            wbName = Application.ActiveWorkbook.Name
        ElseIf wbName = "" Then
            wbName = Application.ActiveWorkbook.Name
        End If

        With Application.Workbooks(wbName).Worksheets(sheetName)
            lastRow = .Cells(.Rows.Count, column).End(xlUp).row
        End With
    End Function

    'Returns last column
    '-------------------
    'Minimum return value = 1
    Function lastColumn( _
                        row As Long, _
                        Optional ByVal sheetName As String = "", _
                        Optional ByVal wbName As String = "" _
                        ) As Integer
        'Set default values
        If sheetName = "" Then
            sheetName = Application.ActiveWorkbook.ActiveSheet.Name
            wbName = Application.ActiveWorkbook.Name
        ElseIf wbName = "" Then
            wbName = Application.ActiveWorkbook.Name
        End If

        With Application.Workbooks(wbName).Worksheets(sheetName)
            lastColumn = .Cells(row, .columns.Count).End(xlToLeft).column
        End With
    End Function

    'Inserts entire row
    '------------------
    Function insertRow( _
                        insertBeforeRow As Long, _
                        Optional ByVal sheetName As String = "", _
                        Optional ByVal wbName As String = "" _
                        ) As Boolean
        'Set default values
        If sheetName = "" Then
            sheetName = Application.ActiveWorkbook.ActiveSheet.Name
            wbName = Application.ActiveWorkbook.Name
        ElseIf wbName = "" Then
            wbName = Application.ActiveWorkbook.Name
        End If

        'Insert row
        Application _
            .Workbooks(wbName) _
            .Worksheets(sheetName) _
            .Cells(insertBeforeRow, 1) _
            .EntireRow.Insert , _
                CopyOrigin:=xlFormatFromLeftOrAbove

        'insertRow success
        insertRow = True
    End Function

    'Inserts entire column
    '---------------------
    Function insertColumn( _
                        insertBeforeColumn As Integer, _
                        Optional ByVal sheetName As String = "", _
                        Optional ByVal wbName As String = "" _
                        ) As Boolean
        'Set default values
        If sheetName = "" Then
            sheetName = Application.ActiveWorkbook.ActiveSheet.Name
            wbName = Application.ActiveWorkbook.Name
        ElseIf wbName = "" Then
            wbName = Application.ActiveWorkbook.Name
        End If

        'Insert row
        Application _
            .Workbooks(wbName) _
            .Worksheets(sheetName) _
            .Cells(1, insertBeforeColumn) _
            .EntireColumn.Insert , _
                CopyOrigin:=xlFormatFromLeftOrAbove

        'insertColumn success
        insertColumn = True
    End Function

    'Remove entire row
    '-----------------
    Function rmRow( _
                    row As Long, _
                    Optional ByVal sheetName As String = "", _
                    Optional ByVal wbName As String = "" _
                    ) As Boolean
            'Set default values
            If sheetName = "" Then
                sheetName = Application.ActiveWorkbook.ActiveSheet.Name
                wbName = Application.ActiveWorkbook.Name
            ElseIf wbName = "" Then
                wbName = Application.ActiveWorkbook.Name
            End If

            'Remove row
            Application _
                .Workbooks(wbName) _
                .Worksheets(sheetName) _
                .Cells(row, 1) _
                .EntireRow.Delete

            'rmRow success
            rmRow = True
    End Function

    'Remove entire column
    '--------------------
    Function rmColumn( _
                    column As Integer, _
                    Optional ByVal sheetName As String = "", _
                    Optional ByVal wbName As String = "" _
                    ) As Boolean
            'Set default values
            If sheetName = "" Then
                sheetName = Application.ActiveWorkbook.ActiveSheet.Name
                wbName = Application.ActiveWorkbook.Name
            ElseIf wbName = "" Then
                wbName = Application.ActiveWorkbook.Name
            End If

            'Remove column
            Application _
                .Workbooks(wbName) _
                .Worksheets(sheetName) _
                .Cells(1, column) _
                .EntireColumn.Delete

            'rmColumn success
            rmColumn = True
    End Function

'Table functions
'===============
    'Set an entire matrix
    '--------------------
    Function setMatrix( _
                        valueMatrix As Variant, _
                        Optional ByVal startRow As Long = 1, _
                        Optional ByVal startColumn As Integer = 1, _
                        Optional ByVal sheetName As String = "", _
                        Optional ByVal wbName As String = "" _
                        ) As Boolean
        Dim i As Integer
        Dim j As Long 'to avoid
        j = 0
        For i = LBound(valueMatrix) To UBound(valueMatrix)
            setRow valueMatrix(i), startRow + j, sheetName, wbName, startColumn
            j = j + 1
        Next i

        'setMatrix success
        setMatrix = True
    End Function

    'Get the last row in a matrix
    '----------------------------
    'For performance reasons endColumn
    'is recommended (defaults to excel limit)
    Function lastRowMatrix( _
                            startColumn As Integer, _
                            Optional ByVal endColumn As Integer = -1, _
                            Optional ByVal sheetName As String = "", _
                            Optional ByVal wbName As String = "" _
                            ) As Variant
        If endColumn = -1 Then
            endColumn = excelMaxColumn()
        End If

        Dim row As Long
        Dim maxRow As Long
        maxRow = 1

        Dim i As Integer
        For i = startColumn To endColumn
            row = lastRow(i, sheetName, wbName)
            If row > maxRow Then
                maxRow = row
            End If
        Next i

        lastRowMatrix = maxRow
    End Function

    'Get the last column in a matrix
    '-------------------------------
    'For performance reasons endRow
    'is recommended (defaults to excel limit)
    Function lastColumnMatrix( _
                            startRow As Long, _
                            Optional ByVal endRow As Long = -1, _
                            Optional ByVal sheetName As String = "", _
                            Optional ByVal wbName As String = "" _
                            ) As Variant
        If endRow = -1 Then
            endRow = excelMaxRow()
        End If

        Dim column As Integer
        Dim maxColumn As Integer
        maxColumn = 1

        Dim i As Long
        For i = startRow To endRow
            column = lastColumn(i, sheetName, wbName)
            If column > maxColumn Then
                maxColumn = column
            End If
        Next i

        lastColumnMatrix = maxColumn
    End Function

    'Return matrix (2d array) of values
    '----------------------------------
    Function getMatrix( _
                        startRow As Long, _
                        startColumn As Integer, _
                        Optional ByVal endRow As Long = -1, _
                        Optional ByVal endColumn As Integer = -1, _
                        Optional ByVal sheetName As String = "", _
                        Optional ByVal wbName As String = "", _
                        Optional ByVal getFormula As Boolean = False _
                        ) As Variant
        'Default last row
        If endRow = -1 Then
            endRow = lastRowMatrix(startColumn, endColumn, sheetName, wbName)
        End If

        Dim result As Variant
        ReDim result(1 To endRow - startRow + 1) As Variant

        Dim i As Long
        For i = 1 To UBound(result)
            result(i) = getRow(startRow + i - 1, sheetName, wbName, startColumn, endColumn, getFormula)
        Next i

        'return 2d array of rows
        getMatrix = result
    End Function
