Attribute VB_Name = "Compile"
'-   Previous month
'~   Previous quarter
'^   Previous year
'@   Current month
'+   Following month
'*   Following year
'$   Sum of previous quarter
'#   Sum of current quarter
'!   Sum of current year
'%   Sum of following year
Sub Cube_Compile()

Dim accountsArray() As Variant
Dim rowsCount As Long
Dim columnsCount As Long
Dim totalColumns As Long
Dim c As Long
Dim columnName As String
Dim formulaColumn As Long
Dim fomulaTextColumn As Long
Dim accountColumn As Long
Dim monthsInput As Long
Dim monthNoColumn As Long
Dim monthNameColumn As Long
Dim quarterNoColumn As Long
Dim quarterNameColumn As Long
Dim yearNoColumn As Long
Dim yearNameColumn As Long
Dim periodNameColumn As Long
Dim basisColumn As Long
Dim outputArray() As Variant
Dim bottomRight As Range
Dim topRow As Long
Dim columnOffset As Long
Dim topLeft As Range
Dim r As Long
Dim m As Long
Dim currentRow As Long
Dim inputFormula As String
Dim outputFormula As String
Dim outputColumn As Long
Dim ch As Long
Dim currentChar As String
Dim parseBracket As Boolean
Dim locationModifier As String
Dim bracketAccount As String
Dim accountName As String
Dim accountRow As Long
Dim y As Long
Dim p As Long
Dim q As Long
Dim periodName As String
Dim periodBasis As String
'variables for start and end search indices
Dim currentProduct As String
Dim currentSegment As String
Dim endIndex() As Long 'index value for current product
Dim endProduct As Long
Dim nextProduct As String
Dim nextSegment As String
Dim previousProduct As String
Dim previousSegment As String
Dim productColumn As Long
Dim segmentColumn As Long
Dim startIndex() As Long 'index value for current product
Dim startProduct As Long 'lookup for starting row of current product

'Grab Input Array
accountsArray = ActiveCell.CurrentRegion.Value
rowsCount = UBound(accountsArray)
columnsCount = UBound(accountsArray, 2)
'find special columns
For c = 1 To columnsCount
    columnName = accountsArray(1, c)
    Select Case columnName
        Case "Formula"
            formulaColumn = c
        Case "Account name"
            accountColumn = c
        Case "Period Basis"
            basisColumn = c
        Case "Product"
            productColumn = c
        Case "Segment"
            segmentColumn = c
    End Select
Next

'Set Variables
'additional columns (in order)
totalColumns = columnsCount
formulaTextColumn = totalColumns + 1
totalColumns = totalColumns + 1
monthNoColumn = totalColumns + 1
totalColumns = totalColumns + 1
monthNameColumn = totalColumns + 1
totalColumns = totalColumns + 1
quarterNoColumn = totalColumns + 1
totalColumns = totalColumns + 1
quarterNameColumn = totalColumns + 1
totalColumns = totalColumns + 1
yearNoColumn = totalColumns + 1
totalColumns = totalColumns + 1
yearNameColumn = totalColumns + 1
totalColumns = totalColumns + 1
periodNameColumn = totalColumns + 1
totalColumns = totalColumns + 1
'user input and output placement
Set topLeft = Application.InputBox("Select top left cell for cube output:", Type:=8).Cells(1, 1)
topRow = topLeft.Row - 1
columnOffset = topLeft.Column - 1
'monthsInput = Application.InputBox("Number of months:", Type:=1) 'change if standard
monthsInput = 36 'change if variable
'Set bottomRight = Cells((rowsCount - 1) * monthsInput + topRow, totalColumns + columnOffset)
Set bottomRight = topLeft.Parent.Cells((rowsCount - 1) * monthsInput + topRow, totalColumns + columnOffset)
ReDim outputArray(1 To (rowsCount - 1) * monthsInput, 1 To totalColumns)
outputColumn = formulaColumn + columnOffset
'length of start and end search index
ReDim startIndex(1 To rowsCount)
ReDim endIndex(1 To rowsCount)

'Generate start search index
'handle first row
previousSegment = accountsArray(1, productColumn)
previousProduct = accountsArray(1, segmentColumn)
startProduct = 1
startIndex(1) = startProduct
'cycle through rows
For r = 2 To rowsCount
    'get product and segment for this row
    currentSegment = accountsArray(r, segmentColumn)
    currentProduct = accountsArray(r, productColumn)
    'beginning of next product
    If Not currentProduct = previousProduct Or Not currentSegment = previousSegment Then
        startProduct = r
    End If
    startIndex(r) = startProduct
    'set product and segment for next row
    previousSegment = currentSegment
    previousProduct = currentProduct
Next

'Generate end search index
'handle last row
nextSegment = accountsArray(rowsCount, segmentColumn)
nextProduct = accountsArray(rowsCount, productColumn)
endProduct = rowsCount
endIndex(rowsCount) = endProduct
'cycle backwards through rows
For r = rowsCount - 1 To 1 Step -1
    'get product and segment for this row
    currentSegment = accountsArray(r, segmentColumn)
    currentProduct = accountsArray(r, productColumn)
    'beginning of next product
    If Not currentProduct = nextProduct Or Not currentSegment = nextSegment Then
        endProduct = r
    End If
    endIndex(r) = endProduct
    'set product and segment for next row
    nextSegment = currentSegment
    nextProduct = currentProduct
Next

'Cycle Rows
For r = 2 To rowsCount
    periodBasis = accountsArray(r, basisColumn)
    startProduct = startIndex(r)
    endProduct = endIndex(r)

'Cycle Months
For m = 1 To monthsInput
    currentRow = (r - 2) * monthsInput + m
    y = Int((m / 12) + 0.99)
    
'Cycle Columns
For c = 1 To columnsCount
    If c = formulaColumn Then
        'Set Formula
        inputFormula = accountsArray(r, c)
        outputFormula = "="
        'Cycle char
        For ch = 1 To Len(inputFormula)
            currentChar = Mid(inputFormula, ch, 1)
            If parseBracket Then 'within bracket account
                Select Case currentChar
                'end bracket account
                Case "}"
                    locationModifier = Left(bracketAccount, 1)
                    accountName = Right(bracketAccount, Len(bracketAccount) - 1)
                    'reserved account name
                    If accountName = "Month" Then
                        bracketAccount = Cells(currentRow + topRow, monthNoColumn + columnOffset).Address(False, True)
                    Else
                    'find account
                    accountRow = 0
                    'only search in current product (i.e. TLD)
                    For a = startProduct To endProduct
                        If accountsArray(a, accountColumn) = accountName Then
                            accountRow = a
                            Exit For
                        End If
                    Next 'account
                    'catch mistakes in spelling and exit
                    If accountRow = 0 Then
                        MsgBox ("The Account Code: " & accountName & " was not found.")
                        Exit Sub
                    End If
                    'set location
                    accountRow = (accountRow - 2) * monthsInput + m + topRow
                    'modify location
                    Select Case locationModifier
                        Case "@" 'Current month
                            bracketAccount = Cells(accountRow, outputColumn).Address(False, True)
                        Case "-" 'Previous month
                            If m > 1 Then
                                bracketAccount = Cells(accountRow - 1, outputColumn).Address(False, True)
                            Else
                                bracketAccount = "0"
                            End If
                        Case "+" 'Following month
                            bracketAccount = Cells(accountRow + 1, outputColumn).Address(False, True)
                        Case "^" 'Previous year
                            If m > 12 Then
                                bracketAccount = Cells(accountRow - 12, outputColumn).Address(False, True)
                            Else
                                bracketAccount = "0"
                            End If
                        Case "~" 'Previous quarter
                            If m > 3 Then
                                bracketAccount = Cells(accountRow - 3, outputColumn).Address(False, True)
                            Else
                                bracketAccount = "0"
                            End If
                        Case "!" 'Sum of current year
                            bracketAccount = "Sum(" & Cells(accountRow - Min_Value(12, m) + 1, outputColumn).Address(False, True) _
                            & ":" & Cells(accountRow, formulaColumn).Address(False, True) _
                            & ")"
                        Case "#" 'Sum of current quarter
                            bracketAccount = "Sum(" & Cells(accountRow - Min_Value(3, m) + 1, outputColumn).Address(False, True) _
                            & ":" & Cells(accountRow, formulaColumn).Address(False, True) _
                            & ")"
                        Case "$" 'Sum of previous quarter
                            If m > 12 Then
                                bracketAccount = "Sum(" & Cells(accountRow - m + 1 + ((y - 2) * 12), outputColumn).Address(False, True) _
                                & ":" & Cells(accountRow - m + ((y - 2) * 12) + 12, outputColumn).Address(False, True) _
                                & ")"
                            Else
                                bracketAccount = "0"
                            End If
                        Case "%" 'Sum of following year
                            bracketAccount = "Sum(" & Cells(accountRow + 1, outputColumn).Address(False, True) _
                            & ":" & Cells(accountRow + 12 + 1, outputColumn).Address(False, True) _
                            & ")"
                        Case "*" 'Following year
                            bracketAccount = Cells(accountRow + 13, outputColumn).Address(False, True)
                    End Select
                    End If 'account name
'                    End If
                    outputFormula = outputFormula & bracketAccount
                    parseBracket = False
                Case Else
                    bracketAccount = bracketAccount & currentChar
                End Select
            Else 'not in bracket account
                Select Case currentChar
                Case "{"
                    If Right(outputFormula, 1) = "'" Then 'escape character
                        outputFormula = Left(outputFormula, Len(outputFormula) - 1) & currentChar
                    Else
                        bracketAccount = ""
                        parseBracket = True 'signal in bracket account
                    End If
                Case Else
                    outputFormula = outputFormula & currentChar
                End Select
            End If
        Next 'char
        outputArray(currentRow, c) = outputFormula
    Else
        outputArray(currentRow, c) = accountsArray(r, c)
    End If
Next 'column

'add time dimensions
p = Int((m - ((y - 1) * 12)) + 0.99)
q = Int((p / 3) + 0.99)
outputArray(currentRow, monthNoColumn) = m
outputArray(currentRow, monthNameColumn) = MonthName(p)
outputArray(currentRow, quarterNoColumn) = q
outputArray(currentRow, quarterNameColumn) = "Q" & CStr(q)
outputArray(currentRow, yearNoColumn) = y
outputArray(currentRow, yearNameColumn) = y + 2013
periodName = CStr(y + 2013)
Select Case periodBasis
    Case "Monthly"
        periodName = periodName & " " & MonthName(p)
    Case "Quarterly"
        periodName = periodName & " Q" & CStr(q)
    Case "Yearly"
        periodName = periodName & " Year End"
End Select
outputArray(currentRow, periodNameColumn) = periodName
outputArray(currentRow, formulaTextColumn) = accountsArray(r, formulaColumn)
'add product dimensions


Next 'period
Next 'row

'Place Output Array
Range(topLeft, bottomRight).Value = outputArray

End Sub
Function Min_Value(valueOne, valueTwo)
    If valueOne < valueTwo Then
        Min_Value = valueOne
    Else
        Min_Value = valueTwo
    End If
End Function




