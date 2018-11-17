'Copyright 2018 Benson Mariro and Bennett Kankuzi, Dept of Computer Science, North-West University, South Africa.
'Code released under a Creative Commons Attribution 4.0 International Licence.

Option Explicit
Option Compare Text
Dim colHeaderString As String
    
Dim rowHeaderString As String

Dim State As Integer

Public Function ReturnState() As Integer
    'State = 1
    ReturnState = State
End Function


Function ColNo2ColRef(ColNo As Integer) As String
    If ColNo < 1 Or ColNo > 256 Then
        ColNo2ColRef = "#VALUE!"
        Exit Function
    End If
    ColNo2ColRef = Cells(1, ColNo).Address(True, False, xlA1)
    ColNo2ColRef = Left(ColNo2ColRef, InStr(1, ColNo2ColRef, "$") - 1)
End Function
Function FullyTranslatedFormula(cellFormula As String) As String

Dim num As Integer
If InStr(cellFormula, "-") > 0 Then

    cellFormula = Replace(cellFormula, "-", " minus ")
End If
    

If InStr(cellFormula, "+") > 0 Then

    cellFormula = Replace(cellFormula, "+", " plus ")
End If
    
If InStr(cellFormula, ":") > 0 Then

    cellFormula = Replace(cellFormula, ":", " to ")
    
End If

If InStr(cellFormula, "^") > 0 Then

    cellFormula = Replace(cellFormula, "^", " exponent ")
    
End If

If InStr(cellFormula, "%") > 0 Then

    cellFormula = Replace(cellFormula, "%", " percent ")
    
End If

If InStr(cellFormula, "=") > 0 Then

    cellFormula = Replace(cellFormula, "=", " equal to")
    
End If

If InStr(cellFormula, ">") > 0 Then

   cellFormula = Replace(cellFormula, ">", "is greater than ")
    
End If

If InStr(cellFormula, "<") > 0 Then

     cellFormula = Replace(cellFormula, "<", "is less than ")
     
End If

If InStr(cellFormula, ">=") > 0 Then

    cellFormula = Replace(cellFormula, ">=", "greater than or equal to")
    
End If

If InStr(cellFormula, "<=") > 0 Then

    cellFormula = Replace(cellFormula, "<=", "less than or equal to")
    
End If

If InStr(cellFormula, "<>") > 0 Then

    cellFormula = Replace(cellFormula, "<>", "not equal to")
    
End If

If InStr(cellFormula, "*") > 0 Then

    cellFormula = Replace(cellFormula, "*", "multiplied by")
    
End If

If InStr(cellFormula, "/") > 0 Then
    
    cellFormula = Replace(cellFormula, "/", " divided by ")
     
End If

If InStr(cellFormula, "SUM") > 0 Then

    cellFormula = Replace(cellFormula, "SUM(", " the sum of ")
    
    'If Comma_Str(cellFormula) > 0 Then
    
    
        ' Dim i As Integer
        'i = Comma_Str(cellFormula)
         'cellFormula = Mid(cellFormula, 1, i - 1) & Replace(cellFormula, ",", " and ", Start:=i)
    'End If
    
    cellFormula = Replace(cellFormula, ")", "")
    
    
End If

If InStr(cellFormula, "AVERAGE") > 0 Then
       
    cellFormula = Replace(cellFormula, "AVERAGE(", "the average of ")
     If Comma_Str(cellFormula) > 0 Then
    
         Dim p As Integer
        p = Comma_Str(cellFormula)
         cellFormula = Mid(cellFormula, 1, p + 1) & Replace(cellFormula, ",", " and ", Start:=p + 1)
    End If
    cellFormula = Replace(cellFormula, ")", "")
            
End If

If InStr(cellFormula, "IF") > 0 Then

    

    cellFormula = Replace(cellFormula, "IF(", " if")
    cellFormula = Replace(cellFormula, "(", "")
    
    Dim cellAdd, colHead, colH As String
    Dim cellAddT() As String
    Dim rowN As Integer
    Dim colN As Integer
    Dim iw, il As Integer
    cellAdd = ActiveCell.Address(True, True)
                'MsgBox cellAdd
                cellAddT() = Split(cellAdd, "$")
                rowN = CInt(cellAddT(2))
                colN = ColRef2ColNo(cellAddT(1))
                 colHead = colHeader(rowN, colN)
                colH = ", " & colHead
                
                
    cellFormula = Replace(cellFormula, ",", colH & " the value is ", 1, 1)
    
    il = InStr(cellFormula, colH)
    iw = InStr(il, cellFormula, ",")
    'MsgBox Mid(cellFormula, 1, iw)
    'MsgBox Replace(cellFormula, ",", ", otherwise ", Start:=iw, Count:=1)
    cellFormula = Mid(cellFormula, 1, iw) & Replace(cellFormula, ",", "; otherwise ", Start:=iw + 1, Count:=1)
    
    'MsgBox cellFormula
   
    cellFormula = Replace(cellFormula, ")", "")
    cellFormula = Replace(cellFormula, Chr(34), "")

End If
    
    FullyTranslatedFormula = cellFormula
End Function
Function Comma_Str(ByRef cellFormula) As Integer
    
    Comma_Str = InStr(cellFormula, ",")
End Function

Function Parser_Formula(ByRef cellFormula As String) As String
Dim My As Object, Ny As Object, M As Object, Match As Object, SB As Object
Dim trans(), transf, transl As String
Dim ArrLen, last, curr, i As Integer


Set My = CreateObject("vbscript.regexp")
My.ignorecase = True
My.Global = True
ReDim trans(Len(cellFormula))
My.Pattern = "[a-zA-z]+\([^()]+\)"
last = UBound(trans)
curr = last
If My.test(cellFormula) Then
While My.test(cellFormula)
    
    Set M = My.Execute(cellFormula)
    
    For Each Match In M
                
        If Len(Match.Value) <= Len(cellFormula) Then
        
            'MsgBox Match.Value
            transl = FullyTranslatedFormula(Match.Value)
            'MsgBox transl
            'i = InStrRev(transl, ",")
            'If i > 0 Then
            
            'transl = Mid(transl, 1, i - 1) & Replace(transl, ",", "; otherwise ", Start:=i, Count:=1)
            'MsgBox Mid(transl, 1, i - 1)
            'End If
            
            'MsgBox transl
            
            cellFormula = Replace(cellFormula, Match.Value, transl)
        
            
            
        End If
       
    Next Match
Wend

Parser_Formula = cellFormula
Else

Parser_Formula = FullyTranslatedFormula(cellFormula)

End If



End Function


Function ColRef2ColNo(ColRef As String) As Integer
    ColRef2ColNo = 0
    On Error Resume Next
    ColRef2ColNo = Range(ColRef & "1").Column
End Function
Function RowHeader(rowNumber As Integer, colNumber As Integer) As String
    Dim counter As Integer
    Dim testCounter As Integer
    Dim myCell As Range
    Dim myCell2 As Range
    
    Dim numberLabel As Boolean
    RowHeader = "---"
    For counter = colNumber To 1 Step -1
        
        'If Not IsEmpty(Cells(rowNumber, counter)) Then
                    'If Application.IsText(Cells(rowNumber, counter)) Then
                        'MsgBox "Text is row --" & Cells(rowNumber, counter).Text
                        'RowHeader = Cells(rowNumber, counter).Text
                        ''Cells(rowNumber, counter).Interior.ColorIndex = 20
                        'Exit For
                        
                    'End If
        'End If
        
        If Not IsEmpty(Cells(rowNumber, counter)) Then
        
                    numberLabel = False
                    Set myCell2 = Nothing
                                
                                
                                
                                
                    
                    If Application.IsText(Cells(rowNumber, counter)) Or Cells(rowNumber, counter).NumberFormat = "@" Or numberLabel = True Then
                        
                        RowHeader = Cells(rowNumber, counter).Text
                        
                        
                        If counter > 1 Then
                            
                            testCounter = counter - 1
                            
                             For testCounter = counter - 1 To 1 Step -1
                                'numberLabel = False
                                'If Cells(rowNumber, testCounter).HasFormula Then
                                   'MsgBox "has formula"
                                   'If Cells(rowNumber, testCounter).DirectPrecedents.Count = 1 Then
                                       'MsgBox "has one precedent"
                                       'For Each myCell In Cells(rowNumber, testCounter).DirectPrecedents.Cells
                                           'MsgBox myCell.Address(True, True)
                                           'If myCell.Address(True, True) = Cells(rowNumber - 1, testCounter).Address(True, True) Then
                                               'MsgBox "am a number label" & Cells(rowNumber, testCounter).Address(False, False)
                                               'RowHeader = Cells(rowNumber, testCounter).Value
                                               'numberLabel = True
                                           'End If
                                       'Next
                                       
                                   'End If
                                 'End If
                              
                              
                              
                              
                              
                              'Or IsDate(Cells(rowNumber, testCounter)) removed Or
                              If Application.IsText(Cells(rowNumber, testCounter)) Or Cells(rowNumber, testCounter).NumberFormat = "@" Or numberLabel = True Then
                                If Trim(Cells(rowNumber, testCounter).Text) = "-" Then
                                    'MsgBox "wawwa"
                                Else
                                    RowHeader = Cells(rowNumber, testCounter).Text & " " & RowHeader
                                End If
                                    'MsgBox RowHeader & " " & Cells(rowNumber, testCounter).Address(True, True)
                                    'If RowHeader = " " Then
                                        'MsgBox "am empty"
                                    'End If
                                'End If
                              Else
                                Exit For
                              End If
                              
                              
                             Next
                             
                             
                             RowHeader = RowHeader
                           'Exit For
                        End If
                        
                        
                        
                        Exit For
                        
                    End If
        End If
    Next
    RowHeader = RowHeader
End Function


Function colHeader(rowNumber As Integer, colNumber As Integer) As String
    Dim counter As Integer
    Dim testCounter As Integer
    Dim myCell As Range
    Dim myCell2 As Range
    Dim numberLabel As Integer
    colHeader = "---"
    
    For counter = rowNumber To 1 Step -1
        
        If Not IsEmpty(Cells(counter, colNumber)) Then
        
        
                    numberLabel = False
                                
                                            
                    
                    If Application.IsText(Cells(counter, colNumber)) Or Cells(counter, colNumber).NumberFormat = "@" Or numberLabel = True Then
                        
                        colHeader = Cells(counter, colNumber).Text
                        
                        'MsgBox colHeader & "column heading"
                        
                        If counter > 1 Then
                            
                            testCounter = counter - 1
                            
                             For testCounter = counter - 1 To 1 Step -1
                             
                              'If Cells(testCounter, colNumber).HasFormula Then
                                'MsgBox "has formula"
                                'If Cells(testCounter, colNumber).DirectPrecedents.Count = 1 Then
                                    'MsgBox "has one precedent"
                                    'For Each myCell In Cells(testCounter, colNumber).DirectPrecedents.Cells
                                        'MsgBox myCell.Address(True, True)
                                        'If myCell.Address(True, True) = Cells(testCounter - 1, colNumber) Then
                                            'MsgBox "am a number label" & Cells(testCounter, colNumber).Address(False, False)
                                        'End If
                                    'Next
                                    
                                'End If
                              'End If
                             
                              'Or IsDate(Cells(testCounter, colNumber)) removed
                              If Application.IsText(Cells(testCounter, colNumber)) Or Cells(testCounter, colNumber).NumberFormat = "@" Then
                                If Trim(Cells(testCounter, colNumber).Text) = "-" Then
                                    'Do Nothing
                                    'MsgBox Cells(testCounter, colNumber).Address(True, True)
                                    'colHeader = colHeader
                                Else
                                
                                    colHeader = Cells(testCounter, colNumber).Text
                                    '& " " & colHeader
                                    
                                End If
                              Else
                                Exit For
                              End If
                              
                              
                             Next
                             
                             
                             'colHeader = colHeader
                           'Exit For
                        End If
                        
                        
                        
                        Exit For
                        
                    End If
        End If
    
    Next
    'colHeader = colHeader
End Function

Public Sub DomainVisualizeAllColumnToRow()

    'visualizeType = 2 if its from column to row
    State = 2
    DomainVisualizeAllFormulaCells (State)
    
End Sub

Sub DomainVisualizeAllRowToColumn()
    'visualizeType = 1 if its from row to column
    State = 2
    DomainVisualizeAllFormulaCells (State)
    
End Sub

Sub DomainVisualizeAllFormulaCells(visualizeType As Integer)

'Application.ScreenUpdating = False

'visualizeType = 1 if its from row to column
'visualizeType = 2 if its from column to row

Dim oWS As Worksheet
Dim oCell As Range
Dim cellAdd As String
Dim val As Variant
Dim cellAddTokens() As String
Dim rowNumber As Integer
Dim colNumber As Integer

Set oWS = ActiveSheet
On Error Resume Next
For Each oCell In oWS.Cells.SpecialCells(xlCellTypeFormulas)
    
    'oCell.Interior.ColorIndex = 36
    'MsgBox oCell.Formula
    cellAdd = oCell.Offset(0, 0).Address(True, True)
    
    cellAddTokens() = Split(cellAdd, "$")
    
    rowNumber = CInt(cellAddTokens(2))
    
    colNumber = ColRef2ColNo(cellAddTokens(1))
    
    Dim cellFormula As String
    cellFormula = oCell.Formula
    cellFormula = Replace(cellFormula, "=", "", 1, 1)
    cellFormula = Replace(cellFormula, "$", "")
    
    'MsgBox cellFormula
    
    
    
    'go column-wise
    'Public colHeaderString As String
    'colHeaderString = colHeader(rowNumber, colNumber)
    'MsgBox "column header " & colHeaderString
    
    'go row-wise
    'Public rowHeaderString As String
    'rowHeaderString = RowHeader(rowNumber, colNumber)
    'MsgBox "row header" & rowHeaderString
    Dim precedentsRange As Range, cel As Range, myCell As Range, myCell2 As Range
    
    Dim directPrecedentsString As String
    Dim cellHeader As String
    Dim tempRow As Integer
    Dim tempCol As Integer
    
    Dim numberLabel As Boolean
    
    
    On Error Resume Next
    If Cells(rowNumber, colNumber).DirectPrecedents.Count Then
        Cells(rowNumber, colNumber).DirectPrecedents.Cells
        
        For Each cel In Cells(rowNumber, colNumber).DirectPrecedents.Cells
            
            'MsgBox "cel  " & cel.Address(False, False)
            numberLabel = False 'in cases like fact(n)
            
            cellAdd = cel.Address(True, True)
            cellAddTokens() = Split(cellAdd, "$")
            rowNumber = CInt(cellAddTokens(2))
            colNumber = ColRef2ColNo(cellAddTokens(1))
            
            
            If numberLabel = False Then
                cellAdd = cel.Address(True, True)
                
                cellAddTokens() = Split(cellAdd, "$")
                rowNumber = CInt(cellAddTokens(2))
                colNumber = ColRef2ColNo(cellAddTokens(1))
                colHeaderString = colHeader(rowNumber, colNumber)
                'MsgBox "column header " & colHeaderString
                rowHeaderString = RowHeader(rowNumber, colNumber)
                'MsgBox "row header" & rowHeaderString
            End If
            If visualizeType = 1 Then
            cellHeader = colHeaderString & " for " & rowHeaderString
            Else
                If visualizeType = 2 Then
                    cellHeader = colHeaderString & " for " & rowHeaderString
                End If
                
            End If
            
            cellFormula = Replace(cellFormula, cel.Address(False, False), " " & cellHeader & " ")
            cellFormula = Replace(cellFormula, ":", " to ")
            
            'cellFormula = Replace(cellFormula, "=", "")
            
            cellFormula = Replace(cellFormula, "--- |", "")
            cellFormula = Replace(cellFormula, "| ---", "")
            cellFormula = Replace(cellFormula, "---", "unnamed")
            cellFormula = Trim(cellFormula)
            
            
            'MsgBox cellFormula
            'MsgBox cellAdd
            'Range(cel.Address(False, False)).Interior.color = 20
            directPrecedentsString = directPrecedentsString & "," & cel.Address(False, False)
        Next
        
        'MsgBox Cells(rowNumber, colNumber).Precedents.Count & " dependancies found."
    Else
        MsgBox "No dependancies found."
    End If
    
    'MsgBox "Equivalent Formula is " & cellFormula
    oCell.ClearComments
    cellFormula = Parser_Formula(cellFormula)
    
    oCell.AddComment cellFormula
    
    oCell.Borders(xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
    oCell.Borders(xlEdgeRight).ColorIndex = 26
    oCell.Borders(xlEdgeRight).Weight = xlThick
    
    If oCell.Borders(xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous Then
     Else
        oCell.Borders(xlEdgeLeft).LineStyle = XlLineStyle.xlLineStyleNone
    End If
    'MsgBox oCell.Comment(1)
    With oCell.Comment
        '.Shape.TextFrame.AutoSize = True
        .Shape.Fill.ForeColor.SchemeColor = 7
        .Shape.TextFrame.Characters.Font.Name = "Arial"
        .Shape.TextFrame.Characters.Font.Size = 14
        '.Shape.TextFrame.Characters.Font.Bold = True
        '.Shape.TextEffect.FontBold = msoTrue
    
               
        '.Shape.ScaleWidth 3, msoFalse, msoScaleFromTopLeft
        .Shape.TextFrame2.WordWrap = msoTrue
        .Shape.Top = .Parent.Top + 18
        
        Dim CommentSize As Long
        .Shape.TextFrame.AutoSize = True
        If .Shape.Width > 250 Then
            CommentSize = .Shape.Width * .Shape.Height
            .Shape.Width = 250
            .Shape.Height = (CommentSize / 180) * 1.2
        End If
        Dim pos As Integer
        Dim startPos As Integer
        Dim cnt As Integer
        
        With .Shape.TextFrame
            
            If InStr(1, cellFormula, "if") > 0 Then
                startPos = 1
                cnt = Len("if")
                Do
                    pos = InStr(startPos, cellFormula, "if ", 1)
                    If pos > 0 Then
                        With .Characters(InStr(startPos, cellFormula, "if ", 1), cnt).Font

                            .color = RGB(64, 0, 128)
                            .Italic = True

                        End With
                    End If
                    startPos = pos + 1
                Loop While (pos > 0)
            End If
            
            If InStr(1, cellFormula, "otherwise") > 0 Then
                startPos = 1
                cnt = Len("otherwise")
                Do
                    pos = InStr(startPos, cellFormula, "otherwise ", 1)
                    If pos > 0 Then
                        With .Characters(InStr(startPos, cellFormula, "otherwise ", 1), cnt).Font

                            .color = RGB(64, 0, 128)
                            .Italic = True

                        End With
                    End If
                    startPos = pos + 1
                Loop While (pos > 0)
            End If
            If InStr(1, cellFormula, "to") > 0 Then
                startPos = 1
                cnt = Len("to")
                Do
                    pos = InStr(startPos, cellFormula, "to ", 1)
                    If pos > 0 Then
                        With .Characters(InStr(startPos, cellFormula, "to ", 1), cnt).Font

                            .color = RGB(64, 0, 128)
                            .Italic = True

                        End With
                    End If
                    startPos = pos + 1
                Loop While (pos > 0)
            End If
            If InStr(1, cellFormula, "sum") > 0 Then
                startPos = 1
                cnt = Len("sum")
                Do
                    pos = InStr(startPos, cellFormula, "sum ", 1)
                    If pos > 0 Then
                        With .Characters(InStr(startPos, cellFormula, "sum ", 1), cnt).Font

                            .color = RGB(64, 0, 128)
                            .Italic = True

                        End With
                    End If
                    startPos = pos + 1
                Loop While (pos > 0)
            End If
            If InStr(1, cellFormula, "average") > 0 Then
                startPos = 1
                cnt = Len("average")
                Do
                    pos = InStr(startPos, cellFormula, "average ", 1)
                    If pos > 0 Then
                        With .Characters(InStr(startPos, cellFormula, "average ", 1), cnt).Font

                            .color = RGB(64, 0, 128)
                            .Italic = True

                        End With
                    End If
                    startPos = pos + 1
                Loop While (pos > 0)
            End If
            If InStr(1, cellFormula, "less than") > 0 Then
                startPos = 1
                cnt = Len("less than")
                Do
                    pos = InStr(startPos, cellFormula, "less than ", 1)
                    If pos > 0 Then
                        With .Characters(InStr(startPos, cellFormula, "less than ", 1), cnt).Font

                            .color = RGB(64, 0, 128)
                            .Italic = True

                        End With
                    End If
                    startPos = pos + 1
                Loop While (pos > 0)
                
            End If
            If InStr(1, cellFormula, "plus") > 0 Then
                startPos = 1
                cnt = Len("plus")
                Do
                    pos = InStr(startPos, cellFormula, "plus ", 1)
                    If pos > 0 Then
                        With .Characters(InStr(startPos, cellFormula, "plus ", 1), cnt).Font

                            .color = RGB(64, 0, 128)
                            .Italic = True

                        End With
                    End If
                    startPos = pos + 1
                Loop While (pos > 0)
                End If
                If InStr(1, cellFormula, "multiplied by") > 0 Then
                startPos = 1
                cnt = Len("multiplied")
                Do
                    pos = InStr(startPos, cellFormula, "multiplied by ", 1)
                    If pos > 0 Then
                        With .Characters(InStr(startPos, cellFormula, "multiplied by ", 1), cnt).Font

                            .color = RGB(64, 0, 128)
                            .Italic = True

                        End With
                    End If
                    startPos = pos + 1
                Loop While (pos > 0)
                End If
                If InStr(1, cellFormula, "divided by") > 0 Then
                startPos = 1
                cnt = Len("plus")
                Do
                    pos = InStr(startPos, cellFormula, "divided by ", 1)
                    If pos > 0 Then
                        With .Characters(InStr(startPos, cellFormula, "divided by ", 1), cnt).Font

                            .color = RGB(64, 0, 128)
                            .Italic = True

                        End With
                    End If
                    startPos = pos + 1
                Loop While (pos > 0)
                End If
        End With
                

'Application.ScreenUpdating = True
    End With
    Next oCell
End Sub
Public Sub DomainVisualizeSingleFormulaCell()

  Dim rCell As Range
    Dim rRng As Range
    Dim strippedCellAddressString As String
    Dim CellAddressString As String
    Dim cellAddressTokens() As String
    Dim activeCellAddressString As String
    Dim resultRange As String
    Dim precedentRange As Range
    Dim precedentCell As Range
    Dim LRandomNumber As Integer
    
    If ActiveCell.HasFormula Then
    
        Dim cellAdd As String
        cellAdd = ActiveCell.Offset(0, 0).Address(True, True)
        Dim cellAddTokens() As String
        cellAddTokens() = Split(cellAdd, "$")
        Dim rowNumber As Integer
        rowNumber = CInt(cellAddTokens(2))
        Dim colNumber As Integer
        colNumber = ColRef2ColNo(cellAddTokens(1))
        
        Dim cellFormula As String
        cellFormula = ActiveCell.Formula
            
        
        
        'go column-wise
        Dim colHeaderString As String
        'colHeaderString = colHeader(rowNumber, colNumber)
        'MsgBox "column header " & colHeaderString
        
        'go row-wise
        Dim rowHeaderString As String
        'rowHeaderString = RowHeader(rowNumber, colNumber)
        'MsgBox "row header" & rowHeaderString
        Dim precedentsRange As Range, cel As Range
        
        Dim directPrecedentsString As String
        Dim cellHeader As String
        
        
        On Error Resume Next
        If Cells(rowNumber, colNumber).DirectPrecedents.Count Then
            Cells(rowNumber, colNumber).DirectPrecedents.Cells
            For Each cel In Cells(rowNumber, colNumber).DirectPrecedents.Cells
                'MsgBox "cel  " & cel.Address(False, False)
                cellAdd = cel.Address(True, True)
                cellAddTokens() = Split(cellAdd, "$")
                rowNumber = CInt(cellAddTokens(2))
                colNumber = ColRef2ColNo(cellAddTokens(1))
                colHeaderString = colHeader(rowNumber, colNumber)
                'MsgBox "column header " & colHeaderString
                rowHeaderString = RowHeader(rowNumber, colNumber)
                'MsgBox "row header" & rowHeaderString
                cellHeader = rowHeaderString & "' " & colHeaderString
                cellFormula = Replace(cellFormula, cel.Address(False, False), cellHeader)
                cellFormula = Replace(cellFormula, ":", " to ")
                'MsgBox cellFormula
                'MsgBox cellAdd
                'Range(cel.Address(False, False)).Interior.color = 20
                directPrecedentsString = directPrecedentsString & "," & cel.Address(False, False)
            Next
            
            'MsgBox Cells(rowNumber, colNumber).Precedents.Count & " dependancies found."
        Else
            MsgBox "No dependencies found."
        End If
        
        'MsgBox "Equivalent Formula is " & cellFormula
        ActiveCell.ClearComments
        ActiveCell.AddComment cellFormula
        With ActiveCell.Comment
            .Shape.TextFrame.AutoSize = True
            
        End With
    Else
    
        MsgBox "Cell has no formula"
        
    End If
    
    
End Sub

Sub SelectedRange()


    Dim rCell As Range
    Dim rRng As Range
    Dim strippedCellAddressString As String
    Dim CellAddressString As String
    Dim cellAddressTokens() As String
    Dim activeCellAddressString As String
    Dim resultRange As String
    Dim precedentRange As Range
    Dim precedentCell As Range
    

    Dim LRandomNumber As Integer
    
    Dim cellAdd As String
    cellAdd = ActiveCell.Offset(0, 0).Address(True, True)
    Dim cellAddTokens() As String
    cellAddTokens() = Split(cellAdd, "$")
    Dim rowNumber As Integer
    rowNumber = CInt(cellAddTokens(2))
    Dim colNumber As Integer
    colNumber = ColRef2ColNo(cellAddTokens(1))
    
    Dim cellFormula As String
    cellFormula = ActiveCell.Formula
    
    MsgBox cellFormula
    
    
    
    'go column-wise
    Dim colHeaderString As String
    colHeaderString = colHeader(rowNumber, colNumber)
    MsgBox "column header " & colHeaderString
    
    'go row-wise
    Dim rowHeaderString As String
    rowHeaderString = RowHeader(rowNumber, colNumber)
    MsgBox "row header" & rowHeaderString
    Dim precedentsRange As Range, cel As Range
    
    Dim directPrecedentsString As String
    Dim cellHeader As String
    
    
    On Error Resume Next
    If Cells(rowNumber, colNumber).DirectPrecedents.Count Then
        Cells(rowNumber, colNumber).DirectPrecedents.Cells
        For Each cel In Cells(rowNumber, colNumber).DirectPrecedents.Cells
            MsgBox "cel  " & cel.Address(False, False)
            cellAdd = cel.Address(True, True)
            cellAddTokens() = Split(cellAdd, "$")
            rowNumber = CInt(cellAddTokens(2))
            colNumber = ColRef2ColNo(cellAddTokens(1))
            colHeaderString = colHeader(rowNumber, colNumber)
            MsgBox "column header " & colHeaderString
            rowHeaderString = RowHeader(rowNumber, colNumber)
            MsgBox "row header" & rowHeaderString
            cellHeader = colHeaderString & rowHeaderString
            cellFormula = Replace(cellFormula, cel.Address(False, False), cellHeader)
            MsgBox cellHeader
            MsgBox cellAdd
            Range(cel.Address(False, False)).Interior.color = 20
            directPrecedentsString = directPrecedentsString & "," & cel.Address(False, False)
        Next
        
        MsgBox Cells(rowNumber, colNumber).Precedents.Count & " dependancies found."
    Else
        MsgBox "No dependancies found."
    End If
    
    MsgBox "Replaced Cell Formula is " & cellFormula
    
    
    
    

    

    
    
    Dim UsedRng As Range
    Dim FirstRow As Long, LastRow As Long, FirstCol As Long, LastCol As Integer
    
    Set UsedRng = ActiveSheet.UsedRange
     
    FirstRow = UsedRng(1).Row
    FirstCol = UsedRng(1).Column
    LastRow = UsedRng(UsedRng.Cells.Count).Row
    LastCol = UsedRng(UsedRng.Cells.Count).Column
     
    'MsgBox "First used row is: " & FirstRow
    'MsgBox "First used column is: " & FirstCol
    'MsgBox "Last used row is: " & LastRow
    'MsgBox "Last used column is: " & LastCol

    'MsgBox Selection.Address(False, False)
    'Set rRange = ActiveSheet.UsedRange.SpecialCells _
     '(xlCellTypeConstants, xlTextValues)
     Set rRange = ActiveSheet.UsedRange.SpecialCells _
     (xlCellTypeFormulas)
     
     'MsgBox "wawaw" & rRange
     
     For Each rCell In rRange

        'rCell.Interior.ColorIndex = 50
        'MsgBox "rrr" & rCell.Address

     Next rCell
     
     
    
    Dim mySelectedRangeString As String
    
    mySelectedRangeString = Selection.Address(False, False)
    
    'MsgBox "rrr" & mySelectedRangeString
    
    Dim mySelectedRange As Range
    
    Set mySelectedRange = ActiveSheet.Range(mySelectedRangeString)
    
    Dim myCell As Range
    Dim color As Integer
    
    Set rRange = mySelectedRange
    
    For Each myCell In rRange
            LRandomNumber = Int((56 - 2 + 1) * Rnd + 2)
            color = LRandomNumber
    
            'Debug.Print rCell.Address, rCell.Value
            'myCell.ShowPrecedents (True)
            'myCell.ShowPrecedents
            
            cellAddressTokens() = Split(myCell.Address, "$")
            'MsgBox (cellAddressTokens(2))
            CellAddressString = myCell.Address
            
            'MsgBox "tta" + CellAddressString
            strippedCellAddressString = Application.WorksheetFunction.Substitute(CellAddressString, "$", "")
        
        
            activeCellAddressString = strippedCellAddressString
            'MsgBox (activeCellAddressString)
            
            
            Dim result1 As Long
            
            'result1 = cellAddressTokens(2) + 1
            
            result1 = FirstRow '4/2/2013
            
            
            'resultRange = ActiveCell + LastRow
            'MsgBox "result 1" & result1
            
            Dim result2 As String
            
            Dim firstCellColumnwise As String
            
            firstCellColumnwise = cellAddressTokens(1) & CStr(result1)
            
            'MsgBox "First cell columnwise " & firstCellColumnwise
            
            Dim lastCellColumnwise As String
            
            'lastCellColumnwise = cellAddressTokens(1) & CStr(LastRow) '04/02/2012
            lastCellColumnwise = cellAddressTokens(1) & cellAddressTokens(2)
            
            'MsgBox "Last cell columnwise " & lastCellColumnwise
            
            Dim myRangeString As String
            
            myRangeString = lastCellColumnwise + ":" + firstCellColumnwise
            
            MsgBox myRangeString
            
            
            Set rRng = ActiveSheet.Range(myRangeString).End(xlUp)
        
            For Each rCell In rRng.Cells
                'Debug.Print rCell.Address, rCell.Value
                rCell.ShowPrecedents (True)
                rCell.ShowPrecedents
                'MsgBox "wawaw" & rCell.Precedents
                'If (rCell.HasFormula) Then
                'Set precedentRange = rCell.DirectPrecedents
                
                '    For Each precedentCell In precedentRange.Cells
                        
                '            precedentCell.Interior.ColorIndex = color
                        
                        'MsgBox "sddd" & precedentCell.Address
                 '   Next precedentCell
                'End If
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                If Not IsEmpty(rCell) Then
                    If Application.IsText(rCell) Then
                        MsgBox "Text" & rCell.Text
                    End If
                    
                    If IsNumeric(rCell) Then
                        MsgBox "Text" & rCell.Text
                        If rCell.HasFormula Then
                            MsgBox "Formula" & rCell.Text
                        End If
                    End If
                    
                    
                    rCell.Interior.ColorIndex = color
                End If
                
            Next rCell
            
            
        
            Dim firstCellRowwise As String
            
            MsgBox "token" + cellAddressTokens(2)
            
            firstCellRowwise = cellAddressTokens(1) & CStr(result1)
            
            Dim columnNumber As Integer
            
            columnNumber = ColRef2ColNo(cellAddressTokens(1))
            
            'columnNumber = columnNumber + 1 ' 04/02/2012
            
            'columnNumber = columnNumber + 1 '04/02/2012
            
            Dim columnNumberString As String
            
            'columnNumberString = ColNo2ColRef(columnNumber)'04/02/2012
            
            
            'firstCellRowwise = columnNumberString & cellAddressTokens(2)'04/02/2012
            
            firstCellRowwise = "A" & cellAddressTokens(2)
            
            MsgBox "First cell rowwise " & firstCellRowwise
            
            
            Dim lastCellRowwise As String
            
            'lastCellRowwise = CStr(ColNo2ColRef(LastCol)) & cellAddressTokens(2) '04/02/2012
            
            lastCellRowwise = cellAddressTokens(1) & cellAddressTokens(2)
            
            'MsgBox "Last cell rowwise " & lastCellRowwise
            
            Dim myRowRangeString As String
            
            myRowRangeString = firstCellRowwise + ":" + lastCellRowwise
            
            MsgBox "Row" & myRowRangeString
            
            'MsgBox ColNo2ColRef(LastCol)
            
            Set rRng = ActiveSheet.Range(myRowRangeString)
        
            For Each rCell In rRng.Cells
                'Debug.Print rCell.Address, rCell.Value
                'rCell.ShowPrecedents (True)
                'rCell.ShowPrecedents
                If Not IsEmpty(rCell) Then
                
                    If Application.IsText(rCell) Then
                        MsgBox "row" & rCell.Text
                    End If
                    
                    If IsNumeric(rCell) Then
                        MsgBox "row" & rCell.Text
                        If rCell.HasFormula Then
                            MsgBox "rowFormula" & rCell.Text
                        End If
                    End If
                    rCell.Interior.ColorIndex = color
                End If
                
                
                

                
                
            Next rCell
                
            
        Next myCell






End Sub


Sub RemoveAllShapes()
    Dim shp As Shape
                
                For Each shp In ActiveSheet.Shapes
                
                    If shp.AutoShapeType = _
                        msoShapeRectangle Then
                        shp.Delete
                    End If
                
                Next
End Sub




Sub Remove_All_Comments_From_Worksheet()

Cells.ClearComments

End Sub








