'Copyright 2018 Benson Mariro and Bennett Kankuzi, Dept of Computer Science, North-West University, South Africa.
'Code released under a Creative Commons Attribution 4.0 International Licence.

Option Explicit
Option Compare Text

Dim previouslyActiveCellAddress As String
Dim ActiveCellAddress As String
Dim previouslyActiveCell As Range
Dim myActiveCellAddress As String
Dim myActiveCellPrecedentsRange As Range
Dim targetChanged As Range
Dim colourIndexArray() As Variant
Dim indexint As Integer
Dim editedComment As String
Dim uneditedComment As String
Dim strippedEditedComment As String
Dim strippedunEditedComment As String
Dim rawEditedComment As String
Dim formulaSystemChanged As Boolean
Dim uneditedCommentPrecedents As Range
Dim myCellEditedComment As Range
Dim gCell As Range
Dim rowSearchNotErrorDetect As Integer
Dim colSearchNotErrorDetect As Integer
Dim gObject As Range
Dim resultCell As Range
Dim It_Is_A_Number As Boolean
Dim findTempRange As Range
Dim findTempRange2 As Range
Dim findAddress As String
Dim foundCount As Integer
Dim iRet As Integer
Dim strPrompt As String
Dim strTitle As String
Dim highlightedColumnLabelCell As Range

Private Sub Workbook_Open()

    Set previouslyActiveCell = Range("A1")
    Set highlightedColumnLabelCell = Nothing
    indexint = 1
    editedComment = ""
    uneditedComment = ""
    rawEditedComment = ""
    formulaSystemChanged = False
    
    'Dim usedCell As Range
    'On Error Resume Next
    'For Each usedCell In ActiveSheet.UsedRange
        'If usedCell.Interior.ColorIndex = 28 Then
           'usedCell.Interior.ColorIndex = xlNone
        'End If
    'Next
    
End Sub



Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
        
        Dim myComment As Object
        Set myComment = ActiveCell.Comment
        If myComment Is Nothing Then
        Else
            ActiveCell.Comment.Visible = False
        End If
        
        Application.EnableEvents = False
        
        Dim myCell As Range
        
        On Error Resume Next
        If ActiveCell.DirectPrecedents.Count > 0 Then
            
            'Set rRange = Range(previouslyActiveCellAddress).DirectPrecedents
            
            On Error Resume Next
            For Each myCell In ActiveCell.DirectPrecedents.Cells
                
                myCell.Interior.ColorIndex = xlNone
            Next
              
        End If
        
        Application.EnableEvents = True
        
    
End Sub



Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    
    Application.DisplayCommentIndicator = xlNoIndicator
    
    formulaSystemChanged = False
    
    ' If the whole row or column has been selected, stop further processing
    Dim myCols As Integer
    myCols = Selection.Columns.Count
    
    If myCols >= 16384 Then
        'MsgBox myCols
        Exit Sub
    End If
    
    Dim myRows As Long
    myRows = Selection.Rows.Count
    
    If myRows >= 1048576 Then
        'MsgBox myRows
        Exit Sub
    End If
    
    ' End - If the whole row or column has been selected, stop further processing
    
    ' we hereafter, do further processing
    
    If previouslyActiveCellAddress = "" Then
    
        'MsgBox previouslyActiveCellAddress
        Set previouslyActiveCell = Range("A1")
        previouslyActiveCellAddress = Trim(previouslyActiveCell.Offset(0, 0).Address(False, False))
    
        Exit Sub
    
    End If
     
    
    'MsgBox previouslyActiveCell.Address
    'Set previouslyActiveCell = Range(previouslyActiveCellAddress)
    On Error Resume Next
    previouslyActiveCellAddress = Trim(previouslyActiveCell.Offset(0, 0).Address(False, False))
    Set previouslyActiveCell = Range(previouslyActiveCellAddress)
    ActiveCellAddress = Trim(ActiveCell.Offset(0, 0).Address(False, False))
    
    Dim myCell As Range
    Dim rRange As Range
    Dim myComment As Object
     
    With Range(previouslyActiveCellAddress)
        
        Set myComment = Range(previouslyActiveCellAddress).Comment
        If myComment Is Nothing Then
        Else
            Range(previouslyActiveCellAddress).Comment.Visible = False
            'MsgBox Range(previouslyActiveCellAddress).Comment.Text
            editedComment = Range(previouslyActiveCellAddress).Comment.Text
            'MsgBox "edi" & editedComment
            
            Dim i
            Dim differenceText As String
            differenceText = ""
            
            
            strippedEditedComment = Replace(editedComment, " ", " ")
            strippedunEditedComment = Replace(uneditedComment, " ", " ")
            rawEditedComment = editedComment
            'MsgBox "len " & Len(strippedEditedComment) & " " & strippedEditedComment
            'MsgBox "len 2 " & Len(strippedunEditedComment) & " " & strippedunEditedComment
            
            'MsgBox "last" & Application.ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
            
            If Len(editedComment) <> Len(uneditedComment) Or (editedComment <> uneditedComment) Then
            
        
            
            Dim strippedEditedCommentSplit As Variant
            Dim intIndex As Integer
            
            strippedEditedComment = Replace(strippedEditedComment, "|", "~|")
            strippedEditedComment = Replace(strippedEditedComment, "+", "~")
            strippedEditedComment = Replace(strippedEditedComment, "-", "~")
            strippedEditedComment = Replace(strippedEditedComment, "*", "~")
            strippedEditedComment = Replace(strippedEditedComment, "/", "~")
            strippedEditedComment = Replace(strippedEditedComment, "(", "~")
            strippedEditedComment = Replace(strippedEditedComment, ")", "~")
            strippedEditedComment = Replace(strippedEditedComment, "...", "~")
            strippedEditedComment = Replace(strippedEditedComment, ",", "~")
            strippedEditedComment = Replace(strippedEditedComment, "<", "~")
            strippedEditedComment = Replace(strippedEditedComment, ">", "~")
            
            
            strippedEditedCommentSplit = Split(strippedEditedComment, "~")
            
            Dim intRow As Long
            intRow = 0
            'Dim searchRange As Range
            'searchRange = ActiveSheet.UsedRange
            Dim cellAddStripped As String
            Dim cellAddTokensStripped() As String
            Dim rowNumberStripped As String
            Dim colNumberStripped As String
            
            
            For intIndex = LBound(strippedEditedCommentSplit) To UBound(strippedEditedCommentSplit)
                'MsgBox "Item " & intIndex & " is " & strippedEditedCommentSplit(intIndex) & _
                '" which is " & Len(strippedEditedCommentSplit(intIndex)) & " characters long", vbInformation
                
                If Len(strippedEditedCommentSplit(intIndex)) = 0 Then
                    'Do Nothing
                Else
                    If strippedEditedCommentSplit(intIndex) Like "*|*" Then
                        'MsgBox "row"
                        strippedEditedCommentSplit(intIndex) = Replace(strippedEditedCommentSplit(intIndex), "|", "")
                        'MsgBox "strip" & strippedEditedCommentSplit(intIndex)
                        'Application.ActiveSheet.UsedRange
                        'cellAddStripped = Range("B5:B15").Find(Trim(strippedEditedCommentSplit(intIndex)), Range("B2"), xlValues, xlWhole, xlByRows, xlNext).Address(True, True)
                        'Set gObject = Range("B5:B18").Find(Trim(strippedEditedCommentSplit(intIndex)), Range("B2"), xlValues, xlWhole, xlRows)
                        'ActiveSheet.UsedRange.Select
                        
                        
                        
                        On Error Resume Next
                        
                        With Range(Cells(1, 1), Cells(50, 50))
                            
                            foundCount = 0
                            Set findTempRange = .Find(What:=Trim(strippedEditedCommentSplit(intIndex)), After:=.Cells(1, 1), LookIn:=xlValues, LookAt:= _
                                            xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                                        , SearchFormat:=False)
                            'Set findTempRange = .Find(Trim(strippedEditedCommentSplit(intIndex)))
                             'If a match is found then
                            If Not findTempRange Is Nothing Then
                                 'Store the address of the cell where the first match is found in a variable
                                 'MsgBox "found count" & foundCount
                                foundCount = foundCount + 1
                                findAddress = findTempRange.Address
                                Do
                                    'MsgBox "found count inside" & foundCount
                                     'Color the cell where a match is found yellow
                                    'findTempRange.Interior.ColorIndex = 6
                                     'Search for the next cell with a matching value
                                    Set findTempRange = .FindNext(findTempRange)
                                    'MsgBox "Found" & foundCount
                                     'Search for all the other occurrences of the item i.e.
                                     'Loop as long matches are found, and the address of the cell where a match is found,
                                     'is different from the address of the cell where the first match is found (FindAddress)
                                     foundCount = foundCount + 1
                                Loop While Not findTempRange Is Nothing And findTempRange.Address <> findAddress
                            End If
                            
                            
                            If foundCount > 2 Then
                                Set findTempRange = .Find(What:=Trim(strippedEditedCommentSplit(intIndex)), After:=.Cells(1, 1), LookIn:=xlValues, LookAt:= _
                                                xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                                            , SearchFormat:=False)
                                'Set findTempRange = .Find(Trim(strippedEditedCommentSplit(intIndex)))
                                 'If a match is found then
                                If Not findTempRange Is Nothing Then
                                     'Store the address of the cell where the first match is found in a variable
                                    findAddress = findTempRange.Address
                                    Do
                                         
                                         'Color the cell where a match is found yellow
                                        findTempRange.Interior.ColorIndex = 6
                                        
                                         'Search for the next cell with a matching value
                                         strPrompt = "There are multiple occurences of " & """" & findTempRange.Value & """" & ". Is the highlighted " & """" & findTempRange.Value & """" & " the one you want?" & _
                                         " Your formula is: (" & editedComment & ")"

 
                                       ' Dialog's Title
                                       strTitle = "Label Occuring Multiple Times"
                                    
                                       'Display MessageBox
                                       iRet = MsgBox(strPrompt, vbYesNo, strTitle)
                                    
                                       ' Check pressed button
                                        If iRet = vbYes Then
                                            'MsgBox "Yes!"
                                            cellAddStripped = findTempRange.Address(True, True)
                                            Set gObject = findTempRange
                                            findTempRange.Interior.ColorIndex = xlNone
                                            If highlightedColumnLabelCell Is Nothing Then
                                                'do nothing
                                            Else
                                                highlightedColumnLabelCell.Interior.ColorIndex = xlNone
                                            End If
                                            Exit Do
                                        Else
                                            cellAddStripped = ""
                                            Set gObject = Nothing
                                            'MsgBox "No!"
                                        End If
                                        findTempRange.Interior.ColorIndex = xlNone
                                        Set findTempRange = .FindNext(findTempRange)
                                        'MsgBox "Found Col" & foundCount
                                         'Search for all the other occurrences of the item i.e.
                                         'Loop as long matches are found, and the address of the cell where a match is found,
                                         'is different from the address of the cell where the first match is found (FindAddress)
                                    Loop While Not findTempRange Is Nothing And findTempRange.Address <> findAddress
                                End If
                            Else
                                'Set gObject = .Find(What:=Trim(strippedEditedCommentSplit(intIndex)), After:=.Cells(1, 1), LookIn:=xlValues, LookAt:= _
                                                xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                                            , SearchFormat:=False)
                                'cellAddStripped = .Find(What:=Trim(strippedEditedCommentSplit(intIndex)), After:=.Cells(1, 1), LookIn:=xlValues, LookAt:= _
                                            xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                                            , SearchFormat:=False).Address(True, True)
                                Set gObject = findTempRange
                                cellAddStripped = findTempRange.Address(True, True)
                                
                            End If
                            
                                
                        End With
            
            
            
                        If gObject Is Nothing Then
                            MsgBox "There is no label " & """" & strippedEditedCommentSplit(intIndex) & """" & ". Please correct accordingly.", vbExclamation
                            rowSearchNotErrorDetect = 1
                        Else
                            'MsgBox "Found" & strippedEditedCommentSplit(intIndex)
                            rowSearchNotErrorDetect = 0
                        End If
                        'MsgBox "ad" & cellAddStripped
                        'For Each gCell In ActiveSheet.UsedRange
                            'If gCell.Text Like strippedEditedCommentSplit(intIndex) Then
                                'MsgBox "Found" & strippedEditedCommentSplit(intIndex)
                                'rowSearchNotErrorDetect = 0
                                'Exit For
                            'Else
                                'cellAddStripped = "~$~$~"
                                'MsgBox "Not Found" & cellAddStripped
                                
                            'End If
                        'Next
                        
                        
                        If rowSearchNotErrorDetect = 0 Then
                            'Do Nothing
                            'MsgBox cellAddStripped
                            cellAddTokensStripped() = Split(cellAddStripped, "$")
                            rowNumberStripped = cellAddTokensStripped(2)
                            'MsgBox rowNumberStripped & "waza"
                            rawEditedComment = Replace(rawEditedComment, strippedEditedCommentSplit(intIndex), rowNumberStripped, 1, 1)
                            'MsgBox rawEditedComment
                        Else
                            'cellAddTokensStripped() = Split(cellAddStripped, "$")
                            'rowNumberStripped = cellAddTokensStripped(2)
                            
                            rowNumberStripped = " " & strippedEditedCommentSplit(intIndex)
                            'MsgBox rowNumberStripped
                            rawEditedComment = Replace(rawEditedComment, strippedEditedCommentSplit(intIndex), rowNumberStripped, 1, 1)
                            
                        End If
                    Else
                        'cellAddStripped = ActiveSheet.UsedRange.Find(Trim(strippedEditedCommentSplit(intIndex)), Range("B2"), xlValues, xlWhole, xlByRows, xlNext).Address(True, True)
                        
                        'Set gObject = ActiveSheet.UsedRange.Find(Trim(strippedEditedCommentSplit(intIndex)), Range("B2"), xlValues, xlWhole, xlByRows)
                        
                        On Error Resume Next

                        With Range(Cells(1, 1), Cells(50, 50))
                            foundCount = 0
                            Set findTempRange2 = .Find(What:=Trim(strippedEditedCommentSplit(intIndex)), After:=.Cells(1, 1), LookIn:=xlValues, LookAt:= _
                                            xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                                        , SearchFormat:=False)
                            'Set findTempRange = .Find(Trim(strippedEditedCommentSplit(intIndex)))
                             'If a match is found then
                            If Not findTempRange2 Is Nothing Then
                                 'Store the address of the cell where the first match is found in a variable
                                foundCount = foundCount + 1
                                findAddress = findTempRange2.Address
                                Do
                                     
                                     'Color the cell where a match is found yellow
                                    'findTempRange.Interior.ColorIndex = 8
                                     'Search for the next cell with a matching value
                                    Set findTempRange2 = .FindNext(findTempRange2)
                                    'MsgBox "Found Col" & foundCount
                                     'Search for all the other occurrences of the item i.e.
                                     'Loop as long matches are found, and the address of the cell where a match is found,
                                     'is different from the address of the cell where the first match is found (FindAddress)
                                     foundCount = foundCount + 1
                                Loop While Not findTempRange2 Is Nothing And findTempRange2.Address <> findAddress
                            End If
                            
                            If foundCount > 2 Then
                                Set findTempRange2 = .Find(What:=Trim(strippedEditedCommentSplit(intIndex)), After:=.Cells(1, 1), LookIn:=xlValues, LookAt:= _
                                                xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                                            , SearchFormat:=False)
                                'Set findTempRange = .Find(Trim(strippedEditedCommentSplit(intIndex)))
                                 'If a match is found then
                                If Not findTempRange2 Is Nothing Then
                                     'Store the address of the cell where the first match is found in a variable
                                    findAddress = findTempRange2.Address
                                    Do
                                         
                                         'Color the cell where a match is found green
                                         If highlightedColumnLabelCell Is Nothing Then
                                                'do nothing
                                            Else
                                                highlightedColumnLabelCell.Interior.ColorIndex = xlNone
                                            End If
                                            
                                        findTempRange2.Interior.ColorIndex = 10
                                        
                                         'Search for the next cell with a matching value
                                         strPrompt = "There are multiple occurences of " & """" & findTempRange2.Value & """" & ". Is the highlighted " & """" & findTempRange2.Value & """" & " the one you want?" & _
                                         " Your formula is : (" & editedComment & ")"
 
                                       ' Dialog's Title
                                       strTitle = "Label Occuring Multiple Times"
                                    
                                       'Display MessageBox
                                       iRet = MsgBox(strPrompt, vbYesNo, strTitle)
                                    
                                       ' Check pressed button
                                        If iRet = vbYes Then
                                            'MsgBox "Yes!"
                                            cellAddStripped = findTempRange2.Address(True, True)
                                            Set gObject = findTempRange2
                                            Set highlightedColumnLabelCell = findTempRange2
                                            'findTempRange2.Interior.ColorIndex = xlNone
                                            Exit Do
                                        Else
                                            cellAddStripped = ""
                                            Set gObject = Nothing
                                            'MsgBox "No!"
                                        End If
                                        findTempRange2.Interior.ColorIndex = xlNone
                                        Set findTempRange2 = .FindNext(findTempRange2)
                                        'MsgBox "Found Col" & foundCount
                                         'Search for all the other occurrences of the item i.e.
                                         'Loop as long matches are found, and the address of the cell where a match is found,
                                         'is different from the address of the cell where the first match is found (FindAddress)
                                    Loop While Not findTempRange2 Is Nothing And findTempRange2.Address <> findAddress
                                End If
                            Else
                                'Set gObject = .Find(What:=Trim(strippedEditedCommentSplit(intIndex)), After:=.Cells(1, 1), LookIn:=xlValues, LookAt:= _
                                                xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                                            , SearchFormat:=False)
                                'cellAddStripped = .Find(What:=Trim(strippedEditedCommentSplit(intIndex)), After:=.Cells(1, 1), LookIn:=xlValues, LookAt:= _
                                            xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                                            , SearchFormat:=False).Address(True, True)
                                Set gObject = findTempRange2
                                cellAddStripped = findTempRange2.Address(True, True)
                                
                            End If
                            
                            

                            
                    
                        End With
                        
                        If gObject Is Nothing Then
                            'MsgBox "There is no label " & """" & strippedEditedCommentSplit(intIndex) & """" & ". Please correct accordingly. The formula has not been changed!", vbExclamation
                            colSearchNotErrorDetect = 1
                            If Not (Trim(strippedEditedCommentSplit(intIndex)) = "SUM") And Not (Trim(strippedEditedCommentSplit(intIndex)) = "AVERAGE") And Not (Trim(strippedEditedCommentSplit(intIndex)) = "IF") Then

                                If Trim(strippedEditedCommentSplit(intIndex)) Like "*""*" Then
                                    'Do Nothing
                                Else
                                    Dim x As Integer
                                    It_Is_A_Number = True
                                    For x = 1 To Len(Trim(strippedEditedCommentSplit(intIndex)))
                                        If Asc(Mid(Trim(strippedEditedCommentSplit(intIndex)), x, 1)) > 57 Or Asc(Mid(Trim(strippedEditedCommentSplit(intIndex)), x, 1)) < 48 Then
                                            It_Is_A_Number = False
                                            Exit For
                                        End If
                                    Next x
                                    If It_Is_A_Number Then
                                        'MsgBox "It's a number."
                                        'Do Nothing
                                    Else
                                        MsgBox "There is no label " & """" & strippedEditedCommentSplit(intIndex) & """" & ". Please correct accordingly. The formula has not been changed!", vbExclamation
                                    End If
                                    
                                End If
                            End If
                            
                        Else
                            'MsgBox "Found" & strippedEditedCommentSplit(intIndex)
                            colSearchNotErrorDetect = 0
                            
                            If IsNumeric(Range(cellAddStripped).Value) Then
                             'MsgBox "numeric"
                             colSearchNotErrorDetect = 3
                            End If
                        End If
                        
                        If colSearchNotErrorDetect = 0 Then
                            'Do Nothing
                            'MsgBox cellAddStripped
                            cellAddTokensStripped() = Split(cellAddStripped, "$")
                            colNumberStripped = cellAddTokensStripped(1)
                            'MsgBox rowNumberStripped & "waza"
                            rawEditedComment = Replace(rawEditedComment, strippedEditedCommentSplit(intIndex), colNumberStripped, 1, 1)
                            'MsgBox rawEditedComment
                        Else
                            'cellAddTokensStripped() = Split(cellAddStripped, "$")
                            'rowNumberStripped = cellAddTokensStripped(2)
                            If colSearchNotErrorDetect = 3 Then
                                colNumberStripped = strippedEditedCommentSplit(intIndex)
                                rawEditedComment = Replace(rawEditedComment, strippedEditedCommentSplit(intIndex), colNumberStripped, 1, 1)
                            Else
                                If strippedEditedCommentSplit(intIndex) Like "*""*" Then
                                    'MsgBox "has"
                                    colNumberStripped = strippedEditedCommentSplit(intIndex)
                                    'MsgBox "fff" & strippedEditedCommentSplit(intIndex) & rawEditedComment
                                    'colNumberStripped = Replace(colNumberStripped, " ", "~")
                                    rawEditedComment = Replace(rawEditedComment, strippedEditedCommentSplit(intIndex), colNumberStripped, 1, 1)
                                Else
                                    colNumberStripped = Replace(strippedEditedCommentSplit(intIndex), " ", "")
                                    'MsgBox "fff" & strippedEditedCommentSplit(intIndex) & rawEditedComment
                                    'colNumberStripped = Replace(colNumberStripped, " ", "~")
                                    rawEditedComment = Replace(rawEditedComment, strippedEditedCommentSplit(intIndex), colNumberStripped, 1, 1)
                                End If
                            End If
                            
                        End If
                        
                        'If Len(cellAddStripped) = 0 Then
                            'Do Nothing
                        'Else
                            'MsgBox cellAddStripped
                            'cellAddTokensStripped() = Split(cellAddStripped, "$")
                            'colNumberStripped = cellAddTokensStripped(1)
                            'MsgBox colNumberStripped
                            'rawEditedComment = Replace(rawEditedComment, strippedEditedCommentSplit(intIndex), colNumberStripped)
                            'MsgBox rawEditedComment
                        'End If
                    End If
                End If
                
            Next ' intIndex = LBound(strippedEditedCommentSplit)
            
            
            rawEditedComment = "=" & Replace(rawEditedComment, "|", "")
            rawEditedComment = Replace(rawEditedComment, "...", ":")
            
            'rawEditedComment = Replace(rawEditedComment, " ", " ")
            'rawEditedComment = Trim(rawEditedComment)
            'MsgBox rawEditedComment
            
            Application.EnableEvents = False
            
            formulaSystemChanged = True
            
            Set uneditedCommentPrecedents = Range(previouslyActiveCellAddress).DirectPrecedents.Cells
            'MsgBox "pre" & uneditedCommentPrecedents.Count
            On Error Resume Next
            If uneditedCommentPrecedents.Count > 0 Then
                
                'Set rRange = Range(previouslyActiveCellAddress).DirectPrecedents
                indexint = 0
                'Dim colorPreviousCellsEditedComment As Variant
                On Error Resume Next
                For Each myCell In uneditedCommentPrecedents.Cells
                    'myCell.Interior.ColorIndex = xlNone
                    'colorPreviousCells = colourIndexArray(indexint)
                    'if previous color was the comment light green colour (28), then change to none
                    
                    If colourIndexArray(indexint) = 28 Then
                        myCell.Interior.ColorIndex = xlNone
                    Else
                        myCell.Interior.ColorIndex = colourIndexArray(indexint)
                    End If
                    
                    indexint = indexint + 1
                Next
                  
            End If
            
            
            'Range(previouslyActiveCellAddress).ClearContents
            If highlightedColumnLabelCell Is Nothing Then
                'do nothing
            Else
                highlightedColumnLabelCell.Interior.ColorIndex = xlNone
            End If
            Range(previouslyActiveCellAddress).Formula = rawEditedComment
            
            'Range(previouslyActiveCellAddress).Formula = "=IF(A1<3000,""Small hjj"", ""Large jkkk"")"
            'Application.Run ("'DomainRealWorldVisualize.xlam'!DomainVisualizeAllColumnToRow")
            Application.EnableEvents = True
            
            'Range(previouslyActiveCellAddress).Formula = rawEditedComment
            
            End If ' If Len(strippedEditedComment) <> Len(strippedunEditedComment) Then
            
            
            
            'Application.Run ("'DomainRealWorldVisualize.xlam'!DomainVisualizeAllColumnToRow")
            'Range("B18").Formula = "=IF(A1<3000,""Small hjj"", ""Large jkkk"")"
            
            'MsgBox "striEdi" & strippedEditedComment
            'MsgBox "striUnEdi" & strippedunEditedComment
            
            
            
            
        End If ' myComment Is Nothing
        
        Application.EnableEvents = False
        
        On Error Resume Next
        If Range(previouslyActiveCellAddress).DirectPrecedents.Count > 0 Then
            
            'Set rRange = Range(previouslyActiveCellAddress).DirectPrecedents
            indexint = 0
            'Dim colorPreviousCells As Variant
            On Error Resume Next
            For Each myCell In Range(previouslyActiveCellAddress).DirectPrecedents.Cells
                'myCell.Interior.ColorIndex = xlNone
                'colorPreviousCells = colourIndexArray(indexint)
                'if previous color was the comment light green colour (28), then change to none
                
                If colourIndexArray(indexint) = 28 Then
                    myCell.Interior.ColorIndex = xlNone
                Else
                    myCell.Interior.ColorIndex = colourIndexArray(indexint)
                End If
                
                indexint = indexint + 1
            Next
              
        End If
        
        
            
        
        ' if the previous cell is empty after deleting its contents,
        ' remove its corresponding comment if any
        
        If IsEmpty(Range(previouslyActiveCellAddress)) Then
            'MsgBox "nothing"
            Set myComment = Range(previouslyActiveCellAddress).Comment
            If myComment Is Nothing Then
            Else
                
                
                Range(previouslyActiveCellAddress).Borders(xlEdgeRight).LineStyle = XlLineStyle.xlLineStyleNone
                Range(previouslyActiveCellAddress).ClearComments
                
            End If
            
        End If
        
        Application.EnableEvents = True
        
    End With
    
    With Range(ActiveCellAddress)
        
        Set myComment = Range(ActiveCellAddress).Comment
        If myComment Is Nothing Then
        Else
            Range(ActiveCellAddress).Comment.Visible = True
            uneditedComment = Range(ActiveCellAddress).Comment.Text
            'MsgBox "un" & uneditedComment
        End If
        
        Application.EnableEvents = False
        
        On Error Resume Next
        If Range(ActiveCellAddress).DirectPrecedents.Count > 0 Then
            
            'Set rRange = Range(previouslyActiveCellAddress).DirectPrecedents
            Set myActiveCellPrecedentsRange = Range(ActiveCellAddress).DirectPrecedents
            myActiveCellAddress = ActiveCellAddress
            indexint = 0
            Dim precedentCurrentColor As Variant
            ReDim Preserve colourIndexArray(0 To Range(ActiveCellAddress).DirectPrecedents.Count)
            On Error Resume Next
            For Each myCell In Range(ActiveCellAddress).DirectPrecedents.Cells
                
                precedentCurrentColor = myCell.Interior.ColorIndex
                colourIndexArray(indexint) = CInt(precedentCurrentColor)
                indexint = indexint + 1
                myCell.Interior.ColorIndex = 28
            Next
            
        End If
        
        Application.EnableEvents = True
        
    End With
    
    
    Set previouslyActiveCell = ActiveCell
    
End Sub


Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)

    ' If the whole row or column has been selected, stop further processing
    
    'whole row selected
    
    Dim myCell As Range
    Dim myComment As Object
    
    Dim myCols As Integer
    myCols = Selection.Columns.Count
    
    If myCols >= 16384 Then
        'MsgBox myCols
            Application.EnableEvents = False
                
                On Error Resume Next
                For Each myCell In myActiveCellPrecedentsRange.Cells
                    myCell.Interior.ColorIndex = xlNone
                    Set myComment = myCell.Comment
                    If myComment Is Nothing Then
                    Else
                        myCell.Borders(xlEdgeRight).LineStyle = XlLineStyle.xlLineStyleNone
                        myCell.ClearComments
                        
                    End If
                Next
                Set previouslyActiveCell = Range("A1")
                Application.EnableEvents = True
        Exit Sub
    End If
    
    'whole column selected
    Dim myRows As Long
    myRows = Selection.Rows.Count
    
    If myRows >= 1048576 Then
        'MsgBox myRows
        
        Application.EnableEvents = False
                
                On Error Resume Next
                For Each myCell In myActiveCellPrecedentsRange.Cells
                    myCell.Interior.ColorIndex = xlNone
                    Set myComment = myCell.Comment
                    If myComment Is Nothing Then
                    Else
                        myCell.Borders(xlEdgeRight).LineStyle = XlLineStyle.xlLineStyleNone
                        myCell.ClearComments
                    End If
                Next
                Set previouslyActiveCell = Range("A1")
                Application.EnableEvents = True
        Exit Sub
    End If
    
    ' End - If the whole row or column has been selected, stop further processing

    'making sure that once a group of cells has been deleted, the corresponding
    'highlighted precedents are de-highlighted
    
    Dim aCell As Range
   
    For Each aCell In Target.Cells '-- Target may contains more than one cells.
        If aCell.Formula = "" Then
            'MsgBox "Cell " & aCell.Address & " in " & Sh.Name & " has been cleared or deleted."
            'MsgBox "My active cell address" & myActiveCellAddress
            'MsgBox " acell.address" & aCell.Address
            If myActiveCellAddress = aCell.Offset(0, 0).Address(False, False) Then
                'MsgBox "Same"
                Application.EnableEvents = False
                'Dim myCell As Range
                On Error Resume Next
                For Each myCell In myActiveCellPrecedentsRange.Cells
                    
                    myCell.Interior.ColorIndex = xlNone
                Next
                Application.EnableEvents = True
                
            End If
            
        Else
            'MsgBox "Cell " & aCell.Address & " in " & Sh.Name & " has been changed."
        End If
    Next
   
   
   ' Executing the domain visualization to regenerate all the comments once  changes
   ' have been made to the spreadsheet - dont regenerate for number inputs
    'If Not (IsNumeric(Target.Value)) Or Target.Formula = "" Then
    
    Dim cel As Range
    Dim dcel As Range
    Dim directDependents As Range
    
    If Target.Cells.Count >= 1 Then
        For Each cel In Target
            'MsgBox cel.Address(False, False)
            On Error Resume Next
            If cel.directDependents.Count > 0 Then
                Set directDependents = cel.directDependents
                For Each dcel In directDependents
                    'MsgBox "direct" & dcel.Address(False, False)
                Next
            End If 'cel.DirectDependents.Count > 0
        Next
    End If 'Target.Cells.Count >= 1
        'MsgBox Target
    
    If Target.Cells.Count = 1 Then
        'MsgBox Target
    
        If IsNumeric(Target) Then
            'do nothing
            If Target.HasFormula Or Target.Formula = "" Then
                If Target.Formula = "" Then
                    Application.Run ("'DomainRealWorldVisualize.xlam'!DomainVisualizeAllColumnToRow")
                Else 'Target.HasFormula
                'MsgBox "wawwa"
                    'For Each cel In Target
                    'GetTargetChanged (cel)
                    If formulaSystemChanged = True Then
                        'MsgBox "system changed"
                    Else
                        Application.Run ("'DomainRealWorldVisualize.xlam'!DomainVisualizeAllColumnToRow")
                    End If
                    'Next
                End If
                
            End If 'Target.HasFormula
        Else
            Application.Run ("'DomainRealWorldVisualize.xlam'!DomainVisualizeAllColumnToRow")
        End If ' IsNumeric(Target)
            
    Else
        If Target.Cells.Count > 1 Then
            Application.Run ("'DomainRealWorldVisualize.xlam'!DomainVisualizeAllColumnToRow")
        End If ' Target.Cells.Count > 1
    End If 'Target.Cells.Count = 1
    
    
        'Dim ResultState As Integer
        
        'ResultState = Application.Run("'DomainRealWorldVisualize.xlam'!ReturnState")
        'ResultState = 2 'make column to row as default
        'If ResultState = 1 Then
            'Application.Run ("'DomainRealWorldVisualize.xlam'!DomainVisualizeAllRowToColumn")
        'Else
            'If ResultState = 2 Then
                'run the default column to row
                'Application.Run ("'DomainRealWorldVisualize.xlam'!DomainVisualizeAllColumnToRow")
            'Else
                
                    'Application.Run ("'DomainRealWorldVisualize.xlam'!DomainVisualizeAllRowToColumn")
            'End If
        'End If
    
    
    
End Sub


































