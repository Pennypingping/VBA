' PART 1: Helper Function (This part is unchanged)
' =====================================================================
Function IsTableUniform_BasedOnUserLogic(ByVal tbl As Table, ByRef columnCount As Long) As Boolean
    Dim maxColumns As Long
    Dim rowCount As Long
    Dim totalCellCountInTable As Long
    Dim oRow As Row

    columnCount = 0
    On Error GoTo HandleError

    maxColumns = 0
    For Each oRow In tbl.Rows
        If oRow.Cells.Count > maxColumns Then
            maxColumns = oRow.Cells.Count
        End If
    Next oRow

    rowCount = tbl.Rows.Count
    totalCellCountInTable = tbl.Range.Cells.Count
    
    If totalCellCountInTable = (rowCount * maxColumns) Then
        IsTableUniform_BasedOnUserLogic = True
        columnCount = maxColumns
    Else
        IsTableUniform_BasedOnUserLogic = False
    End If
    
    Exit Function

HandleError:
    IsTableUniform_BasedOnUserLogic = False
End Function


' =====================================================================
' PART 2: Main Subroutine (This is the updated macro)
' =====================================================================
Sub 表格修改2()
    Dim tbl As Table
    Dim cel As Cell
    Dim i As Long, j As Long, k As Long
    Dim currentColCount As Long
    
    Dim skippedComplexTableCount As Long, skippedSmallTableCount As Long
    Dim firstColHasText As Boolean
    Dim shouldDelete As Boolean
    Dim isCompletelyBlank As Boolean
    Dim cellText As String
    Dim tableWasModified As Boolean
    Dim blankPara As Range, nextItem As Range
    Dim doNotDelete As Boolean
    Dim textToCheck As String
    Dim prefixes As Variant, prefix As Variant
    
    Dim isSpecialCol As Boolean

    ' Initialize counters
    skippedComplexTableCount = 0
    skippedSmallTableCount = 0
    
    Application.ScreenUpdating = False

    For i = ActiveDocument.Tables.Count To 1 Step -1
        Set tbl = ActiveDocument.Tables(i)
        
        If IsTableUniform_BasedOnUserLogic(tbl, currentColCount) Then
            
            '==================================================
            ' START OF STEP 1: Deletion Logic (Unchanged)
            '==================================================
            tableWasModified = False
            If currentColCount >= 6 Then
                firstColHasText = False
                For Each cel In tbl.Columns(1).Cells
                    cellText = Trim(Replace(cel.Range.Text, vbCr & Chr(7), ""))
                    If cellText <> "" Then firstColHasText = True: Exit For
                Next cel
                If firstColHasText Then
                    For j = tbl.Columns.Count To 1 Step -1
                        shouldDelete = True
                        isCompletelyBlank = True
                        If tbl.Rows.Count > 1 Then
                            For k = 2 To tbl.Rows.Count
                                Set cel = tbl.Cell(k, j)
                                cellText = Trim(Replace(cel.Range.Text, vbCr & Chr(7), ""))
                                If cellText <> "-" And cellText <> "" Then shouldDelete = False: Exit For
                                If cellText <> "" Then isCompletelyBlank = False
                            Next k
                        Else
                            shouldDelete = False
                        End If
                        If shouldDelete And Not isCompletelyBlank Then
                            If j < tbl.Columns.Count Then
                                tbl.Columns(j).Delete
                                tbl.Columns(j).Delete
                            Else
                                tbl.Columns(j).Delete
                            End If
                            tableWasModified = True
                        End If
                    Next j
                End If
                If tableWasModified Then
                    For j = 1 To tbl.Columns.Count
                        isCompletelyBlank = True
                        For Each cel In tbl.Columns(j).Cells
                            cellText = Trim(Replace(cel.Range.Text, vbCr & Chr(7), ""))
                            If cellText <> "" Then isCompletelyBlank = False: Exit For
                        Next cel
                        If isCompletelyBlank Then
                            tbl.Columns(j).PreferredWidth = CentimetersToPoints(0.26)
                        End If
                    Next j
                End If
            Else
                skippedSmallTableCount = skippedSmallTableCount + 1
            End If
            '==================================================
            ' END OF STEP 1
            '==================================================


            '==================================================
            ' START OF STEP 2: Final Formatting
            '==================================================
            Dim finalColCount As Long
            If IsTableUniform_BasedOnUserLogic(tbl, finalColCount) Then
                
                tbl.AllowAutoFit = False
                tbl.Columns.PreferredWidthType = wdPreferredWidthPoints
                On Error Resume Next
                
                '==================================================
                ' *** NEW REPLACEMENT LOGIC STARTS HERE ***
                ' This block applies to ALL tables that enter the Select Case below.
                '==================================================
                
                ' 1. Replace all single spaces with nothing.
                With tbl.Range.Find
                    .ClearFormatting
                    .Replacement.ClearFormatting
                    .Text = " "
                    .Replacement.Text = ""
                    .Forward = True
                    .Wrap = wdFindStop
                    .Execute Replace:=wdReplaceAll
                End With
                
                ' 2. Replace opening parenthesis to add padding for negative numbers.
                With tbl.Range.Find
                    .ClearFormatting
                    .Replacement.ClearFormatting
                    .Text = "("
                    .Replacement.Text = "(  "
                    .Forward = True
                    .Wrap = wdFindStop
                    .Execute Replace:=wdReplaceAll
                End With
                
                '==================================================
                ' *** NEW REPLACEMENT LOGIC ENDS HERE ***
                '==================================================

                Select Case finalColCount
                    Case 6 To 8
                        tbl.Columns(1).Width = Application.CentimetersToPoints(8)
                        For Each cel In tbl.Columns(1).Cells
                            cel.Range.ParagraphFormat.CharacterUnitLeftIndent = 8
                        Next cel
                    Case 9
                        tbl.Columns(1).Width = Application.CentimetersToPoints(5)
                    Case 10
                        tbl.Columns(1).Width = Application.CentimetersToPoints(5)
                        For Each cel In tbl.Columns(1).Cells
                            cel.Range.ParagraphFormat.CharacterUnitLeftIndent = 2
                        Next cel
                        
                    Case 11 To 13
                        tbl.Range.Font.Size = 11
                        tbl.Columns.PreferredWidth = CentimetersToPoints(2.5)
                        For j = 1 To finalColCount
                            isSpecialCol = True
                            For k = 1 To tbl.Rows.Count
                                cellText = Trim(Replace(tbl.Cell(k, j).Range.Text, vbCr & Chr(7), ""))
                                If cellText <> "" And cellText <> "$" Then isSpecialCol = False: Exit For
                            Next k
                            If isSpecialCol Then tbl.Columns(j).PreferredWidth = CentimetersToPoints(0.15)
                        Next j

                    Case 14 To 15
                        tbl.Range.Font.Size = 10
                        tbl.Columns.PreferredWidth = CentimetersToPoints(2.2)
                        For j = 1 To finalColCount
                            isSpecialCol = True
                            For k = 1 To tbl.Rows.Count
                                cellText = Trim(Replace(tbl.Cell(k, j).Range.Text, vbCr & Chr(7), ""))
                                If cellText <> "" And cellText <> "$" Then isSpecialCol = False: Exit For
                            Next k
                            If isSpecialCol Then tbl.Columns(j).PreferredWidth = CentimetersToPoints(0.15)
                        Next j
                        
                    Case 16 To 18
                        tbl.Range.Font.Size = 8.5
                        tbl.Columns.PreferredWidth = CentimetersToPoints(1.9)
                        For j = 1 To finalColCount
                            isSpecialCol = True
                            For k = 1 To tbl.Rows.Count
                                cellText = Trim(Replace(tbl.Cell(k, j).Range.Text, vbCr & Chr(7), ""))
                                If cellText <> "" And cellText <> "$" Then isSpecialCol = False: Exit For
                            Next k
                            If isSpecialCol Then tbl.Columns(j).PreferredWidth = CentimetersToPoints(0.11)
                        Next j
                        
                    Case Is >= 19
                        tbl.Range.Font.Size = 8
                        For j = 1 To finalColCount
                            isSpecialCol = True
                            For k = 1 To tbl.Rows.Count
                                cellText = Trim(Replace(tbl.Cell(k, j).Range.Text, vbCr & Chr(7), ""))
                                If cellText <> "" And cellText <> "$" Then isSpecialCol = False: Exit For
                            Next k
                            If isSpecialCol Then tbl.Columns(j).PreferredWidth = CentimetersToPoints(0.1)
                        Next j
                        
                End Select
                On Error GoTo 0
            End If
            '==================================================
            ' END OF STEP 2
            '==================================================
            
        Else
            skippedComplexTableCount = skippedComplexTableCount + 1
        End If
        
    Next i
    
    Application.ScreenUpdating = True
    
    ' --- Final Report and Second Loop for deleting blank lines (Unchanged) ---
    Dim report As String
    report = "所有處理程序已完成！" & vbCrLf & vbCrLf
    report = report & "已成功處理所有符合條件的表格。" & vbCrLf
    If skippedComplexTableCount > 0 Then
        report = report & "有 " & skippedComplexTableCount & " 個複雜表格 (含合併儲存格) 被安全跳過。" & vbCrLf
    End If
    If skippedSmallTableCount > 0 Then
        report = report & "有 " & skippedSmallTableCount & " 個小型表格 (<6欄) 被安全跳過。" & vbCrLf
    End If
    MsgBox report, vbInformation, "處理報告"
    Application.ScreenUpdating = False
    For i = ActiveDocument.Tables.Count To 1 Step -1
        Set tbl = ActiveDocument.Tables(i)
        Set blankPara = tbl.Range.Next(Unit:=wdParagraph, Count:=1)
        If Not blankPara Is Nothing And Trim(Replace(blankPara.Text, vbCr, "")) = "" Then
            Set nextItem = blankPara.Next(Unit:=wdParagraph, Count:=1)
            If Not nextItem Is Nothing Then
                doNotDelete = False
                If nextItem.Tables.Count > 0 Then
                    doNotDelete = True
                Else
                    textToCheck = Trim(nextItem.Text)
                    prefixes = Array("(一)", "(二)", "(三)", "(四)", "(五)", "(六)", "(七)", "(八)", "(九)", "(十)", "(十一)", "(十二)", "(十三)", "(十四)", "(十五)", "(十六)", "(十七)", "(十八)", "(十九)", "(二十)", "(二十一)", "(二十二)", "(二十三)", "(二十四)", "(二十五)", "一、", "二、", "三、", "四、", "五、", "六、", "七、", "八、", "九、", "十、", "十一、", "十二、", "十三、", "十四、", "十五、", "十六、", "十七、", "十八、", "十九、", "二十、")
                    For Each prefix In prefixes
                        If textToCheck Like prefix & "*" Then
                            doNotDelete = True
                            Exit For
                        End If
                    Next prefix
                End If
                If Not doNotDelete Then
                    blankPara.Delete
                End If
            End If
        End If
    Next i
    Application.ScreenUpdating = True
    MsgBox "智慧型刪除空白行處理完成！", vbInformation, "提示"
End Sub

