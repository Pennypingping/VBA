Sub FormatTables_Conditionally()
    ' 宣告變數
    Dim oTbl As Word.Table
    Dim oCell As Word.Cell
    Dim colWidthPoints As Single
    Dim totalTableCount As Long
    Dim modifiedTableCount As Long

    ' --- 1. 定義要設定的數值 ---
    ' 設定第一欄的寬度為 10 公分
    colWidthPoints = CentimetersToPoints(10)

    ' 暫時忽略錯誤，以處理含有合併儲存格的複雜表格
    On Error Resume Next

    ' 檢查文件中是否有表格
    totalTableCount = ActiveDocument.Tables.Count
    If totalTableCount > 0 Then
        ' 初始化修改計數器
        modifiedTableCount = 0
        
        ' 遍歷文件中的每一個表格
        For Each oTbl In ActiveDocument.Tables

            ' --- 【新增的條件判斷】---
            ' 檢查表格的總欄數是否小於或等於 5
            If oTbl.Columns.Count <= 5 Then
            
                ' --- 動作一：設定表格整體格式 (寬度與縮排) ---
                ' 移除表格本身的左縮排，使其靠左對齊
                oTbl.Rows.LeftIndent = 0
                
                ' 設定第一欄的寬度為 10 公分
                oTbl.Columns(1).PreferredWidthType = wdPreferredWidthPoints
                oTbl.Columns(1).PreferredWidth = colWidthPoints
                
                ' --- 動作二：設定第一欄儲存格內部文字的縮排 ---
                ' 遍歷第一欄中的每一個儲存格
                For Each oCell In oTbl.Columns(1).Cells
                    ' 直接使用 CharacterUnitLeftIndent 屬性設定 6 個字元的縮排
                    oCell.Range.ParagraphFormat.CharacterUnitLeftIndent = 6
                Next oCell
                
                ' 如果條件符合並執行修改，計數器加 1
                modifiedTableCount = modifiedTableCount + 1
                
            End If
            ' 如果表格欄數超過 5，則直接跳到下一個表格

        Next oTbl
        
        ' 恢復正常的錯誤處理
        On Error GoTo 0
        
        ' 完成後顯示更詳細的成功訊息
        MsgBox "處理完成！" & vbCrLf & vbCrLf & _
               "共檢查 " & totalTableCount & " 個表格。" & vbCrLf & _
               "已修改 " & modifiedTableCount & " 個欄數小於或等於5的表格。", vbInformation

    Else
        ' 如果文件中沒有表格，則顯示提示訊息
        MsgBox "此文件中找不到任何表格。", vbInformation
    End If
End Sub
