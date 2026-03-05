Attribute VB_Name = "modSubProcess"
Option Explicit

'*************
'データの統合
'*************
Public Sub subInputData()
    Dim startFilePathPos As Long, endFilePathPos As Long
    Dim searchSheetName As String, targetSheetName As String
    
    With wbInfo
        'ファイルパスを取得
        If Trim(.wsMain.Range("oplusFilePath").Value) = "" Then
            Call modCommonProcess.sub_msg_error(1, "読込対象ファイルパス")
        Else
            .filePathTarget = Trim(.wsMain.Range("oplusFilePath").Value)
        End If
        
        '現在のシフト表を削除
        .wsMain.Range(.wsMain.Cells(.rowStartShift, .colStartShift), _
                        .wsMain.Cells(.rowEndShift, .colEndShift)).ClearContents
        
        .wsMain.Range(.wsMain.Cells(.rowStartShift, .colStartShift), _
                        .wsMain.Cells(.rowEndShift, .colEndShift)).Borders.LineStyle = xlLineStyleNone
        
        .wsMain.Range(.wsMain.Cells(.rowStartShift, .colStartShift), _
                        .wsMain.Cells(.rowEndShift, .colEndShift)).Borders(xlEdgeTop).LineStyle = xlLineStyleNone
        
        .wsMain.Range(.wsMain.Cells(.rowStartShift, .colStartShift), _
                        .wsMain.Cells(.rowEndShift, .colEndShift)).Font.FontStyle = "標準"
        
        'ファイル名を取得
        searchSheetName = Range("oplusFilePath").Value
        startFilePathPos = InStrRev(searchSheetName, "\")
        endFilePathPos = InStrRev(searchSheetName, ".")
        targetSheetName = Mid(searchSheetName, startFilePathPos + 1, endFilePathPos - startFilePathPos - 1)
        
        'シフト表CSVを開く
        Set .wbTarget = Workbooks.Open(.filePathTarget)
        Set .wsTarget = .wbTarget.Worksheets(targetSheetName)
        
        '対象の月を取得
        .targetMonth = Split(.wsTarget.Cells(1, 3).Value, "/")(0)
        
        'シフト表最終行の取得
        .lastRowOplusShift = .wsTarget.Cells(Rows.Count, 1).End(xlUp).Row
        
        'シフト表最終列の取得
        .lastColOplusShift = .wsTarget.Cells(1, Columns.Count).End(xlToLeft).Column
        
        'メインシートにコピーペースト
        .wsTarget.Range(.wsTarget.Cells(2, 1), .wsTarget.Cells(.lastRowOplusShift, .lastColOplusShift)).Copy
        .wsMain.Range("targetPaste").PasteSpecial Paste:=xlPasteValues
        
        'oplusアウトプットファイル閉じる
        .wbTarget.Close savechanges:=False
        
    End With
    
End Sub

'*********
'体裁変更
'*********
'シフト表
Public Sub changeShiftForm()
        
    With wbInfo
        .rowEndShift = .wsMain.Cells(Rows.Count, .colStartShift).End(xlUp).Row
        
        .wsMain.Range(.wsMain.Cells(.rowEndShift + 1, 1), _
                        .wsMain.Cells(Rows.Count, Columns.Count)).ClearContents
        
        .wsMain.Range(.wsMain.Cells(.rowEndShift + 1, 1), _
                        .wsMain.Cells(Rows.Count, Columns.Count)).Interior.ColorIndex = xlNone
        
        .wsMain.Range(.wsMain.Cells(.rowEndShift + 1, 1), _
                        .wsMain.Cells(Rows.Count, Columns.Count)).Borders.LineStyle = xlLineStyleNone
        
        .wsMain.Range(.wsMain.Cells(.rowStartShift, .colStartShift), _
                        .wsMain.Cells(.rowEndShift, .colEndShift)).Borders.LineStyle = xlContinuous
        
        .wsMain.Range(.wsMain.Cells(.rowStartShift, .colStartShift), _
                        .wsMain.Cells(.rowEndShift, .colEndShift)).BorderAround Weight:=xlMedium
        
        .wsMain.Range(.wsMain.Cells(.rowStartShift, .colStartShift + 1), _
                        .wsMain.Cells(.rowEndShift, .colStartShift + 1)).Borders(xlEdgeRight).Weight = xlMedium
        
    End With
    
End Sub

'集計行
Public Sub changeAggregationForm()
    With wbInfo
        .wsMain.Range(.wsMain.Cells(.rowStartAggre, .colStartAggre), _
                        .wsMain.Cells(.rowEndShift, .colEndAggre)).Borders.LineStyle = xlLineStyleNone
        
        .wsMain.Range(.wsMain.Cells(.rowStartAggre, .colStartAggre), _
                        .wsMain.Cells(.rowEndShift, .colEndAggre)).Borders.LineStyle = xlContinuous
        
        .wsMain.Range(.wsMain.Cells(.rowStartAggre, .colStartAggre), _
                        .wsMain.Cells(.rowEndShift, .colEndAggre)).BorderAround Weight:=xlMedium
        
    End With
End Sub

'*************
'シフトの集計
'*************
Public Sub shiftAggregation()
    With wbInfo
        .wsMain.Range(.wsMain.Cells(.rowStartAggre, .colStartAggre), _
                        .wsMain.Cells(.rowStartAggre, .colEndAggre)).Copy
        
        .wsMain.Range(.wsMain.Cells(.rowStartAggre, .colStartAggre), _
                        .wsMain.Cells(.rowEndShift, .colEndAggre)).PasteSpecial
        
    End With
End Sub

'*********************
'ファイルの出力・保存
'*********************
'既存ファイル
Public Sub dataOutput()
    Dim s As Shape
    Dim sheetCount As Long
    Dim saveasAnswer As Integer
    Dim openFilePath As String
    Dim searchWorkSheetName As String
    Dim myWorkSheet As Worksheet
    
    With wbInfo
        'シフト表シートを既存ブックに保存
        Select Case Range("createPosition").Value
            Case 0
                MsgBox "シート作成場所を指定してください。", vbInformation
                Exit Sub
            Case 1
                openFilePath = Range("saveFilePath").Value
                Workbooks.Open filename:=openFilePath
                
                Set .wbOutput = ActiveWorkbook
                
                .wsMain.Copy before:=.wbOutput.Worksheets(1)
                
                searchWorkSheetName = .wbOutput.Worksheets(1).Range("outputDay").Value
                
                For Each myWorkSheet In Worksheets
                    If myWorkSheet.Name = searchWorkSheetName Then
                        saveasAnswer = MsgBox("「" & searchWorkSheetName & "」と同名のシートがあります。" _
                                        & vbCrLf & "上書きしますか。", vbYesNo)
                    End If
                Next
                
                If saveasAnswer = 0 Then
                    Worksheets(1).Name = searchWorkSheetName
                ElseIf saveasAnswer = 6 Then
                    Worksheets(searchWorkSheetName).Delete
                    Worksheets(1).Name = searchWorkSheetName
                ElseIf saveasAnswer = 7 Then
                    Worksheets(1).Delete
                    MsgBox "処理を中止します。" & vbCrLf & "再度実行してください。"
                    Exit Sub
                End If
            Case 2
                openFilePath = Range("saveFilePath").Value
                Workbooks.Open filename:=openFilePath
                
                Set .wbOutput = ActiveWorkbook
                
                sheetCount = .wbOutput.Worksheets.Count
                .wsMain.Copy after:=.wbOutput.Worksheets(sheetCount)
                
                searchWorkSheetName = .wbOutput.Worksheets(sheetCount + 1).Range("outputDay").Value
                
                For Each myWorkSheet In Worksheets
                    If myWorkSheet.Name = searchWorkSheetName Then
                        saveasAnswer = MsgBox("「" & searchWorkSheetName & "」と同名のシートがあります。" _
                                            & vbCrLf & "上書きしますか。", vbYesNo)
                    End If
                Next
                
                If saveasAnswer = 0 Then
                    Worksheets(sheetCount + 1).Name = searchWorkSheetName
                ElseIf saveasAnswer = 6 Then
                    Worksheets(searchWorkSheetName).Delete
                    Worksheets(sheetCount).Name = searchWorkSheetName
                ElseIf saveasAnswer = 7 Then
                    Worksheets(sheetCount + 1).Delete
                    MsgBox "処理を中止します。" & vbCrLf & "再度実行してください。"
                    Exit Sub
                End If
            
        End Select
        
        '体制調整（ボタンオブジェクト削除）
        For Each s In .wbOutput.Worksheets(1).Shapes
            s.Delete
        Next s
        
        '体裁調整（列削除）
        .wbOutput.Worksheets(1).Range("1:8").Delete
        
        'シートの上書き保存
        .wbOutput.Close savechanges:=True
        
    End With
End Sub

'別ファイル
Public Sub dataAnotherFileOutput()
    Dim s As Shape
    Dim filePath  As String, targetFileName As String
    
    With wbInfo
        'シフト表シートを新規のブックとして保存
        Call modCommonProcess.saveFileOther(SHEETNAMECREATE)
        
        '体裁調整（ボタンオブジェクト削除）
        For Each s In .wbOutput.Worksheets(1).Shapes
            s.Delete
        Next s
        
        '体裁調整（列削除）
        .wbOutput.Worksheets(1).Range("1:8").Delete
        
        'シート名の変更（yyyy年mm月）
        .wbOutput.Worksheets(1).Name = Range("outputDay").Value
        
        filePath = Application.GetSaveAsFilename(filefilter:="Excelブック(*.xlsx),*.xlsx")
        
        If Not filePath = "False" Then
            .wbOutput.SaveAs filename:=filePath, FileFormat:=xlWorkbookDefault
            .wbOutput.Close savechanges:=False
        Else
            MsgBox "処理を中止します。", vbInformation, "処理中止"
            .wbOutput.Saved = True
            .wbOutput.Close
        End If
    End With
End Sub

'*********
'情報削除
'*********
'カレンダー
Public Sub clearCalendar()
    With wbInfo
        .wsMain.Range(.wsMain.Cells(.rowStartCalendar, .colStartCalendar), _
                        .wsMain.Cells(.rowEndCalendar, .colEndCalendar)).ClearContents
                        
        .wsMain.Range(.wsMain.Cells(.rowStartCalendar, .colStartCalendar), _
                        .wsMain.Cells(.rowEndCalendar, .colEndCalendar)).Interior.ColorIndex = xlNone
        
        .wsMain.Range(.wsMain.Cells(.rowStartCalendar, .colStartCalendar), _
                        .wsMain.Cells(.rowEndCalendar, .colEndCalendar)).Font.Color = RGB(0, 0, 0)
        
    End With
    
End Sub

'作業内容
Public Sub clearWork()
    With wbInfo
        .wsMain.Range(.wsMain.Cells(.rowStartCalendar + 2, .colStartCalendar), _
                        .wsMain.Cells(.rowEndCalendar, .colEndCalendar)).ClearContents
        
        .wsMain.Range(.wsMain.Cells(.rowStartCalendar + 2, .colStartCalendar), _
                        .wsMain.Cells(Rows.Count, .colEndCalendar)).Interior.ColorIndex = xlNone
        
        .wsMain.Range(.wsMain.Cells(.rowStartCalendar + 3, .colStartCalendar), _
                        .wsMain.Cells(.rowEndCalendar, .colEndCalendar)).Borders.ColorIndex = xlNone
        
        .wsMain.Range(.wsMain.Cells(.rowStartCalendar + 3, .colStartCalendar), _
                        .wsMain.Cells(.rowEndCalendar, .colEndCalendar)).Borders.Weight = xlThin
        
        .wsMain.Range(.wsMain.Cells(.rowStartCalendar + 3, .colStartCalendar), _
                        .wsMain.Cells(.rowEndCalendar, .colEndCalendar)).BorderAround Weight:=xlMedium
        
    End With
End Sub

'シフト表
Public Sub clearShift()
    With wbInfo
        If .rowEndShift > Range("targetPaste").Row Then
            .wsMain.Range(.wsMain.Cells(.rowStartShift, .colStartShift), _
                            .wsMain.Cells(.rowEndShift, .colEndShift)).ClearContents
            
            .wsMain.Range(.wsMain.Cells(.rowStartShift, .colStartShift), _
                            .wsMain.Cells(.rowEndShift, .colEndShift)).Interior.ColorIndex = xlNone
            
            .wsMain.Range(.wsMain.Cells(.rowStartShift, .colStartShift), _
                            .wsMain.Cells(.rowEndShift, .colEndShift)).Borders.ColorIndex = xlNone
            
            .wsMain.Range(.wsMain.Cells(.rowStartShift, .colStartShift), _
                            .wsMain.Cells(.rowEndShift, .colEndShift)).Borders.Weight = xlThin
            
            .wsMain.Range(.wsMain.Cells(.rowStartShift, .colStartShift), _
                            .wsMain.Cells(.rowEndShift, .colEndShift)).BorderAround Weight:=xlMedium
            
           .wsMain.Range(.wsMain.Cells(.rowStartShift, .colStartShift + 1), _
                            .wsMain.Cells(.rowEndShift, .colStartShift + 1)).Borders(xlEdgeRight).Weight = xlMedium
            
        End If
    End With
End Sub

'シフト表(職員名削除無)
Public Sub clearShiftStaffLeave()
    With wbInfo
        If .rowEndShift > Range("targetPaste").Row Then
            .wsMain.Range(.wsMain.Cells(.rowStartShift, .colStartShift + 2), _
                            .wsMain.Cells(.rowEndShift, .colEndShift)).ClearContents
            
            .wsMain.Range(.wsMain.Cells(.rowStartShift, .colStartShift + 2), _
                            .wsMain.Cells(.rowEndShift, .colEndShift)).Interior.ColorIndex = xlNone
            
            .wsMain.Range(.wsMain.Cells(.rowStartShift, .colStartShift + 2), _
                            .wsMain.Cells(.rowEndShift, .colEndShift)).Borders.ColorIndex = xlNone
            
            .wsMain.Range(.wsMain.Cells(.rowStartShift, .colStartShift + 2), _
                            .wsMain.Cells(.rowEndShift, .colEndShift)).Borders.Weight = xlThin
            
            .wsMain.Range(.wsMain.Cells(.rowStartShift, .colStartShift + 2), _
                            .wsMain.Cells(.rowEndShift, .colEndShift)).BorderAround Weight:=xlMedium
            
           .wsMain.Range(.wsMain.Cells(.rowStartShift, .colStartShift + 1), _
                            .wsMain.Cells(.rowEndShift, .colStartShift + 1)).Borders(xlEdgeRight).Weight = xlMedium
            
        End If
    End With
End Sub

'閉所情報反映
Public Sub reflectionClose()
    Dim c As Integer
    
    For c = 1 To 31
        If Cells(14, c + 3).Value Like "*閉所*" Then
            Range(Cells(14, c + 3), Cells(Rows.Count, c + 3)).Interior.Color = RGB(230, 230, 230)
        End If
    Next c
    
End Sub

'*********
'凡例変更
'*********
'シフト表
Public Sub roleChange()
    Dim r As Long, c As Long, e As Long
    Dim targetStr As String
    
    With wbInfo
        
        For c = .colStartCalendar To .colEndCalendar
            For r = .rowStartShift To .rowEndShift
                targetStr = .wsMain.Cells(r, c).Value
                
                '「凡例_シフト」シートより変換文字列検索
                For e = 2 To .lastRowShiftRole
                    If targetStr = "" Then
                        Exit For
                    ElseIf .wsShiftRole.Cells(e, .colTargetShiftRole).Value = targetStr Then
                        .wsMain.Cells(r, c).Value = .wsShiftRole.Cells(e, .colChangeShiftRole).Value
                        Exit For
                    End If
                Next e
            Next r
        Next c
    End With
End Sub

'*************
'予定の色付け
'*************
'シフト表
Public Sub changeShiftColoring()
    Dim r As Long, c As Long, e As Long
    Dim targetStr As String
    
    With wbInfo
        For c = .colStartCalendar To .colEndCalendar
            For r = .rowStartShift To .rowEndShift
                targetStr = .wsMain.Cells(r, c).Value
                
                '「凡例_シフト」シートより表示スタイル検索
                For e = 5 To .lastRowShiftRole
                    If targetStr = "" Then
                        Exit For
                    ElseIf .wsShiftRole.Cells(e, .colTargetShiftRole).Value = targetStr Then
                    '----------
                    'RGBの格納
                    '----------
                    '背景色(R)の格納
                        Select Case Cells(e, .colShiftRoleInteriorRed).Value
                            Case ""
                                .shiftRoleInteriorRed = 255
                            Case Else
                                .shiftRoleInteriorRed = Cells(e, .colShiftRoleInteriorRed).Value
                        End Select
                    '背景色(G)の格納
                        Select Case Cells(e, .colShiftRoleInteriorGreen).Value
                            Case ""
                                .shiftRoleInteriorGreen = 255
                            Case Else
                                .shiftRoleInteriorGreen = Cells(e, .colShiftRoleInteriorGreen).Value
                        End Select
                    '背景色(B)の格納
                        Select Case Cells(e, .colShiftRoleInteriorBlue).Value
                            Case ""
                                .shiftRoleInteriorBlue = 255
                            Case Else
                                .shiftRoleInteriorBlue = Cells(e, .colShiftRoleInteriorBlue).Value
                        End Select
                    '文字色(R)の格納
                        Select Case Cells(e, .colShiftRoleFontRed).Value
                            Case ""
                                .shiftRoleFontRed = 0
                            Case Else
                                .shiftRoleFontRed = Cells(e, .colShiftRoleFontRed).Value
                        End Select
                    '文字色(G)の格納
                        Select Case Cells(e, .colShiftRoleFontGreen).Value
                            Case ""
                                .shiftRoleFontGreen = 0
                            Case Else
                                .shiftRoleFontGreen = Cells(e, .colShiftRoleFontGreen).Value
                        End Select
                    '文字色(B)の格納
                        Select Case Cells(e, .colShiftRoleFontBlue).Value
                            Case ""
                                .shiftRoleFontBlue = 0
                            Case Else
                                .shiftRoleFontBlue = Cells(e, .colShiftRoleFontBlue).Value
                        End Select
                        
                    '---------------
                    '色付け
                    '---------------
                        Cells(r, c).Interior.Color = _
                            RGB(.shiftRoleInteriorRed, .shiftRoleInteriorGreen, .shiftRoleInteriorBlue)
                        Cells(r, c).Font.Color = _
                            RGB(.shiftRoleFontRed, .shiftRoleFontGreen, .shiftRoleFontBlue)
                        
                        Select Case Cells(e, .colShiftRoleFontStyle).Value
                            Case "太字"
                                Cells(r, c).Font.FontStyle = "太字"
                            Case ""
                                Cells(r, c).Font.FontStyle = "標準"
                        End Select
                        
                        Exit For
                    End If
                Next e
            Next r
        Next c
    End With
End Sub

'作業内容
Public Sub changeWorkColoring()
    Dim r As Long, c As Long, e As Long
    Dim targetStr As String
    
    With wbInfo
        For c = .colStartCalendar To .colEndCalendar
            For r = .rowStartCalendar + 3 To .rowEndCalendar
                targetStr = .wsMain.Cells(r, c).Value
                
                '「凡例_作業内容」シートより表示スタイル検索
                For e = 5 To .lastRowWorkRole
                    If targetStr = "" Then
                        Exit For
                    ElseIf .wsWorkRole.Cells(e, .colSearchRole).Value = targetStr Then
                    '----------
                    'RGBの格納
                    '----------
                    '背景色(R)の格納
                        Select Case Cells(e, .colWorkRoleInteriorRed).Value
                            Case ""
                                .workRoleInteriorRed = 255
                            Case Else
                                .workRoleInteriorRed = Cells(e, .colWorkRoleInteriorRed).Value
                        End Select
                    '背景色(G)の格納
                        Select Case Cells(e, .colWorkRoleInteriorGreen).Value
                            Case ""
                                .workRoleInteriorGreen = 255
                            Case Else
                                .workRoleInteriorGreen = Cells(e, .colWorkRoleInteriorGreen).Value
                        End Select
                    '背景色(B)の格納
                        Select Case Cells(e, .colWorkRoleInteriorBlue).Value
                            Case ""
                                .workRoleInteriorBlue = 255
                            Case Else
                                .workRoleInteriorBlue = Cells(e, .colWorkRoleInteriorBlue).Value
                        End Select
                    '文字色(R)の格納
                        Select Case Cells(e, .colWorkRoleFontRed).Value
                            Case ""
                                .workRoleFontRed = 0
                            Case Else
                                .workRoleFontRed = Cells(e, .colWorkRoleFontRed).Value
                        End Select
                    '文字色(G)の格納
                        Select Case Cells(e, .colWorkRoleFontGreen).Value
                            Case ""
                                .workRoleFontGreen = 0
                            Case Else
                                .workRoleFontGreen = Cells(e, .colWorkRoleFontGreen).Value
                        End Select
                    '文字色(B)の格納
                        Select Case Cells(e, .colWorkRoleFontBlue).Value
                            Case ""
                                .workRoleFontBlue = 0
                            Case Else
                                .workRoleFontBlue = Cells(e, .colWorkRoleFontBlue).Value
                        End Select
                        
                    '---------------
                    '色付け
                    '---------------
                        Cells(r, c).Interior.Color = _
                            RGB(.workRoleInteriorRed, .workRoleInteriorGreen, .workRoleInteriorBlue)
                        Cells(r, c).Font.Color = _
                            RGB(.workRoleFontRed, .workRoleFontGreen, .workRoleFontBlue)
                        
                        Select Case Cells(e, .colWorkRoleFontStyle).Value
                            Case "太字"
                                Cells(r, c).Font.FontStyle = "太字"
                            Case ""
                                Cells(r, c).Font.FontStyle = "標準"
                        End Select
                        
                        Exit For
                    End If
                Next e
            Next r
        Next c
    End With

End Sub

'***********
'予定表作成
'***********
Public Sub inputSceduleThisMonth()
    Dim tarYear As Integer, tarMonth As Integer, i As Integer
    Dim maxDay As Long
    Dim searchDay As String
    Dim dataArea As Range
    
    tarYear = Range("targetYear").Value
    tarMonth = Range("targetMonth").Value
    
    Range("outputDay").Value = tarYear & "年" & tarMonth & "月"
    maxDay = DateSerial(tarYear, tarMonth + 1, 1)
    maxDay = DAY(maxDay - 1)
    
    Set dataArea = Worksheets("祝日一覧").Range("B2").CurrentRegion
    
    For i = 1 To maxDay
    
        Range("startPosition").Offset(0, i - 1).Value = i
        searchDay = DateSerial(tarYear, tarMonth, i)
        Range("startPosition").Offset(1, i - 1).Value = WeekdayName(Weekday(searchDay), True)
        
        '曜日によりフォントの色を変更する
        Select Case Weekday(searchDay)
            Case 1
                Range("startPosition").Offset(0, i - 1).Resize(2, 1).Interior.Color = RGB(255, 199, 206)
                Range("startPosition").Offset(0, i - 1).Resize(2, 1).Font.Color = RGB(156, 0, 6)
            Case 7
                Range("startPosition").Offset(0, i - 1).Resize(2, 1).Interior.Color = RGB(211, 235, 247)
                Range("startPosition").Offset(0, i - 1).Resize(2, 1).Font.Color = RGB(0, 112, 192)
            Case Else
                Range("startPosition").Offset(0, i - 1).Resize(2, 1).Font.Color = RGB(0, 0, 0)
        End Select
        
        '不要な日付の削除
        Select Case i
            Case maxDay
                If i = 28 Then
                    Range("startPosition").Offset(0, i).Resize(2, 3).ClearContents
                ElseIf i = 29 Then
                    Range("startPosition").Offset(0, i).Resize(2, 2).ClearContents
                ElseIf i = 30 Then
                    Range("startPosition").Offset(0, i).Resize(2, 1).ClearContents
                End If
        End Select
        
        '祝日の場合フォントの色を変更する
        Select Case WorksheetFunction.CountIf(dataArea.Columns(1), searchDay)
            Case 1
                Range("startPosition").Offset(0, i - 1).Resize(2, 1).Interior.Color = RGB(255, 199, 206)
                Range("startPosition").Offset(0, i - 1).Resize(2, 1).Font.Color = RGB(156, 0, 6)
        End Select
        
    Next i
    
End Sub
