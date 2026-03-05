Attribute VB_Name = "modVariable"
Option Explicit

    Public Const SHEETNAMECREATE      As String = "シフト表"
    Public Const SHEETNAMESHIFTROLE   As String = "凡例_シフト"
    Public Const SHEETNAMEWORKROLE    As String = "凡例_作業内容"
    Public Const SHEETNAMEPLANROLE    As String = "凡例_行事予定"
    Public Const SHEETNAMEHOLIDAY     As String = "祝日一覧"
    
    Public wbInfo As wbInfo
    
    Type wbInfo
        wb                            As Workbook
        wbTarget                      As Workbook
        wbOutput                      As Workbook
        wsTarget                      As Worksheet
        wsMain                        As Worksheet
        wsShiftRole                   As Worksheet
        wsWorkRole                    As Worksheet
        wsPlanRole                    As Worksheet
        wsEmployee                    As Worksheet
        wsHoliday                     As Worksheet
        
        fso                           As FileSystemObject
        filePathTarget                As String
        targetMonth                   As String
    
    '----------------------
    'メインシート情報
    '----------------------
    'シフト表_開始終了行
        rowStartShift                 As Long
        rowEndShift                   As Long
    'シフト表_開始終了列
        colStartShift                 As Long
        colEndShift                   As Long
    'カレンダー開始終了行
        rowStartCalendar              As Long
        rowEndCalendar                As Long
    'カレンダー開始終了列
        colStartCalendar              As Long
        colEndCalendar                As Long
    '集計表_開始終了行
        rowStartAggre                 As Long
        rowEndAggre                   As Long
    '集計表_開始終了列
        colStartAggre                 As Long
        colEndAggre                   As Long
    
    '----------------------
    'インプットシート情報
    '----------------------
    'oplusシフトデータ終了行
        lastRowOplusShift             As Long
    'oplusシフトデータ終了列
        lastColOplusShift             As Long
    
    '----------------------
    'シフト凡例シート情報
    '----------------------
    '最終行
        lastRowShiftRole              As Long
    'oplus
        colTargetShiftRole            As Long
    'excelシフト
        colChangeShiftRole            As Long
    'サンプル表示
        colShiftRoleTarget            As Long
    '背景色
        colShiftRoleInteriorRed       As Long
        colShiftRoleInteriorGreen     As Long
        colShiftRoleInteriorBlue      As Long
        shiftRoleInteriorRed          As Long
        shiftRoleInteriorGreen        As Long
        shiftRoleInteriorBlue         As Long
    '文字色
        colShiftRoleFontRed           As Long
        colShiftRoleFontGreen         As Long
        colShiftRoleFontBlue          As Long
        shiftRoleFontRed              As Long
        shiftRoleFontGreen            As Long
        shiftRoleFontBlue             As Long
    '文字スタイル
        colShiftRoleFontStyle         As Long
    
    '----------------------
    '作業内容凡例シート情報
    '----------------------
    '最終行
        lastRowWorkRole               As Long
    'Excel
        colSearchRole                 As Long
    'サンプル表示
        colWorkRoleTarget             As Long
    '背景色
        colWorkRoleInteriorRed        As Long
        colWorkRoleInteriorGreen      As Long
        colWorkRoleInteriorBlue       As Long
        workRoleInteriorRed           As Long
        workRoleInteriorGreen         As Long
        workRoleInteriorBlue          As Long
    '文字色
        colWorkRoleFontRed            As Long
        colWorkRoleFontGreen          As Long
        colWorkRoleFontBlue           As Long
        workRoleFontRed               As Long
        workRoleFontGreen             As Long
        workRoleFontBlue              As Long
    '文字スタイル
        colWorkRoleFontStyle          As Long
        
    '-------------------
    '祝日一覧シート情報
    '-------------------
    'ダウンロードファイルURL
        targetURL                 As String
    '保存先ファイルパス
        saveFilePath              As String
    'ダウンロード結果
        resultDownload            As Long
    
End Type
