Attribute VB_Name = "modVariable"
Option Explicit
    Public wbInfo As wbInfo
    
    Type wbInfo
        wb                            As Workbook
        wbTarget                      As Workbook
        wbHoliday                     As Workbook
        wsExport                      As Worksheet
        wsShift                       As Worksheet
        wsMaster                      As Worksheet
        wsCondition                   As Worksheet
        wsCheck                       As Worksheet
        wsHoliday                     As Worksheet
        wsTarget                      As Worksheet
        wsHolidayInput                As Worksheet
        
        formatConditionHoliday        As FormatCondition
        formatConditionWorkday        As FormatCondition
        formatConditionWorkTime       As FormatCondition
        formatConditionShiftNight     As FormatCondition
        formatConditionShiftHoli      As FormatCondition
        formatConditionShiftOver      As FormatCondition
        
    '作成対象月
        startDay                      As Date
        createDay                     As Date
        targetDay                     As Date
        targetMonth                   As String
    '検索ファイルパス
        targetFilePath                As String
        decodeURL                     As String
    'ファイル選択
        myFile                        As String
    
    '----------------------
    'フラグ情報
    '----------------------
    '変形労働適用単位
        flagTypeSystem                As Long
    '作成可否
        flagTypeCreate                As Long
    'ファイル選択
        flagSelectFile                As Long
    '作成対象日
        flagCreateDay                 As Long
        
    '----------------------
    '配列情報
    '----------------------
    '勤務日区分
        classWorkDay                  As Object
    '勤務区分
        classWorkType                 As Object
    '転記
        classWorkTranslate            As Object
        
    '----------------------
    'リシテア取込Excel
    '----------------------
    '項番
        rowTranslateNumber            As Long
        colTranslateNumber            As Long
    '氏名
        colTranslateName              As Long
    '個人CD
        colTranslateParsonalCode      As Long
        
    '----------------------
    'リシテア取込シート
    '----------------------
    'シフト
        rowExportShift                As Long
        rowExportStartShift           As Long
        rowExportEndShift             As Long
        colExportShift                As Long
    'カレンダー
        rowExportCalendar             As Long
        rowExportCalendarWeekday      As Long
        colExportStartCalendar        As Long
        colExportEndCalendar          As Long
        colExportCalendar             As Long
    '項番
        countExport                   As Long
    '氏名
        colExportName                 As Long
    '個人CD
        colExportParsonalCode         As Long
        
    '----------------------
    'シフト表シート
    '----------------------
    'シフト
        rowShiftRole                  As Long
        rowShiftStartRole             As Long
        rowShiftEndRole               As Long
    'カレンダー
        shiftCalendar                 As Date
        rowShiftCalendar              As Long
        rowShiftCalendarWeekday       As Long
        colShiftCalendar              As Long
        colShiftStartCalendar         As Long
        colShiftEndCalendar           As Long
        colEndCalendar                As Long
    '氏名
        colShiftName                  As Long
        shiftName                     As String
    '個人CD
        colShiftParsonalCode          As Long
    '労働時間
        colShiftWorkTime              As Long
    '労働日数
        colShiftWorkDay               As Long
    
    '----------------------
    '条件シート
    '----------------------
    '対象日
        rowConditionDay               As Long
        rowConditionStartDay          As Long
        rowConditionEndDay            As Long
        colConditionDay               As Long
        colConditionStartDay          As Long
        colConditionEndDay            As Long
    '所定労働時間（1日あたり）
        colConditionWorkTimePerday    As Long
        
    '----------------------
    '労働時間チェックシート
    '----------------------
    '対象月
        rowCheckMonth                 As Long
        rowCheckStartMonth            As Long
        rowCheckEndMonth              As Long
        colCheckMonth                 As Long
    '労働日数
        colCheckWorkDay               As Long
    '所定労働時間
        colCheckWorkTime              As Long
        
    '----------------------
    'マスタシート
    '----------------------
    '勤務区分凡例
        rowMasterEndShiftRole         As Long
    '内容
        colMasterShiftContent         As Long
        
    '----------------------
    '検索シート
    '----------------------
    'シフト表_開始終了行
        targetShift                   As String
        rowTargetShift                As Long
        rowTargetStartShift           As Long
        rowTargetEndShift             As Long
    'カレンダー開始終了列
        targetCalendar                As Date
        colTargetCalendar             As Long
        colTargetStartCalendar        As Long
        colTargetEndCalendar          As Long
    '土・祝振休判定
        targetDate                    As String
        targetWeekday                 As Long
        
    '----------------------
    '祝日一覧シート
    '----------------------
    'ダウンロードファイルURL
        targetURL                     As String
    '保存先ファイルパス
        activePath                    As String
    'ダウンロード結果
        resultDownload                As Long
    '開始終了行（社休日）
        rowHolidayListStart           As Long
        rowHolidayListEnd             As Long
    '開始終了列（社休日）
        colHolidayListStart           As Long
        colHolidayListEnd             As Long
    End Type
