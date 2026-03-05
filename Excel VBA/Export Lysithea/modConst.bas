Attribute VB_Name = "modConst"
Option Explicit
    '*****************
    'シート名
    '*****************
    Public Const SHEETNAMEEXPORT         As String = "リシテア取込"
    Public Const SHEETNAMESHIFT          As String = "シフト表"
    Public Const SHEETNAMEMASTER         As String = "マスタ"
    Public Const SHEETNAMECONDITION      As String = "条件"
    Public Const SHEETNAMECHECK          As String = "労働時間チェック"
    Public Const SHEETNAMEHOLIDAY        As String = "社休日一覧"
    Public Const SHEETNAMETARGET         As String = "勤務予定"
    '*****************
    'リシテア
    '*****************
    '-----------
    '勤務事由
    '-----------
    '休日
    Public Const CATEGORYHOLIDAY_1       As String = "社休(日曜)"
    Public Const CATEGORYHOLIDAY_2       As String = "社休(土祝)"
    '平日
    Public Const CATEGORYWEEKDAY         As String = "平日"
    '-----------
    '勤務形態
    '-----------
    '休日
    Public Const CATEGORYREFLECT_1       As String = "変形休"
    Public Const CATEGORYREFLECT_2       As String = "土祝振"
    Public Const CATEGORYREFLECT_3       As String = "日振"
    Public Const CATEGORYREFLECT_4       As String = "代休"
    Public Const CATEGORYREFLECT_5       As String = "年休"
    Public Const CATEGORYREFLECT_6       As String = "明"
    '1か月_日勤
    Public Const CATEGORYWORKDAY_M1      As String = "月変4定"
    Public Const CATEGORYWORKDAY_M2      As String = "月変6定"
    Public Const CATEGORYWORKDAY_M3      As String = "月変7定"
    Public Const CATEGORYWORKDAY_M4      As String = "月変7.75定"
    Public Const CATEGORYWORKDAY_M5      As String = "月変8定"
    Public Const CATEGORYWORKDAY_M6      As String = "月変8.5定"
    Public Const CATEGORYWORKDAY_M7      As String = "月変9定"
    Public Const CATEGORYWORKDAY_M8      As String = "月変9.5定"
    Public Const CATEGORYWORKDAY_M9      As String = "月変10定"
    Public Const CATEGORYWORKDAY_M10     As String = "月変11定"
    Public Const CATEGORYWORKDAY_M11     As String = "月変12定"
    Public Const CATEGORYWORKDAY_M12     As String = "月変16定"
    '1か月_夜勤
    Public Const CATEGORYWORKNIGHT_M1    As String = "月変4非"
    Public Const CATEGORYWORKNIGHT_M2    As String = "月変6非"
    Public Const CATEGORYWORKNIGHT_M3    As String = "月変7非"
    Public Const CATEGORYWORKNIGHT_M4    As String = "月変7.75非"
    Public Const CATEGORYWORKNIGHT_M5    As String = "月変8非"
    Public Const CATEGORYWORKNIGHT_M6    As String = "月変8.5非"
    Public Const CATEGORYWORKNIGHT_M7    As String = "月変9非"
    Public Const CATEGORYWORKNIGHT_M8    As String = "月変9.5非"
    Public Const CATEGORYWORKNIGHT_M9    As String = "月変10非"
    Public Const CATEGORYWORKNIGHT_M10   As String = "月変11非"
    Public Const CATEGORYWORKNIGHT_M11   As String = "月変12非"
    Public Const CATEGORYWORKNIGHT_M12   As String = "月変16非"
    '1年_日勤
    Public Const CATEGORYWORKDAY_Y1      As String = "年変4定"
    Public Const CATEGORYWORKDAY_Y2      As String = "年変6定"
    Public Const CATEGORYWORKDAY_Y3      As String = "年変7定"
    Public Const CATEGORYWORKDAY_Y4      As String = "年変7.75定"
    Public Const CATEGORYWORKDAY_Y5      As String = "年変8定"
    Public Const CATEGORYWORKDAY_Y6      As String = "年変8.5定"
    Public Const CATEGORYWORKDAY_Y7      As String = "年変9定"
    Public Const CATEGORYWORKDAY_Y8      As String = "年変9.5定"
    Public Const CATEGORYWORKDAY_Y9      As String = "年変10定"
    '1年_夜勤
    Public Const CATEGORYWORKNIGHT_Y1    As String = "年変4非"
    Public Const CATEGORYWORKNIGHT_Y2    As String = "年変6非"
    Public Const CATEGORYWORKNIGHT_Y3    As String = "年変7非"
    Public Const CATEGORYWORKNIGHT_Y4    As String = "年変7.75非"
    Public Const CATEGORYWORKNIGHT_Y5    As String = "年変8非"
    Public Const CATEGORYWORKNIGHT_Y6    As String = "年変8.5非"
    Public Const CATEGORYWORKNIGHT_Y7    As String = "年変9非"
    Public Const CATEGORYWORKNIGHT_Y8    As String = "年変9.5非"
    Public Const CATEGORYWORKNIGHT_Y9    As String = "年変10非"
    '*****************
    'デフォルトシフト
    '*****************
    '勤務
    Public Const CATEGORYWORK_1          As String = "昼4"
    Public Const CATEGORYWORK_2          As String = "昼6"
    Public Const CATEGORYWORK_3          As String = "昼7"
    Public Const CATEGORYWORK_4          As String = "昼7.75"
    Public Const CATEGORYWORK_5          As String = "昼8"
    Public Const CATEGORYWORK_6          As String = "昼8.5"
    Public Const CATEGORYWORK_7          As String = "昼9"
    Public Const CATEGORYWORK_8          As String = "昼9.5"
    Public Const CATEGORYWORK_9          As String = "昼10"
    Public Const CATEGORYWORK_10         As String = "夜8"
    '休日
    Public Const CATEGORYREST_1          As String = "休"
    Public Const CATEGORYREST_2          As String = "代"
    Public Const CATEGORYREST_3          As String = "明"
    Public Const CATEGORYREST_4          As String = "年休"
    Public Const CATEGORYREST_5          As String = "日振"
    Public Const CATEGORYREST_6          As String = "土祝振"
    '*****************
    '条件付き書式
    '*****************
    Public Const FC_HOLIDAY              As String = "=AND($AJ10<>""－"",DAY(MAX($D$8:$AH$8))-4<$AJ10)"
    Public Const FC_WORKDAY              As String = "=AND($AI$3<>"""",$AJ10<>""－"",$AJ10>$AI$3)"
    Public Const FC_WORKTIME             As String = "=AND($AI10<>""－"",$AI10>$AI$2)"
    Public Const FC_SHIFT_NIGHT_1        As String = "=COUNTIF(OFFSET(マスタ!$H$2,MATCH(""夜勤"",マスタ!$I$3:$I$"
    Public Const FC_SHIFT_NIGHT_2        As String = ",0),,COUNTIF(マスタ!$I$3:$I$"
    Public Const FC_SHIFT_NIGHT_3        As String = ",""夜勤"")),D$"
    Public Const FC_SHIFT_HOLIDAY_1      As String = "=COUNTIF(OFFSET(マスタ!$H$2,MATCH(""休日"",マスタ!$I$3:$I$"
    Public Const FC_SHIFT_HOLIDAY_2      As String = ",0),,COUNTIF(マスタ!$I$3:$I$"
    Public Const FC_SHIFT_HOLIDAY_3      As String = ",""休日"")),D$"
    Public Const FC_SHIFT_OVER_1         As String = "=COUNTIF(OFFSET(マスタ!$H$2,MATCH(""通し"",マスタ!$I$3:$I$"
    Public Const FC_SHIFT_OVER_2         As String = ",0),,COUNTIF(マスタ!$I$3:$I$"
    Public Const FC_SHIFT_OVER_3         As String = ",""通し"")),D$"
    '*****************
    '現場シフト
    '*****************
    '勤務
    Public Const CATEGORYSITEWORK_1      As String = "★"
    Public Const CATEGORYSITEWORK_2      As String = "●"
    Public Const CATEGORYSITEWORK_3      As String = "▲"
    Public Const CATEGORYSITEWORK_4      As String = "■"
    Public Const CATEGORYSITEWORK_5      As String = "夜勤"
    '休暇
    Public Const CATEGORYSITEREST_1      As String = "休"
    Public Const CATEGORYSITEREST_2      As String = "代休"
    Public Const CATEGORYSITEREST_3      As String = "/"
    Public Const CATEGORYSITEREST_4      As String = "年休"
    Public Const CATEGORYSITEREST_5      As String = "振休"
    Public Const CATEGORYSITEREST_6      As String = "休希望"

