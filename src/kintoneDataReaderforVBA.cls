VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "kintoneDataReaderforVBA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' -----------------------------------------------------------------------
' kintoneDataReaderforVBA v1.3.3
'  -Dictionary.cls と JsonConverter.bas のインポートが必要です。下記から取得してください。
' 　  Dictionary.cls：v1.4.1で動作確認
' 　    https://github.com/VBA-tools/VBA-Dictionary
' 　  JsonConverter.bas：v2.3.1で動作確認
' 　    https://github.com/VBA-tools/VBA-JSON
'  -Windows only
'
'
' -----------------------------------------------------------------------
Option Explicit
' -----------------------------------------------------------------------
' Public Enum
' -----------------------------------------------------------------------
'リッチエディタのタグ除去オプション
Public Enum RemoveTagsEnum
    Original            '取り除かない
    RemoveAll           '改行も取り除く
    KeepNewLine         '改行あり
End Enum
'エラー定義
Public Enum KTN_VBA_EXCEPTION
    PROP_NOT_EXISTS = 601 + vbObjectError       '指定されたプロパティは存在しません
    FCODE_NOT_EXISTS = 602 + vbObjectError      '指定されたフィールドコードは存在しません
    PROP_NOT_FOUND = 603 + vbObjectError        'プロパティ情報が存在しません
    LIMIT_MAX = 604 + vbObjectError             'リミットは1～LIMIT_MAX_VALで指定してください
    OFFSET_MIN = 605 + vbObjectError            '0以上の指定をしてください
    DATA_NOT_FOUND = 606 + vbObjectError        'レコードは取得されていません
    SETTING_NOT_FOUND = 607 + vbObjectError     'サブドメインまたはアプリIDが設定されていません
    HTTP_REQ_ERR = 608 + vbObjectError          'データ取得時エラー
    RECORD_NOT_FOUND = 609 + vbObjectError      'recordsが存在しません
    NOT_SUBTABLE = 610 + vbObjectError          'サブテーブルではありません
    UNSUPPORTED_SUBTABLE = 611 + vbObjectError  'このメソッドはサブテーブル内フィールドに対応していません
    INVALID_UTC_DIFF = 612 + vbObjectError      'UTCとの時間差が正しくありません(+09:00の形式で指定)
End Enum

' -----------------------------------------------------------------------
' Constants
' -----------------------------------------------------------------------
Private Const LIMIT_MAX_VAL As Long = 500
' -----------------------------------------------------------------------
' Private Variables
' -----------------------------------------------------------------------
Private kvu_Domain As String
Private kvu_SubDomain As String
Private kvu_AppId As String
Private kvu_isGuestSpace As Boolean
Private kvu_GuestSpaceId As String
Private kvu_AppFields As Dictionary
Private kvu_AppSubtableFields As Dictionary  'Key:サブテーブルのフィールドコード,Item:親サブテーブルのフィールドコード
Private kvu_Limit As Long       'デフォルト500で設定
Private kvu_Offset As Long      '0で設定
Private kvu_Lang As String
Private kvu_UserAuth As String
Private kvu_RemoveTag As RemoveTagsEnum
Private kvu_isName As Boolean
Private kvu_isJoinArrayData As Boolean
Private kvu_AppConf As Collection
Private kvu_Timezone As String  'UTC時間との差

'レコード取得後の変数
Private kvu_RecordCount As Long
Private kvu_TotalRecordCount As Long
Private kvu_Recordset() As Dictionary  'Dictionaryの配列

' =============================================
'  Properties
' =============================================

' ---------------------------------------------
' kintoneログインユーザー情報
'   ユーザーID:パスワード の形式
' ---------------------------------------------
Public Property Get UserAuth() As String
    UserAuth = kvu_UserAuth
End Property

' ---------------------------------------------
' アプリID
' ---------------------------------------------
Public Property Get AppId() As String
    AppId = kvu_AppId
End Property

' ---------------------------------------------
' Language
' ---------------------------------------------
Public Property Let Lang(ByVal lng As String)
    kvu_Lang = lng
End Property
Public Property Get Lang() As String
    Lang = kvu_Lang
End Property

' ---------------------------------------------
' Record取得時のlimit
' ---------------------------------------------
Public Property Let Limit(ByVal num As Long)
    If num > 0 Or num <= LIMIT_MAX_VAL Then
        kvu_Limit = num
    Else
        Err.Raise KTN_VBA_EXCEPTION.LIMIT_MAX, Description:=getErrorMessage(KTN_VBA_EXCEPTION.LIMIT_MAX)
    End If
End Property
Public Property Get Limit() As Long
    Limit = kvu_Limit
End Property

' ---------------------------------------------
' Record取得時のoffset
' ---------------------------------------------
Public Property Let Offset(ByVal index As Long)
    If index >= 0 Then
        kvu_Offset = index
    Else
        Err.Raise KTN_VBA_EXCEPTION.OFFSET_MIN, Description:=getErrorMessage(KTN_VBA_EXCEPTION.OFFSET_MIN)
    End If
End Property
Public Property Get Offset() As Long
    Offset = kvu_Offset
End Property

' ---------------------------------------------
' 取得したレコードの件数
' ---------------------------------------------
Public Property Get RecordCount() As Long
    RecordCount = kvu_RecordCount
End Property

' ---------------------------------------------
' limitを指定しなかった場合は、指定無し時抽出対象となるレコード件数
' ---------------------------------------------
Public Property Get TotalRecordCount() As Long
    TotalRecordCount = kvu_TotalRecordCount
End Property

' ---------------------------------------------
' リッチエディタのタグ除去
' ---------------------------------------------
Public Property Let RemoveTag(ByVal typ As RemoveTagsEnum)
    kvu_RemoveTag = typ
End Property
Public Property Get RemoveTag() As RemoveTagsEnum
    RemoveTag = kvu_RemoveTag
End Property

' ---------------------------------------------
' ユーザー選択、組織、グループを名前で取得するかコードで取得するか
' ---------------------------------------------
Public Property Let IsName(ByVal n As Boolean)
    kvu_isName = n
End Property
Public Property Get IsName() As Boolean
    IsName = kvu_isName
End Property
' ---------------------------------------------
' ユーザー選択、組織、グループ、チェックボックス、複数選択、ファイル、カテゴリを結合して取得するか
' ---------------------------------------------
Public Property Let IsJoinArrayData(ByVal n As Boolean)
    kvu_isJoinArrayData = n
End Property
Public Property Get IsJoinArrayData() As Boolean
    IsJoinArrayData = kvu_isJoinArrayData
End Property

' ---------------------------------------------
' UTCとの時間差(+09:00の形式) 日時フィールドをLocalに合わせたいときに指定
' ---------------------------------------------
Public Property Let Timezone(ByVal df As String)
    If Not (Left(df, 1) = "+" Or Left(df, 1) = "-") Then
        Err.Raise KTN_VBA_EXCEPTION.INVALID_UTC_DIFF, Description:=getErrorMessage(KTN_VBA_EXCEPTION.INVALID_UTC_DIFF)
    End If
    If Not (IsNumeric(Mid(df, 1, 2)) And IsNumeric(Mid(df, 5, 2))) Then
        Err.Raise KTN_VBA_EXCEPTION.INVALID_UTC_DIFF, Description:=getErrorMessage(KTN_VBA_EXCEPTION.INVALID_UTC_DIFF)
    End If
    kvu_Timezone = df
End Property
Public Property Get Timezone() As String
    Timezone = kvu_Timezone
End Property

' =============================================
'  Public Methods
' =============================================

''' <summary>
''' Setup
''' </summary>
''' <param name="subdomain">サブドメイン</param>
''' <param name="app">アプリID</param>
''' <param name="auth">ユーザーID:パスワード</param>
''' <param name="lng">指定しない場合はdefault</param>
''' <param name="isGuestspace">ゲストスペースの場合はTrue</param>
''' <param name="guestspaceid">ゲストスペースID</param>
''' <param name="domain">ドメイン</param>
Public Sub Setup(ByVal subdomain As String, ByVal app As String, ByVal auth As String, _
                 Optional ByVal lng As String = "", _
                 Optional ByVal isGuestspace As Boolean = False, _
                 Optional ByVal guestspaceid As String, _
                 Optional ByVal domain As String = "cybozu.com")
    kvu_SubDomain = subdomain
    kvu_AppId = app
    kvu_UserAuth = auth
    kvu_Domain = domain
    kvu_isGuestSpace = isGuestspace
    kvu_GuestSpaceId = guestspaceid
    Me.Lang = lng
    Me.IsName = True
    Me.IsJoinArrayData = True
    Me.RemoveTag = RemoveTagsEnum.Original
    Call kvu_GetAppConf
    Call kvu_GetFieldConf
End Sub

''' <summary>
''' Teardown
''' </summary>
Public Sub Teardown()
    kvu_SubDomain = ""
    kvu_AppId = ""
    kvu_UserAuth = ""
    kvu_Domain = ""
    kvu_isGuestSpace = False
    kvu_GuestSpaceId = ""
    Me.Lang = ""
    Me.IsName = True
    Me.IsJoinArrayData = True
    Me.RemoveTag = RemoveTagsEnum.Original
    Me.Limit = LIMIT_MAX_VAL
    Me.Offset = 0
    Me.Timezone = "+09:00"
    Call kvu_ClearRecordset
    Set kvu_AppConf = Nothing
    Set kvu_AppFields = Nothing
    Set kvu_AppSubtableFields = Nothing
End Sub

''' <summary>
''' AppFCodes
''' </summary>
''' <returns>
''' アプリのフィールドコード一覧の配列
''' サブテーブルの中は展開しない。サブテーブルまで展開したものはAppFCodesExpandを使用
''' </returns>
Public Function AppFCodes() As String()
    Dim fieldList As Variant
    Dim resultFields() As String
    
    If kvu_AppFields.Exists("properties") Then
        fieldList = kvu_AppFields("properties").Keys
        resultFields = kvu_ConvertVariant2StringArray(fieldList)
    Else
        Err.Raise KTN_VBA_EXCEPTION.PROP_NOT_FOUND, Description:=getErrorMessage(KTN_VBA_EXCEPTION.PROP_NOT_FOUND)
    End If
    AppFCodes = resultFields
End Function

''' <summary>
''' AppFCodesExpand
''' </summary>
''' <returns>
''' アプリのサブテーブルの中まで展開したフィールドコード一覧の配列
''' </returns>
Public Function AppFCodesExpand() As String()
    Dim orgFields() As String
    Dim subtblFields() As String
    Dim resultFields() As String
    Dim i As Long, j As Long, k As Long
    
    orgFields = Me.AppFCodes
    
    ReDim resultFields(LBound(orgFields) To UBound(orgFields))
    
    For i = LBound(orgFields) To UBound(orgFields)
        resultFields(i) = orgFields(i)
        If Me.FieldProperty(orgFields(i), "type") = "SUBTABLE" Then
            subtblFields = Me.AppSubtableFCodes(orgFields(i))
            ReDim Preserve resultFields(UBound(orgFields) + UBound(subtblFields) - LBound(subtblFields) + 1)
            k = 1
            For j = LBound(subtblFields) To UBound(subtblFields)
               resultFields(i + k) = subtblFields(j)
               k = k + 1
            Next
        End If
    Next
    
    Erase orgFields
    Erase subtblFields
    
    AppFCodesExpand = resultFields
End Function

''' <summary>
''' AppSubtableFCodes
''' </summary>
''' <param name="fcode">サブテーブルのフィールドコード</param>
''' <returns>
''' サブテーブル内のフィールドコード一覧の配列
''' </returns>
Public Function AppSubtableFCodes(ByVal fcode As String) As String()
    Dim resultFields() As String
    Dim subtblFields As Variant
    Dim i As Long
     
    If Me.FieldProperty(fcode, "type") <> "SUBTABLE" Then
         Err.Raise KTN_VBA_EXCEPTION.NOT_SUBTABLE, Description:=getErrorMessage(KTN_VBA_EXCEPTION.NOT_SUBTABLE)
    End If
   
    subtblFields = Me.FieldProperty(fcode, "fields")
    
    ReDim resultFields(LBound(subtblFields) To UBound(subtblFields))
    
    For i = LBound(subtblFields) To UBound(subtblFields)
        resultFields(i) = subtblFields(i)("code")
    Next
    
    AppSubtableFCodes = resultFields
End Function

''' <summary>
''' FieldType
''' </summary>
''' <param name="fcode">フィールドコード</param>
''' <returns>フィールドコードに対応するフィールドタイプ</returns>
Public Function FieldType(ByVal fcode As String) As String
    Dim stblfcode As String
    'サブテーブル内フィールドかどうかチェック
    stblfcode = kvu_FindParentSubtableFcode(fcode)

    If stblfcode = "" Then
        FieldType = Me.FieldProperty(fcode, "type")
    Else
        FieldType = Me.SubtableFieldProperty(stblfcode, fcode, "type")
    End If
End Function

''' <summary>
''' FieldLabel
''' </summary>
''' <param name="fcode">フィールドコード</param>
''' <returns>
''' フィールドコードに対応するラベル
''' ラベルがない場合はフィールドコードを返す
''' </returns>
Public Function FieldLabel(ByVal fcode As String) As String
    Dim stblfcode As String
    'サブテーブル内フィールドかどうかチェック
    stblfcode = kvu_FindParentSubtableFcode(fcode)
    
On Error GoTo subSetFCode

    If stblfcode = "" Then
        FieldLabel = Me.FieldProperty(fcode, "label")
    Else
        FieldLabel = Me.SubtableFieldProperty(stblfcode, fcode, "label")
    End If
    Exit Function
    
subSetFCode:
    If Err.Number = KTN_VBA_EXCEPTION.PROP_NOT_EXISTS Then
        FieldLabel = fcode
        Err.Clear
    Else
        Err.Raise (Err.Number)
    End If
End Function

''' <summary>
''' FieldProperty
''' </summary>
''' <param name="fcode">フィールドコード</param>
''' <param name="prop">プロパティ名</param>
''' <returns>
''' フィールドコードに対応する指定したプロパティ
''' optionsの場合は配列で返す
''' fieldsの場合はコレクション型の配列
''' サブテーブル内フィールドの場合は、SubtableFieldPropertyを使用
''' </returns>
Public Function FieldProperty(ByVal fcode As String, ByVal prop As String) As Variant
    Dim subtableArray() As Collection
    Dim i As Long, k As Long
    Dim colRec As Collection
            
    If kvu_AppFields.Exists("properties") Then
        If kvu_AppFields("properties").Exists(fcode) Then
            If kvu_AppFields("properties").Item(fcode).Exists(prop) Then
                Select Case prop
                    Case "options"  'ドロップダウンの選択肢など
                        FieldProperty = kvu_AppFields("properties").Item(fcode)(prop).Keys
                    
                    Case "fields"   'サブテーブル内のフィールド
                        Dim subtableFields As Variant
                        subtableFields = kvu_AppFields("properties").Item(fcode)(prop).Keys
                        ReDim subtableArray(LBound(subtableFields) To UBound(subtableFields))

                        For i = LBound(subtableFields) To UBound(subtableFields)
                            Dim subtableFieldsProp As Variant
                            subtableFieldsProp = kvu_AppFields("properties").Item(fcode)(prop)(subtableFields(i)).Keys
                            Set colRec = New Collection

                            '各項目値をコレクションにセット
                            For k = LBound(subtableFieldsProp) To UBound(subtableFieldsProp)
                                If subtableFieldsProp(k) = "options" Then
                                    colRec.Add Item:=kvu_AppFields("properties").Item(fcode)(prop)(subtableFields(i))(subtableFieldsProp(k)).Keys, Key:=subtableFieldsProp(k)
                                Else
                                    colRec.Add Item:=kvu_AppFields("properties").Item(fcode)(prop)(subtableFields(i))(subtableFieldsProp(k)), Key:=subtableFieldsProp(k)
                                End If
                            Next
                            'コレクションの配列にセット
                            Set subtableArray(i) = colRec
                        Next
                        
                        FieldProperty = subtableArray
                    Case Else
                        FieldProperty = kvu_AppFields("properties").Item(fcode)(prop)
                End Select
            Else
                Err.Raise KTN_VBA_EXCEPTION.PROP_NOT_EXISTS, Description:=getErrorMessage(KTN_VBA_EXCEPTION.PROP_NOT_EXISTS) & "(" & prop & ")"
            End If
        Else
            If kvu_FindParentSubtableFcode(fcode) <> "" Then
                Err.Raise KTN_VBA_EXCEPTION.DATA_NOT_FOUND, Description:=getErrorMessage(KTN_VBA_EXCEPTION.DATA_NOT_FOUND)
            Else
                If fcode = "$id" Or fcode = "$revision" Then
                    Select Case prop
                        Case "type"
                            If fcode = "$id" Then
                                FieldProperty = "__ID__"
                            ElseIf fcode = "$revision" Then
                                FieldProperty = "__REVISION__"
                            End If
                        Case Else
                            Err.Raise KTN_VBA_EXCEPTION.PROP_NOT_EXISTS, Description:=getErrorMessage(KTN_VBA_EXCEPTION.PROP_NOT_EXISTS) & "(" & prop & ")"
                    End Select
                Else
                    Err.Raise KTN_VBA_EXCEPTION.FCODE_NOT_EXISTS, Description:=getErrorMessage(KTN_VBA_EXCEPTION.FCODE_NOT_EXISTS) & "(" & fcode & ")"
                End If
            End If
        End If
    Else
        Err.Raise KTN_VBA_EXCEPTION.PROP_NOT_FOUND, Description:=getErrorMessage(KTN_VBA_EXCEPTION.PROP_NOT_FOUND)
    End If
End Function

''' <summary>
''' SubtableFieldProperty
''' </summary>
''' <param name="fcode">フィールドコード</param>
''' <param name="stblfcode">サブテーブル内フィールドコード</param>
''' <param name="prop">プロパティ名</param>
''' <returns>
''' フィールドコードに対応する指定したプロパティ
''' optionsの場合は配列で返す
''' </returns>
Public Function SubtableFieldProperty(ByVal fcode As String, ByVal stblfcode As String, ByVal prop As String) As Variant
    Dim subfields As Variant
    Dim i As Long
    subfields = Me.FieldProperty(fcode, "fields")
    For i = LBound(subfields) To UBound(subfields)
        If subfields(i)("code") = stblfcode Then
            If prop = "options" Then 'ドロップダウンの選択肢など
                SubtableFieldProperty = subfields(i)(prop).Keys
            Else
                SubtableFieldProperty = subfields(i)(prop)
            End If
            Exit For
        End If
    Next
End Function

''' <summary>
''' RecordFields
''' </summary>
''' <param name="existsPropOnly">$id、$revisionなどAppFCodesで取得できないものは除く場合にTrue</param>
''' <returns>取得したレコードのフィールドコードの配列</returns>
Public Function RecordFields(Optional ByVal existsPropOnly As Boolean = True) As String()
    Dim resultFields() As String
    Dim f As Variant
    Dim isTarget As Boolean
    
    If Me.RecordCount > 0 Then
        '不要なものを除く
        For Each f In kvu_Recordset(0).Keys
            If existsPropOnly Then
                isTarget = kvu_AppFields("properties").Exists(f)
            Else
                isTarget = True
            End If
            If isTarget Then
                Call incrimentArray(resultFields)
                resultFields(UBound(resultFields)) = CStr(f)
            End If
        Next
        
        RecordFields = resultFields
    Else
        Err.Raise KTN_VBA_EXCEPTION.DATA_NOT_FOUND, Description:=getErrorMessage(KTN_VBA_EXCEPTION.DATA_NOT_FOUND)
    End If
End Function

''' <summary>
''' FieldValue
''' </summary>
''' <param name="index">レコード位置</param>
''' <param name="fcode">フィールドコード</param>
''' <returns>
''' 指定されたindexの位置のデータ
''' 値が複数あるものについては、IsJoinArrayDataプロパティがFalseの場合、配列。データ無しの場合は、空文字
''' SUBTABLEはコレクション型の配列。データ無しの場合は、空文字
''' </returns>
Public Function FieldValue(ByVal index As Long, ByVal fcode As String) As Variant
    Dim i As Long
    Dim subtableArray() As Collection
    Dim subtableRecCnt As Long  'サブテーブル内のレコード件数
    Dim myType As String
    
    If Me.RecordCount > 0 Then
        
        If kvu_FindParentSubtableFcode(fcode) <> "" Then
            Err.Raise KTN_VBA_EXCEPTION.DATA_NOT_FOUND, Description:=getErrorMessage(KTN_VBA_EXCEPTION.DATA_NOT_FOUND)
        End If
        
        myType = Me.FieldType(fcode)
        
        Select Case myType
            Case "SUBTABLE"
            
                subtableRecCnt = kvu_Recordset(index)(fcode)("value").Count
        
                If subtableRecCnt > 0 Then
        
                    '戻り値の配列をレコード件数で確保
                    ReDim subtableArray(0 To subtableRecCnt - 1)
                    Dim dt As Dictionary
                    Dim colRec As Collection
                    
                    'サブテーブルのフィールド一覧を取得
                    Dim s As Variant
                    'サブテーブルのレコード件数分ループ
                    For i = 1 To subtableRecCnt
                        Set colRec = New Collection
                        '各項目値をコレクションにセット
                        For Each s In kvu_Recordset(index)(fcode)("value")(1)("value").Keys '1からスタート
                            Set dt = kvu_Recordset(index)(fcode)("value")(i)("value")(s)
                            colRec.Add Item:=kvu_GetFieldValue(dt("type"), dt("value")), Key:=s
                        Next
                        'コレクションの配列にセット
                        Set subtableArray(i - 1) = colRec
                    Next
                    FieldValue = subtableArray
               Else
                    FieldValue = ""
               End If
        
        Case "GROUP"
            FieldValue = ""

        Case Else
            If IsEmpty(kvu_Recordset(index)(fcode)) Then
                FieldValue = Empty
            Else
                FieldValue = kvu_GetFieldValue(myType, kvu_Recordset(index)(fcode)("value"))
            End If
        End Select
    Else
        Err.Raise KTN_VBA_EXCEPTION.DATA_NOT_FOUND, Description:=getErrorMessage(KTN_VBA_EXCEPTION.DATA_NOT_FOUND)
    End If
    
End Function

''' <summary>
''' RecordArray
''' </summary>
''' <param name="addFieldLabel">配列の最初にフィールドラベルを追加するか</param>
''' <param name="existsPropOnly">$id、$revisionなどAppFCodesで取得できないものは除く場合にTrue</param>
''' <returns>
''' 取得データ全ての二次元配列（サブテーブルの値は取得しない、複数選択系はカンマ区切り）
''' </returns>
Public Function RecordArray(Optional ByVal addFieldLabel As Boolean = True, Optional ByVal existsPropOnly As Boolean = True) As Variant
    Dim varData As Variant
    Dim i As Long, k As Long
    Dim adj As Integer: adj = IIf(addFieldLabel, 1, 0)
    Dim fieldsOrg() As String
    Dim fields() As String  'サブテーブル除外
    Dim fVal As Variant
    
    If Me.RecordCount > 0 Then
        'レコードのフィールド一覧取得し、サブテーブルを除去
        fieldsOrg = Me.RecordFields(existsPropOnly)
        For k = LBound(fieldsOrg) To UBound(fieldsOrg)
            If Me.FieldType(CStr(fieldsOrg(k))) <> "SUBTABLE" Then
                Call incrimentArray(fields)
                fields(UBound(fields)) = fieldsOrg(k)
            End If
        Next
        
        ReDim varData(0 To Me.RecordCount - 1 + adj, LBound(fields) To UBound(fields))
        For i = 0 To Me.RecordCount - 1
            For k = LBound(fields) To UBound(fields)
                If addFieldLabel And i = 0 Then
                    varData(0, k) = Me.FieldLabel(CStr(fields(k)))
                End If
                fVal = Me.FieldValue(i, fields(k))
                If IsArray(fVal) Then
                    '複数選択、チェックボックス、ユーザー選択など複数ある場合は、カンマ区切りで
                    varData(i + adj, k) = Join(fVal, ",")
                Else
                    varData(i + adj, k) = IIf(Left(fVal, 1) = "=", "'" & fVal, fVal)
                End If
            Next
        Next
        RecordArray = varData
    Else
        Err.Raise KTN_VBA_EXCEPTION.DATA_NOT_FOUND, Description:=getErrorMessage(KTN_VBA_EXCEPTION.DATA_NOT_FOUND)
    End If

End Function

''' <summary>
''' GetRecords
''' </summary>
''' <param name="isAll">全レコード取得する場合はTrue、Limitを指定する場合はFalse</param>
''' <param name="sFields">取得フィールド。複数の場合はカンマ区切りで</param>
''' <param name="query">limitとoffsetは含めずクラスのプロパティで設定</param>
''' <returns>
''' データ量が多い場合に全件取得すると、PCによってはメモリ不足になります。
''' </returns>
Public Sub GetRecords(ByVal isAll As Boolean, Optional ByVal sFields As String = "", Optional ByVal query As String = "")
    Dim strQuery As String
    Dim jsonData As Dictionary
    Dim returnData() As Dictionary
    Dim dataCount As Long
    Dim offsetIdx As Long
    Dim i As Long
    Dim isFirst As Boolean: isFirst = True
    
    'クリア
    Call kvu_ClearRecordset
    Erase returnData
    
    'limitで指定した件数（MAX500件）しか取れない
    offsetIdx = IIf(isAll, 0, Me.Offset)
    dataCount = Me.Limit

    Do Until dataCount < Me.Limit
        'クエリの作成
        strQuery = IIf(sFields <> "", "&fields=" & knt_EncodeUrl(sFields), "")
        'WebHelperのUrlEncodeは日本語NG
        strQuery = strQuery & "&query=" & knt_EncodeUrl(query & " limit " & Me.Limit & " offset " & offsetIdx)   'limit 500 offset 0
        
        Set jsonData = kvu_Execute(kvu_RecordsUrl & strQuery & "&totalCount=true")
        
        If Not jsonData.Exists("records") Then
            Dim errMsg As String
            If jsonData.Exists("message") Then
               errMsg = " - " & jsonData("message")
            End If
            Err.Raise KTN_VBA_EXCEPTION.RECORD_NOT_FOUND, Description:=getErrorMessage(KTN_VBA_EXCEPTION.RECORD_NOT_FOUND) & "　" & errMsg & "(GetRecords)"
        End If
        
        dataCount = jsonData("records").Count
        
        kvu_TotalRecordCount = jsonData("totalCount")
  
        If isFirst Then
            If dataCount > 0 Then
                ReDim returnData(0 To IIf(isAll, kvu_TotalRecordCount - 1, dataCount - 1))
                isFirst = False
            Else
                Exit Sub
            End If
        End If
        
        For i = 1 To dataCount
           Set returnData(IIf(isAll, offsetIdx, 0) + i - 1) = jsonData("records")(i)
        Next i
        
        offsetIdx = offsetIdx + dataCount
        Set jsonData = Nothing
        
        If Not isAll Then Exit Do
    Loop
    
    kvu_RecordCount = UBound(returnData) + 1
    kvu_Recordset = returnData
    
End Sub

''' <summary>
''' リッチエディタのタグ除去
''' </summary>
''' <param name="str"></param>
''' <param name="typ"></param>
''' <returns></returns>
Public Function RemoveTags(ByVal str As String, ByVal typ As RemoveTagsEnum) As String
    If typ = RemoveTagsEnum.Original Then
        RemoveTags = str
    Else
        Dim doc As Object
        Set doc = CreateObject("htmlfile")
        Call doc.write(str)
        Dim retStr As String
        If doc.FirstChild Is Nothing Then
        Else
            retStr = doc.FirstChild.innerText
            If typ = RemoveTagsEnum.RemoveAll Then
                retStr = Replace(Replace(Replace(retStr, vbCr, ""), vbCrLf, ""), vbLf, "")
            End If
        End If
        RemoveTags = retStr
    End If
 End Function
 
''' <summary>
''' UTCをLoacl日時に変更
''' </summary>
''' <param name="dtUtc">2012-01-11T11:30:00Z のフォーマット</param>
''' <returns></returns>
Public Function UTC2Local(ByVal dtUtc As String) As Date

    Dim localDate As Date, utDate As Date
    If Trim(dtUtc) = "" Then
        UTC2Local = Empty
        Exit Function
    End If
    utDate = CDate(Replace(Replace(dtUtc, "T", " "), "Z", ""))
    localDate = utDate
    
    If kvu_Timezone <> "" Then
        If Left(kvu_Timezone, 1) = "+" Then
            localDate = utDate + TimeValue(Mid(kvu_Timezone, 2))
        Else
            localDate = utDate - TimeValue(Mid(kvu_Timezone, 2))
        End If
    End If
    
    UTC2Local = localDate

End Function

' =============================================
' Private Methods
' =============================================
''' <summary>
''' kintoneBaseURL
''' </summary>
Public Function kvu_BaseUrl() As String
    kvu_BaseUrl = "https://???." & kvu_Domain & "/"
End Function

''' <summary>
''' アプリ情報の取得
''' </summary>
Private Sub kvu_GetAppConf()
    Dim result As Object
    Dim r As Object
    
    Set result = kvu_Execute(kvu_AppUrl)
    Set kvu_AppConf = New Collection
    
    kvu_AppConf.Add result("code"), "code"
    kvu_AppConf.Add result("name"), "name"
    kvu_AppConf.Add result("description"), "description"
    kvu_AppConf.Add result("spaceId"), "spaceId"
    kvu_AppConf.Add result("threadId"), "threadId"
    kvu_AppConf.Add result("createdAt"), "createdAt"
    kvu_AppConf.Add result("creator")("code"), "creator.code"
    kvu_AppConf.Add result("creator")("name"), "creator.name"
    kvu_AppConf.Add result("modifiedAt"), "modifiedAt"
    kvu_AppConf.Add result("modifier")("code"), "modifier.code"
    kvu_AppConf.Add result("modifier")("name"), "modifier.name"
    
    Set result = Nothing
End Sub

''' <summary>
''' フィールド情報の取得
''' </summary>
Private Sub kvu_GetFieldConf()
    Dim fieldList() As String
    Dim subtfields() As String
    Dim i As Long, k As Long
    
    Set kvu_AppFields = kvu_Execute(kvu_FieldsUrl)
    
    'サブテーブルの情報を保存
    Set kvu_AppSubtableFields = New Dictionary
    fieldList = Me.AppFCodes
    For i = LBound(fieldList) To UBound(fieldList)
        If Me.FieldProperty(fieldList(i), "type") = "SUBTABLE" Then
            subtfields = Me.AppSubtableFCodes(fieldList(i))
            For k = LBound(subtfields) To UBound(subtfields)
                kvu_AppSubtableFields.Add subtfields(k), fieldList(i)
            Next
        End If
    Next
    Erase fieldList
    Erase subtfields
End Sub

''' <summary>
''' データ取得共通処理
''' </summary>
''' <param name="url"></param>
''' <returns></returns>
Private Function kvu_Execute(ByVal url As String) As Object
    Dim objHttpReq As Object    'XMLHTTP オブジェクト
    Dim strJson As String       'レスポンスで受け取るJSONデータ
    
    If kvu_SubDomain = "" Or kvu_AppId = "" Then
        Err.Raise KTN_VBA_EXCEPTION.SETTING_NOT_FOUND, Description:=getErrorMessage(KTN_VBA_EXCEPTION.SETTING_NOT_FOUND)
    End If
    
    Set objHttpReq = CreateObject("MSXML2.XMLHTTP")
    objHttpReq.Open "GET", url, False

    'リクエストヘッダー指定
    With objHttpReq
        .SetRequestHeader "X-Cybozu-Authorization", knt_EncodeBase64(Me.UserAuth)   'ログイン認証
        .SetRequestHeader "Host", kvu_SubDomain & ".cybozu.com" + ":443"            'ドメイン名：ポート番号
        .SetRequestHeader "If-Modified-Since", "Thu, 01 Jun 1970 00:00:00 GMT"      'キャッシュ対策（常にレスポンスが取得できる状態にする）
    End With
    
    'リクエスト送信
    objHttpReq.Send (Null)

    'レスポンス取得
    strJson = objHttpReq.ResponseText

    If objHttpReq.status <> 200 Then
        Dim errObj As Object
        Set errObj = ParseJson(strJson)
        If errObj.Exists("message") Then
            Err.Raise KTN_VBA_EXCEPTION.HTTP_REQ_ERR, Description:=CStr(objHttpReq.status) & " - " & errObj("message") & "(kvu_Execute)"
        Else
            Err.Raise KTN_VBA_EXCEPTION.HTTP_REQ_ERR, Description:=CStr(objHttpReq.status) & "(kvu_Execute)"
        End If
    End If

    Set objHttpReq = Nothing
    
    'Parseしてデータセット
    Set kvu_Execute = ParseJson(strJson)

End Function

''' <summary>
'''
''' </summary>
''' <param name="fieldType"></param>
''' <param name="rec"></param>
''' <returns></returns>
Private Function kvu_GetFieldValue(ByVal FieldType As String, ByVal rec As Variant) As Variant
    Dim i As Long
    Dim valueArry() As String
    
    Select Case FieldType
       
        Case "CREATOR", "MODIFIER"
            Application.Wait [now()+"00:00:01"] 'Excelが落ちてしまう場合があるので回避策
            If kvu_isName Then
                kvu_GetFieldValue = rec("name")
            Else
                kvu_GetFieldValue = rec("code")
            End If

        Case "USER_SELECT", "STATUS_ASSIGNEE", "GROUP_SELECT", "ORGANIZATION_SELECT", "FILE"
            If IsNull(rec) Then
                kvu_GetFieldValue = ""
            Else
                If rec.Count > 0 Then
                    ReDim valueArry(0 To rec.Count - 1)
                    For i = 1 To rec.Count
                        If FieldType <> "FILE" And Not kvu_isName Then
                            valueArry(i - 1) = rec(i)("code")
                        Else
                            valueArry(i - 1) = rec(i)("name") '1からスタート,ファイルの場合はファイル名を返す
                        End If
                    Next
                    If kvu_isJoinArrayData Then
                        kvu_GetFieldValue = Join(valueArry, ",")
                    Else
                        kvu_GetFieldValue = valueArry
                    End If
                Else
                    kvu_GetFieldValue = ""
                End If
            End If
     
        Case "CHECK_BOX", "MULTI_SELECT", "CATEGORY"
            If IsNull(rec) Then
                kvu_GetFieldValue = ""
            Else
                If rec.Count > 0 Then
                    ReDim valueArry(0 To rec.Count - 1)
                    For i = 1 To rec.Count
                        valueArry(i - 1) = rec(i) '1からスタート
                    Next
                    If kvu_isJoinArrayData Then
                        kvu_GetFieldValue = Join(valueArry, ",")
                    Else
                        kvu_GetFieldValue = valueArry
                    End If
                Else
                    kvu_GetFieldValue = ""
                End If
            End If

        Case Else
            Dim val As Variant
            If IsNull(rec) Then
                val = Empty
            Else
                val = rec
            End If
            If FieldType = "RICH_TEXT" And Me.RemoveTag <> RemoveTagsEnum.Original Then
                kvu_GetFieldValue = RemoveTags(val, Me.RemoveTag)
            ElseIf FieldType = "DATETIME" Or FieldType = "CREATED_TIME" Or FieldType = "UPDATED_TIME" Then
                kvu_GetFieldValue = CStr(Format(UTC2Local(CStr(val)), "YYYY/MM/DD HH:MM:SS"))
            Else
                kvu_GetFieldValue = val
            End If
    End Select
    Erase valueArry
End Function

''' <summary>
''' フィールドコードの親サブテーブルを検索
''' </summary>
''' <param name="stblfcode">サブテーブル内フィールドのフィールドコード</param>
''' <returns>親サブテーブルのフィールドコード、該当するサブテーブルがなければ空文字を返す</returns>
Private Function kvu_FindParentSubtableFcode(ByVal stblfcode As String) As String
    If kvu_AppSubtableFields.Exists(stblfcode) Then
        kvu_FindParentSubtableFcode = kvu_AppSubtableFields(stblfcode)
    Else
        kvu_FindParentSubtableFcode = ""
    End If
End Function

''' <summary>
''' アプリ取得のURL
''' </summary>
''' <returns></returns>
Private Function kvu_AppUrl() As String
    If kvu_isGuestSpace Then
        kvu_AppUrl = kvu_BaseUrl & "k/guest/" & kvu_GuestSpaceId
    Else
        kvu_AppUrl = kvu_BaseUrl & "k"
    End If
    kvu_AppUrl = Replace(kvu_AppUrl & "/v1/app.json", "???", kvu_SubDomain)
    kvu_AppUrl = kvu_AppUrl & "?&id=" & kvu_AppId
End Function

''' <summary>
''' レコード取得のURL
''' </summary>
''' <returns></returns>
Private Function kvu_RecordsUrl() As String
    If kvu_isGuestSpace Then
        kvu_RecordsUrl = kvu_BaseUrl & "k/guest/" & kvu_GuestSpaceId
    Else
        kvu_RecordsUrl = kvu_BaseUrl & "k"
    End If
    kvu_RecordsUrl = Replace(kvu_RecordsUrl & "/v1/records.json", "???", kvu_SubDomain)
    kvu_RecordsUrl = kvu_RecordsUrl & "?&app=" & kvu_AppId
End Function

''' <summary>
''' フィールド取得のURL
''' </summary>
''' <returns></returns>
Private Function kvu_FieldsUrl() As String
    If kvu_isGuestSpace Then
        kvu_FieldsUrl = kvu_BaseUrl & "k/guest/" & kvu_GuestSpaceId
    Else
        kvu_FieldsUrl = kvu_BaseUrl & "k"
    End If
    kvu_FieldsUrl = Replace(kvu_FieldsUrl & "/v1/app/form/fields.json", "???", kvu_SubDomain)
    kvu_FieldsUrl = kvu_FieldsUrl & "?app=" & kvu_AppId & IIf(Me.Lang <> "", "&lang=" & Me.Lang, "")
End Function

''' <summary>
''' レコード取得前クリア処理
''' </summary>
Private Sub kvu_ClearRecordset()
    kvu_RecordCount = 0
    kvu_TotalRecordCount = 0
    Application.Wait [now()+"00:00:01"] 'Excelが落ちてしまう場合があるので回避策
    DoEvents
    Erase kvu_Recordset
End Sub
 
''' <summary>
''' Variant配列をString配列に変換
''' </summary>
''' <param name="varArray"></param>
''' <returns>
''' アプリのフィールドコード一覧（サブテーブルの中まで展開）の配列
''' </returns>
Private Function kvu_ConvertVariant2StringArray(ByVal varArray As Variant) As String()
    Dim resultArray() As String
    Dim i As Long
    
    ReDim resultArray(LBound(varArray) To UBound(varArray))
    For i = LBound(varArray) To UBound(varArray)
        resultArray(i) = CStr(varArray(i))
    Next
    kvu_ConvertVariant2StringArray = resultArray
End Function

''' <summary>
''' EncodeUrl
''' </summary>
''' <param name="sWord"></param>
''' <returns></returns>
Private Function knt_EncodeUrl(ByVal sWord As String) As String
    Dim d As Object
    Dim elm As Object
  
    sWord = Replace(sWord, "\", "\\")
    sWord = Replace(sWord, "'", "\'")
    Set d = CreateObject("htmlfile")
    Set elm = d.CreateElement("span")
    elm.setAttribute "id", "result"
    d.body.appendChild elm
    d.parentWindow.execScript "document.getElementById('result').innerText = encodeURIComponent('" & sWord & "');", "JScript"
    knt_EncodeUrl = elm.innerText
End Function

''' <summary>
''' EncodeBase64
''' </summary>
''' <param name="txt"></param>
''' <returns></returns>
Private Function knt_EncodeBase64(ByVal txt As String) As String
    Dim XMLDocument As Object
    Dim elm As Object
    Dim byteArray() As Byte
    Dim resultTxt As String
    
    byteArray = StrConv(txt, vbFromUnicode)
    'byteArray = CreateObject("System.Text.UTF8Encoding").GetBytes_4(txt)
    
    Set XMLDocument = CreateObject("MSXML2.DOMDocument")
    Set elm = XMLDocument.CreateElement("b64")

    elm.DataType = "bin.base64"
    elm.NodeTypedValue = byteArray
    resultTxt = elm.Text

    Set elm = Nothing
    Set XMLDocument = Nothing
    
    knt_EncodeBase64 = Replace$(resultTxt, vbLf, "")
End Function

''' <summary>
''' getErrorMessage
''' </summary>
''' <param name="num"></param>
''' <returns></returns>
Private Function getErrorMessage(ByVal num As KTN_VBA_EXCEPTION) As String
    Dim result As String
    
    Select Case num
        Case DATA_NOT_FOUND
            result = "レコードは取得されていません"
        Case FCODE_NOT_EXISTS
            result = "指定されたフィールドコードは存在しません"
        Case HTTP_REQ_ERR
            result = "データ取得時エラー"
        Case INVALID_UTC_DIFF
            result = "UTCとの時間差が正しくありません(+09:00の形式で指定)"
        Case LIMIT_MAX
            result = "リミットは1～" & CStr(LIMIT_MAX_VAL) & "で指定してください"
        Case OFFSET_MIN
            result = "0以上の指定をしてください"
        Case PROP_NOT_EXISTS
            result = "指定されたプロパティは存在しません"
        Case PROP_NOT_FOUND
            result = "プロパティ情報が存在しません"
        Case RECORD_NOT_FOUND
            result = "recordsが存在しません"
        Case SETTING_NOT_FOUND
            result = "サブドメインまたはアプリIDが設定されていません"
        Case NOT_SUBTABLE
            result = "サブテーブルではありません"
        Case UNSUPPORTED_SUBTABLE
            result = "このメソッドはサブテーブル内フィールドに対応していません"
    End Select
    getErrorMessage = result
End Function

''' <summary>
''' incrimentArray
''' </summary>
''' <param name="rArray"></param>
Private Sub incrimentArray(ByRef rArray() As String)

On Error GoTo ErrHandler
    ReDim Preserve rArray(UBound(rArray) + 1)
    
    Exit Sub
    
ErrHandler:
    If Err.Number = 9 Then
        ReDim rArray(0)
    Else
        Err.Raise (Err.Number)
    End If
End Sub
Private Sub Class_Initialize()
    Me.Lang = ""
    Me.IsName = True
    Me.IsJoinArrayData = True
    Me.RemoveTag = RemoveTagsEnum.Original
    Me.Limit = LIMIT_MAX_VAL
    Me.Offset = 0
    Me.Timezone = "+09:00"  'デフォルトを日本
    Call kvu_ClearRecordset
End Sub

Private Sub Class_Terminate()
    Call kvu_ClearRecordset
End Sub


