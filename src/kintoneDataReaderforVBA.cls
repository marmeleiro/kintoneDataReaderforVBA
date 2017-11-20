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
' kintoneDataReaderforVBA v1.1.0
'  -Dictionary.cls �� JsonConverter.bas �̃C���|�[�g���K�v�ł��B���L����擾���Ă��������B
' �@  Dictionary.cls�Fv1.4.1�œ���m�F
' �@    https://github.com/VBA-tools/VBA-Dictionary
' �@  JsonConverter.bas�Fv2.2.2�œ���m�F
' �@    https://github.com/VBA-tools/VBA-JSON
'  -Windows only
'
'
' -----------------------------------------------------------------------
Option Explicit
' -----------------------------------------------------------------------
' Public Enum
' -----------------------------------------------------------------------
'���b�`�G�f�B�^�̃^�O�����I�v�V����
Public Enum RemoveTagsEnum
    Original            '��菜���Ȃ�
    RemoveAll           '��菜��
    RemoveChangeNewLine '<br /> ��</div>�����s�ɒu�������Ď�菜��
End Enum
'�G���[��`
Public Enum KTN_VBA_EXCEPTION
    PROP_NOT_EXISTS = 601 + vbObjectError       '�w�肳�ꂽ�v���p�e�B�͑��݂��܂���
    FCODE_NOT_EXISTS = 602 + vbObjectError      '�w�肳�ꂽ�t�B�[���h�R�[�h�͑��݂��܂���
    PROP_NOT_FOUND = 603 + vbObjectError        '�v���p�e�B��񂪑��݂��܂���
    LIMIT_MAX = 604 + vbObjectError             '���~�b�g��1�`LIMIT_MAX_VAL�Ŏw�肵�Ă�������
    OFFSET_MIN = 605 + vbObjectError            '0�ȏ�̎w������Ă�������
    DATA_NOT_FOUND = 606 + vbObjectError        '���R�[�h�͎擾����Ă��܂���
    SETTING_NOT_FOUND = 607 + vbObjectError     '�T�u�h���C���܂��̓A�v��ID���ݒ肳��Ă��܂���
    HTTP_REQ_ERR = 608 + vbObjectError          '�f�[�^�擾���G���[
    RECORD_NOT_FOUND = 609 + vbObjectError      'records�����݂��܂���
    NOT_SUBTABLE = 610 + vbObjectError          '�T�u�e�[�u���ł͂���܂���
    UNSUPPORTED_SUBTABLE = 611 + vbObjectError  '���̃��\�b�h�̓T�u�e�[�u�����t�B�[���h�ɑΉ����Ă��܂���
    INVALID_UTC_DIFF = 612 + vbObjectError      'UTC�Ƃ̎��ԍ�������������܂���(+09:00�̌`���Ŏw��)
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
Private kvu_AppSubtableFields As Dictionary  'Key:�T�u�e�[�u���̃t�B�[���h�R�[�h,Item:�e�T�u�e�[�u���̃t�B�[���h�R�[�h
Private kvu_Limit As Long       '�f�t�H���g500�Őݒ�
Private kvu_Offset As Long      '0�Őݒ�
Private kvu_Lang As String
Private kvu_UserAuth As String
Private kvu_RemoveTag As RemoveTagsEnum
Private kvu_isName As Boolean
Private kvu_isJoinArrayData As Boolean
Private kvu_AppConf As Collection
Private kvu_Timezone As String  'UTC���ԂƂ̍�

'���R�[�h�擾��̕ϐ�
Private kvu_RecordCount As Long
Private kvu_TotalRecordCount As Long
Private kvu_Recordset() As Dictionary  'Dictionary�̔z��

' =============================================
'  Properties
' =============================================

' ---------------------------------------------
' kintone���O�C�����[�U�[���
'   ���[�U�[ID:�p�X���[�h �̌`��
' ---------------------------------------------
Public Property Get UserAuth() As String
    UserAuth = kvu_UserAuth
End Property

' ---------------------------------------------
' �A�v��ID
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
' Record�擾����limit
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
' Record�擾����offset
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
' �擾�������R�[�h�̌���
' ---------------------------------------------
Public Property Get RecordCount() As Long
    RecordCount = kvu_RecordCount
End Property

' ---------------------------------------------
' limit���w�肵�Ȃ������ꍇ�́A�w�薳�������o�ΏۂƂȂ郌�R�[�h����
' ---------------------------------------------
Public Property Get TotalRecordCount() As Long
    TotalRecordCount = kvu_TotalRecordCount
End Property

' ---------------------------------------------
' ���b�`�G�f�B�^�̃^�O����
' ---------------------------------------------
Public Property Let RemoveTag(ByVal typ As RemoveTagsEnum)
    kvu_RemoveTag = typ
End Property
Public Property Get RemoveTag() As RemoveTagsEnum
    RemoveTag = kvu_RemoveTag
End Property

' ---------------------------------------------
' ���[�U�[�I���A�g�D�A�O���[�v�𖼑O�Ŏ擾���邩�R�[�h�Ŏ擾���邩
' ---------------------------------------------
Public Property Let IsName(ByVal n As Boolean)
    kvu_isName = n
End Property
Public Property Get IsName() As Boolean
    IsName = kvu_isName
End Property
' ---------------------------------------------
' ���[�U�[�I���A�g�D�A�O���[�v�A�`�F�b�N�{�b�N�X�A�����I���A�t�@�C���A�J�e�S�����������Ď擾���邩
' ---------------------------------------------
Public Property Let IsJoinArrayData(ByVal n As Boolean)
    kvu_isJoinArrayData = n
End Property
Public Property Get IsJoinArrayData() As Boolean
    IsJoinArrayData = kvu_isJoinArrayData
End Property

' ---------------------------------------------
' UTC�Ƃ̎��ԍ�(+09:00�̌`��) �����t�B�[���h��Local�ɍ��킹�����Ƃ��Ɏw��
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
''' <param name="subdomain">�T�u�h���C��</param>
''' <param name="app">�A�v��ID</param>
''' <param name="auth">���[�U�[ID:�p�X���[�h</param>
''' <param name="lng">�w�肵�Ȃ��ꍇ��default</param>
''' <param name="isGuestspace">�Q�X�g�X�y�[�X�̏ꍇ��True</param>
''' <param name="guestspaceid">�Q�X�g�X�y�[�XID</param>
''' <param name="domain">�h���C��</param>
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
''' �A�v���̃t�B�[���h�R�[�h�ꗗ�̔z��
''' �T�u�e�[�u���̒��͓W�J���Ȃ��B�T�u�e�[�u���܂œW�J�������̂�AppFCodesExpand���g�p
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
''' �A�v���̃T�u�e�[�u���̒��܂œW�J�����t�B�[���h�R�[�h�ꗗ�̔z��
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
''' <param name="fcode">�T�u�e�[�u���̃t�B�[���h�R�[�h</param>
''' <returns>
''' �T�u�e�[�u�����̃t�B�[���h�R�[�h�ꗗ�̔z��
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
''' <param name="fcode">�t�B�[���h�R�[�h</param>
''' <returns>�t�B�[���h�R�[�h�ɑΉ�����t�B�[���h�^�C�v</returns>
Public Function FieldType(ByVal fcode As String) As String
    Dim stblfcode As String
    '�T�u�e�[�u�����t�B�[���h���ǂ����`�F�b�N
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
''' <param name="fcode">�t�B�[���h�R�[�h</param>
''' <returns>
''' �t�B�[���h�R�[�h�ɑΉ����郉�x��
''' ���x�����Ȃ��ꍇ�̓t�B�[���h�R�[�h��Ԃ�
''' </returns>
Public Function FieldLabel(ByVal fcode As String) As String
    Dim stblfcode As String
    '�T�u�e�[�u�����t�B�[���h���ǂ����`�F�b�N
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
''' <param name="fcode">�t�B�[���h�R�[�h</param>
''' <param name="prop">�v���p�e�B��</param>
''' <returns>
''' �t�B�[���h�R�[�h�ɑΉ�����w�肵���v���p�e�B
''' options�̏ꍇ�͔z��ŕԂ�
''' fields�̏ꍇ�̓R���N�V�����^�̔z��
''' �T�u�e�[�u�����t�B�[���h�̏ꍇ�́ASubtableFieldProperty���g�p
''' </returns>
Public Function FieldProperty(ByVal fcode As String, ByVal prop As String) As Variant
    Dim subtableArray() As Collection
    Dim i As Long, k As Long
    Dim colRec As Collection
            
    If kvu_AppFields.Exists("properties") Then
        If kvu_AppFields("properties").Exists(fcode) Then
            If kvu_AppFields("properties").Item(fcode).Exists(prop) Then
                Select Case prop
                    Case "options"  '�h���b�v�_�E���̑I�����Ȃ�
                        FieldProperty = kvu_AppFields("properties").Item(fcode)(prop).Keys
                    
                    Case "fields"   '�T�u�e�[�u�����̃t�B�[���h
                        Dim subtableFields As Variant
                        subtableFields = kvu_AppFields("properties").Item(fcode)(prop).Keys
                        ReDim subtableArray(LBound(subtableFields) To UBound(subtableFields))

                        For i = LBound(subtableFields) To UBound(subtableFields)
                            Dim subtableFieldsProp As Variant
                            subtableFieldsProp = kvu_AppFields("properties").Item(fcode)(prop)(subtableFields(i)).Keys
                            Set colRec = New Collection

                            '�e���ڒl���R���N�V�����ɃZ�b�g
                            For k = LBound(subtableFieldsProp) To UBound(subtableFieldsProp)
                                If subtableFieldsProp(k) = "options" Then
                                    colRec.Add Item:=kvu_AppFields("properties").Item(fcode)(prop)(subtableFields(i))(subtableFieldsProp(k)).Keys, Key:=subtableFieldsProp(k)
                                Else
                                    colRec.Add Item:=kvu_AppFields("properties").Item(fcode)(prop)(subtableFields(i))(subtableFieldsProp(k)), Key:=subtableFieldsProp(k)
                                End If
                            Next
                            '�R���N�V�����̔z��ɃZ�b�g
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
''' <param name="fcode">�t�B�[���h�R�[�h</param>
''' <param name="stblfcode">�T�u�e�[�u�����t�B�[���h�R�[�h</param>
''' <param name="prop">�v���p�e�B��</param>
''' <returns>
''' �t�B�[���h�R�[�h�ɑΉ�����w�肵���v���p�e�B
''' options�̏ꍇ�͔z��ŕԂ�
''' </returns>
Public Function SubtableFieldProperty(ByVal fcode As String, ByVal stblfcode As String, ByVal prop As String) As Variant
    Dim subfields As Variant
    Dim i As Long
    subfields = Me.FieldProperty(fcode, "fields")
    For i = LBound(subfields) To UBound(subfields)
        If subfields(i)("code") = stblfcode Then
            If prop = "options" Then '�h���b�v�_�E���̑I�����Ȃ�
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
''' <param name="existsPropOnly">$id�A$revision�Ȃ�AppFCodes�Ŏ擾�ł��Ȃ����̂͏����ꍇ��True</param>
''' <returns>�擾�������R�[�h�̃t�B�[���h�R�[�h�̔z��</returns>
Public Function RecordFields(Optional ByVal existsPropOnly As Boolean = True) As String()
    Dim resultFields() As String
    Dim f As Variant
    Dim isTarget As Boolean
    
    If Me.RecordCount > 0 Then
        '�s�v�Ȃ��̂�����
        For Each f In kvu_Recordset(0).Keys
            If existsPropOnly Then
                isTarget = kvu_AppFields("properties").Exists(f)
            Else
                isTarget = True
            End If
            If isTarget Then
                If Sgn(resultFields) = 0 Then
                    ReDim resultFields(0)
                Else
                    ReDim Preserve resultFields(UBound(resultFields) + 1)
                End If
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
''' <param name="index">���R�[�h�ʒu</param>
''' <param name="fcode">�t�B�[���h�R�[�h</param>
''' <returns>
''' �w�肳�ꂽindex�̈ʒu�̃f�[�^
''' �l������������̂ɂ��ẮAIsJoinArrayData�v���p�e�B��False�̏ꍇ�A�z��B�f�[�^�����̏ꍇ�́A�󕶎�
''' SUBTABLE�̓R���N�V�����^�̔z��B�f�[�^�����̏ꍇ�́A�󕶎�
''' </returns>
Public Function FieldValue(ByVal index As Long, ByVal fcode As String) As Variant
    Dim i As Long
    Dim subtableArray() As Collection
    Dim subtableRecCnt As Long  '�T�u�e�[�u�����̃��R�[�h����
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
        
                    '�߂�l�̔z������R�[�h�����Ŋm��
                    ReDim subtableArray(0 To subtableRecCnt - 1)
                    Dim dt As Dictionary
                    Dim colRec As Collection
                    
                    '�T�u�e�[�u���̃t�B�[���h�ꗗ���擾
                    Dim s As Variant
                    '�T�u�e�[�u���̃��R�[�h���������[�v
                    For i = 1 To subtableRecCnt
                        Set colRec = New Collection
                        '�e���ڒl���R���N�V�����ɃZ�b�g
                        For Each s In kvu_Recordset(index)(fcode)("value")(1)("value").Keys '1����X�^�[�g
                            Set dt = kvu_Recordset(index)(fcode)("value")(i)("value")(s)
                            colRec.Add Item:=kvu_GetFieldValue(dt("type"), dt("value")), Key:=s
                        Next
                        '�R���N�V�����̔z��ɃZ�b�g
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
''' <param name="addFieldLabel">�z��̍ŏ��Ƀt�B�[���h���x����ǉ����邩</param>
''' <param name="existsPropOnly">$id�A$revision�Ȃ�AppFCodes�Ŏ擾�ł��Ȃ����̂͏����ꍇ��True</param>
''' <returns>
''' �擾�f�[�^�S�Ă̓񎟌��z��i�T�u�e�[�u���̒l�͎擾���Ȃ��A�����I���n�̓J���}��؂�j
''' </returns>
Public Function RecordArray(Optional ByVal addFieldLabel As Boolean = True, Optional ByVal existsPropOnly As Boolean = True) As Variant
    Dim varData As Variant
    Dim i As Long, k As Long
    Dim adj As Integer: adj = IIf(addFieldLabel, 1, 0)
    Dim fieldsOrg() As String
    Dim fields() As String  '�T�u�e�[�u�����O
    Dim fVal As Variant
    
    If Me.RecordCount > 0 Then
        '���R�[�h�̃t�B�[���h�ꗗ�擾���A�T�u�e�[�u��������
        fieldsOrg = Me.RecordFields(existsPropOnly)
        For k = LBound(fieldsOrg) To UBound(fieldsOrg)
            If Me.FieldType(CStr(fieldsOrg(k))) <> "SUBTABLE" Then
                If Sgn(fields) = 0 Then
                    ReDim fields(0)
                Else
                    ReDim Preserve fields(UBound(fields) + 1)
                End If
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
                    '�����I���A�`�F�b�N�{�b�N�X�A���[�U�[�I���ȂǕ�������ꍇ�́A�J���}��؂��
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
''' <param name="isAll">�S���R�[�h�擾����ꍇ��True�ALimit���w�肷��ꍇ��False</param>
''' <param name="sFields">�擾�t�B�[���h�B�����̏ꍇ�̓J���}��؂��</param>
''' <param name="query">limit��offset�͊܂߂��N���X�̃v���p�e�B�Őݒ�</param>
''' <returns>
''' �f�[�^�ʂ������ꍇ�ɑS���擾����ƁAPC�ɂ���Ă̓������s���ɂȂ�܂��B
''' </returns>
Public Sub GetRecords(ByVal isAll As Boolean, Optional ByVal sFields As String = "", Optional ByVal query As String = "")
    Dim strQuery As String
    Dim jsonData As Dictionary
    Dim returnData() As Dictionary
    Dim dataCount As Long
    Dim offsetIdx As Long
    Dim i As Long
    Dim isFirst As Boolean: isFirst = True
    
    '�N���A
    Call kvu_ClearRecordset
    Erase returnData
    
    'limit�Ŏw�肵�������iMAX500���j�������Ȃ�
    offsetIdx = IIf(isAll, 0, Me.Offset)
    dataCount = Me.Limit

    Do Until dataCount < Me.Limit
        '�N�G���̍쐬
        strQuery = IIf(sFields <> "", "&fields=" & knt_EncodeUrl(sFields), "")
        'WebHelper��UrlEncode�͓��{��NG
        strQuery = strQuery & "&query=" & knt_EncodeUrl(query & " limit " & Me.Limit & " offset " & offsetIdx)   'limit 500 offset 0
        
        Set jsonData = kvu_Execute(kvu_RecordsUrl & strQuery & "&totalCount=true")
        
        If Not jsonData.Exists("records") Then
            Dim errMsg As String
            If jsonData.Exists("message") Then
               errMsg = " - " & jsonData("message")
            End If
            Err.Raise KTN_VBA_EXCEPTION.RECORD_NOT_FOUND, Description:=getErrorMessage(KTN_VBA_EXCEPTION.RECORD_NOT_FOUND) & "�@" & errMsg & "(GetRecords)"
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
''' ���b�`�G�f�B�^�̃^�O����
''' </summary>
''' <param name="str"></param>
''' <param name="typ"></param>
''' <returns></returns>
Public Function RemoveTags(ByVal str As String, ByVal typ As RemoveTagsEnum) As String
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    Dim targetStr As String: targetStr = str
    If typ = RemoveTagsEnum.RemoveChangeNewLine Then
        targetStr = Replace(targetStr, "<br />", "<br />" & vbCrLf, compare:=vbTextCompare)
        targetStr = Replace(targetStr, "</div>", "</div>" & vbCrLf, compare:=vbTextCompare)
    End If
    regEx.Pattern = "<[^>]*>"
    regEx.Global = True
    RemoveTags = regEx.Replace(targetStr, "")
 End Function
 
''' <summary>
''' UTC��Loacl�����ɕύX
''' </summary>
''' <param name="dtUtc">2012-01-11T11:30:00Z �̃t�H�[�}�b�g</param>
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
''' �A�v�����̎擾
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
''' �t�B�[���h���̎擾
''' </summary>
Private Sub kvu_GetFieldConf()
    Dim fieldList() As String
    Dim subtfields() As String
    Dim i As Long, k As Long
    
    Set kvu_AppFields = kvu_Execute(kvu_FieldsUrl)
    
    '�T�u�e�[�u���̏���ۑ�
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
''' �f�[�^�擾���ʏ���
''' </summary>
''' <param name="url"></param>
''' <returns></returns>
Private Function kvu_Execute(ByVal url As String) As Object
    Dim objHttpReq As Object    'XMLHTTP �I�u�W�F�N�g
    Dim strJson As String       '���X�|���X�Ŏ󂯎��JSON�f�[�^
    
    If kvu_SubDomain = "" Or kvu_AppId = "" Then
        Err.Raise KTN_VBA_EXCEPTION.SETTING_NOT_FOUND, Description:=getErrorMessage(KTN_VBA_EXCEPTION.SETTING_NOT_FOUND)
    End If
    
    Set objHttpReq = CreateObject("MSXML2.XMLHTTP")
    objHttpReq.Open "GET", url, False

    '���N�G�X�g�w�b�_�[�w��
    With objHttpReq
        .SetRequestHeader "X-Cybozu-Authorization", knt_EncodeBase64(Me.UserAuth)   '���O�C���F��
        .SetRequestHeader "Host", kvu_SubDomain & ".cybozu.com" + ":443"            '�h���C�����F�|�[�g�ԍ�
        .SetRequestHeader "If-Modified-Since", "Thu, 01 Jun 1970 00:00:00 GMT"      '�L���b�V���΍�i��Ƀ��X�|���X���擾�ł����Ԃɂ���j
    End With
    
    '���N�G�X�g���M
    objHttpReq.Send (Null)

    '���X�|���X�擾
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
    
    'Parse���ăf�[�^�Z�b�g
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
                            valueArry(i - 1) = rec(i)("name") '1����X�^�[�g,�t�@�C���̏ꍇ�̓t�@�C������Ԃ�
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
                        valueArry(i - 1) = rec(i) '1����X�^�[�g
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
''' �t�B�[���h�R�[�h�̐e�T�u�e�[�u��������
''' </summary>
''' <param name="stblfcode">�T�u�e�[�u�����t�B�[���h�̃t�B�[���h�R�[�h</param>
''' <returns>�e�T�u�e�[�u���̃t�B�[���h�R�[�h�A�Y������T�u�e�[�u�����Ȃ���΋󕶎���Ԃ�</returns>
Private Function kvu_FindParentSubtableFcode(ByVal stblfcode As String) As String
    If kvu_AppSubtableFields.Exists(stblfcode) Then
        kvu_FindParentSubtableFcode = kvu_AppSubtableFields(stblfcode)
    Else
        kvu_FindParentSubtableFcode = ""
    End If
End Function

''' <summary>
''' �A�v���擾��URL
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
''' ���R�[�h�擾��URL
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
''' �t�B�[���h�擾��URL
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
''' ���R�[�h�擾�O�N���A����
''' </summary>
Private Sub kvu_ClearRecordset()
    kvu_RecordCount = 0
    kvu_TotalRecordCount = 0
    Erase kvu_Recordset
End Sub
 
''' <summary>
''' Variant�z���String�z��ɕϊ�
''' </summary>
''' <param name="varArray"></param>
''' <returns>
''' �A�v���̃t�B�[���h�R�[�h�ꗗ�i�T�u�e�[�u���̒��܂œW�J�j�̔z��
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
    d.appendChild elm
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
            result = "���R�[�h�͎擾����Ă��܂���"
        Case FCODE_NOT_EXISTS
            result = "�w�肳�ꂽ�t�B�[���h�R�[�h�͑��݂��܂���"
        Case HTTP_REQ_ERR
            result = "�f�[�^�擾���G���["
        Case INVALID_UTC_DIFF
            result = "UTC�Ƃ̎��ԍ�������������܂���(+09:00�̌`���Ŏw��)"
        Case LIMIT_MAX
            result = "���~�b�g��1�`" & CStr(LIMIT_MAX_VAL) & "�Ŏw�肵�Ă�������"
        Case OFFSET_MIN
            result = "0�ȏ�̎w������Ă�������"
        Case PROP_NOT_EXISTS
            result = "�w�肳�ꂽ�v���p�e�B�͑��݂��܂���"
        Case PROP_NOT_FOUND
            result = "�v���p�e�B��񂪑��݂��܂���"
        Case RECORD_NOT_FOUND
            result = "records�����݂��܂���"
        Case SETTING_NOT_FOUND
            result = "�T�u�h���C���܂��̓A�v��ID���ݒ肳��Ă��܂���"
        Case NOT_SUBTABLE
            result = "�T�u�e�[�u���ł͂���܂���"
        Case UNSUPPORTED_SUBTABLE
            result = "���̃��\�b�h�̓T�u�e�[�u�����t�B�[���h�ɑΉ����Ă��܂���"
    End Select
    getErrorMessage = result
End Function
Private Sub Class_Initialize()
    Me.Lang = ""
    Me.IsName = True
    Me.IsJoinArrayData = True
    Me.RemoveTag = RemoveTagsEnum.Original
    Me.Limit = LIMIT_MAX_VAL
    Me.Offset = 0
    Me.Timezone = "+09:00"  '�f�t�H���g����{
    Call kvu_ClearRecordset
End Sub

Private Sub Class_Terminate()
    Call kvu_ClearRecordset
    Set kvu_AppConf = Nothing
    Set kvu_AppFields = Nothing
    Set kvu_AppSubtableFields = Nothing
End Sub