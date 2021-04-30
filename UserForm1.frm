VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "名簿管理"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'**
'* イベントプロシージャ: UserForm_Initialize
'*
Private Sub UserForm_Initialize()
    Call Sheet1.LoadData
    Call LoadIdList
End Sub

'**
'* イベントプロシージャ: CommandButtonUpdate_Click
'*
Private Sub CommandButtonUpdate_Click()
    If CheckFields Then
        Dim p As Person: Set p = New Person
        
        p.Name = TextBoxName.Text
        p.Birthday = TextBoxBirthday.Value
        p.Gender = "女"
        If OptionButtonMale.Value = True Then p.Gender = "男"
        If (IsNull(CheckBoxActive.Value)) Then
            p.Active = False
        Else
            p.Active = CheckBoxActive.Value
        End If
        
        If ComboBoxId.Value = "New" Then
            p.Id = Sheet1.MaxId + 1
            Call Sheet1.AddPerson(p)
        Else
            p.Id = ComboBoxId.Value
            Call Sheet1.UpdatePerson(p)
        End If
        
        Call LoadFields(p.Id)
        Call LoadIdList
    End If
    
End Sub

'**
'* イベントプロシージャ: CommandButtonClose_Click
'*
Private Sub CommandButtonClose_Click()
    Unload Me
End Sub

'**
'* コンボボックスComboBoxIdのリストを読み込む
'*
Private Sub LoadIdList()
    
    With Sheet1.ListObjects(1)
        If .ListRows.Count > 1 Then
            Dim Lists As Variant: Lists = .ListColumns(1).DataBodyRange
            ComboBoxId.List = Lists
        End If
    End With
    ComboBoxId.AddItem "New"
End Sub

'**
'* イベントプロシージャ: ComboBoxId_Change
'*
Private Sub ComboBoxId_Change()
    With ComboBoxId
        If IsValidId Then
            If IsNumeric(.Value) Then
                Call LoadFields(.Value)
            Else
                Call ClearFields
            End If
        End If
    End With
    
End Sub

'**
'* コンボボックスComboBoxIdの値が適正かどうか
'*
'* @return {Boolean} コンボボックスComboBoxIdの値が1以上IDの最大以下、または"New"かどうか\
'*
Private Property Get IsValidId() As Boolean
    IsValidId = False
    With ComboBoxId
        If (.Value > 0 And .Value <= Sheet1.MaxId) Or (.Value = "New") Then IsValidId = True
    End With
    
End Property

'**
'* 各コントロールの値として指定したIDのレコードデータを呼び出す
'*
'* @return myId {Long} 呼び出すレコードのID
'*
Private Sub LoadFields(ByVal myId As Long)
    
    With Sheet1.Persons.Item(myId)
        ComboBoxId.Value = myId
        TextBoxName.Value = .Name
        Call SetGender(.Gender)
        TextBoxBirthday.Value = .Birthday
        LabelAge.Caption = .Age
        CheckBoxActive.Value = .Active
    End With
    
End Sub

'**
'* 性別を表す文字列("男"または"女")をもとにオプションボタンの値を設定する
'*
'* @param myGender {String} 性別を表す文字列
'*
Private Sub SetGender(ByVal myGender As String)

    OptionButtonFemale.Value = True
    If myGender = "男" Then OptionButtonMale.Value = True
End Sub

'**
'* 各コントロールの値をクリアする
'*
Private Sub ClearFields()
    TextBoxName.Value = ""
    OptionButtonMale.Value = ""
    TextBoxBirthday.Value = ""
    LabelAge.Caption = ""
    CheckBoxActive.Value = ""
End Sub

'**
'* 各コントロールの値が正しく入力されているかどうかを判定する
'*
'* @return {Boolean} すべてのコントロールの値が正しく入力されているかどうか
'*
Private Function CheckFields() As Boolean
    
    CheckFields = True
    
    If Not IsValidId Then
        MsgBox "「ID」は1以上の最大値以下の数値または""New""を入力してください", vbInformation
        CheckFields = False
    End If
    
    If Len(TextBoxName.Text) = 0 Then
        MsgBox "「名前」に入力してください", vbInformation
        CheckFields = False
    End If
    
    If Not IsDate(TextBoxBirthday.Value) Then
        MsgBox "「誕生日」に日付を入力してください", vbInformation
        CheckFields = False
    End If
    
End Function
