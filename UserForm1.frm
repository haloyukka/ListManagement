VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "����Ǘ�"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'**
'* �C�x���g�v���V�[�W��: UserForm_Initialize
'*
Private Sub UserForm_Initialize()
    Call Sheet1.LoadData
    Call LoadIdList
End Sub

'**
'* �C�x���g�v���V�[�W��: CommandButtonUpdate_Click
'*
Private Sub CommandButtonUpdate_Click()
    If CheckFields Then
        Dim p As Person: Set p = New Person
        
        p.Name = TextBoxName.Text
        p.Birthday = TextBoxBirthday.Value
        p.Gender = "��"
        If OptionButtonMale.Value = True Then p.Gender = "�j"
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
'* �C�x���g�v���V�[�W��: CommandButtonClose_Click
'*
Private Sub CommandButtonClose_Click()
    Unload Me
End Sub

'**
'* �R���{�{�b�N�XComboBoxId�̃��X�g��ǂݍ���
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
'* �C�x���g�v���V�[�W��: ComboBoxId_Change
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
'* �R���{�{�b�N�XComboBoxId�̒l���K�����ǂ���
'*
'* @return {Boolean} �R���{�{�b�N�XComboBoxId�̒l��1�ȏ�ID�̍ő�ȉ��A�܂���"New"���ǂ���\
'*
Private Property Get IsValidId() As Boolean
    IsValidId = False
    With ComboBoxId
        If (.Value > 0 And .Value <= Sheet1.MaxId) Or (.Value = "New") Then IsValidId = True
    End With
    
End Property

'**
'* �e�R���g���[���̒l�Ƃ��Ďw�肵��ID�̃��R�[�h�f�[�^���Ăяo��
'*
'* @return myId {Long} �Ăяo�����R�[�h��ID
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
'* ���ʂ�\��������("�j"�܂���"��")�����ƂɃI�v�V�����{�^���̒l��ݒ肷��
'*
'* @param myGender {String} ���ʂ�\��������
'*
Private Sub SetGender(ByVal myGender As String)

    OptionButtonFemale.Value = True
    If myGender = "�j" Then OptionButtonMale.Value = True
End Sub

'**
'* �e�R���g���[���̒l���N���A����
'*
Private Sub ClearFields()
    TextBoxName.Value = ""
    OptionButtonMale.Value = ""
    TextBoxBirthday.Value = ""
    LabelAge.Caption = ""
    CheckBoxActive.Value = ""
End Sub

'**
'* �e�R���g���[���̒l�����������͂���Ă��邩�ǂ����𔻒肷��
'*
'* @return {Boolean} ���ׂẴR���g���[���̒l�����������͂���Ă��邩�ǂ���
'*
Private Function CheckFields() As Boolean
    
    CheckFields = True
    
    If Not IsValidId Then
        MsgBox "�uID�v��1�ȏ�̍ő�l�ȉ��̐��l�܂���""New""����͂��Ă�������", vbInformation
        CheckFields = False
    End If
    
    If Len(TextBoxName.Text) = 0 Then
        MsgBox "�u���O�v�ɓ��͂��Ă�������", vbInformation
        CheckFields = False
    End If
    
    If Not IsDate(TextBoxBirthday.Value) Then
        MsgBox "�u�a�����v�ɓ��t����͂��Ă�������", vbInformation
        CheckFields = False
    End If
    
End Function
