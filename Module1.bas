Attribute VB_Name = "Module1"
Option Explicit

Sub MySub()
'    '�C���X�^���X�̐����Ə����ݒ�̓���m�F
'    Dim p As Person: Set p = New Person
'    p.Initialize Sheet1.ListObjects(1).ListRows(1).Range
'
'    Stop

'    'Age�v���p�e�B�̓���m�F
'    Dim p As Person: Set p = New Person
'    p.Initialize Sheet1.ListObjects(1).ListRows(1).Range
'
'    Debug.Print p.Age
    
'    '�f�[�^�̓ǂݏ����m�F
'    With Sheet1
'        .LoadData
'
'        With .Persons(1)
'            .Name = "�����@��"
'            .Birthday = #3/30/1988#
'            .Active = True
'        End With
'
'        .ApplyData
'    End With

'    '���R�[�h�̍X�V�ƒǉ����m�F����
'    With Sheet1
'        .LoadData
'
'        Dim p As Person: Set p = New Person
'        With p
'            .Id = 1
'            .Name = "�����@��"
'            .Gender = "�j"
'            .Birthday = #3/30/1988#
'            .Active = True
'        End With
'
'        .UpdatePerson p
'
'        p.Id = .MaxId + 1
'        .AddPerson p
'
'        End With

    '�e�[�u���X�V�̎��s���x�𑪒肷��
    Dim start As Date: start = Time
    Call ShowUserForm
    Call Sheet1.ApplyData
    Unload UserForm1
    
    Dim finish As Date: finish = Time
    Debug.Print Minute(finish - start) * 60 + Second(finish - start)
        
End Sub

'**
'* ���[�U�[�t�H�[��UserForm1��\������
'* (Sheet1�̃R���g���[���{�^���u����Ǘ��v�Ƀ}�N���o�^�j
'*
Sub ShowUserForm()
    UserForm1.Show vbModeless
End Sub
