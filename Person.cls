VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Person"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Id As Long
Public Name As String
Public Gender As String
Public Birthday As Date
Public Active As Boolean


'**
'* Person�N���X�̃C���X�^���X�ɏ����l��ݒ肷��
'*
'* @param myRange {Range} �ݒ肷�郌�R�[�h�̃Z���͈�
'*
Public Sub Initialize(ByVal myRange As Range)
    Id = myRange(eId).Value
    Name = myRange(eName).Value
    Gender = myRange(eGender).Value
    Birthday = myRange(eBirthday).Value
    Active = myRange(eActive).Value
End Sub

'**
'* Person�N���X�̔N����擾����
'*
'* @return {Long} �N��
'*
Public Property Get Age() As Long
    Dim myAge As Long
    myAge = DateDiff("yyyy", Me.Birthday, Date)
    
    If Date < DateSerial(Year(Now), Month(Me.Birthday), Day(Me.Birthday)) Then
        myAge = myAge - 1
    End If
    
    Age = myAge
End Property