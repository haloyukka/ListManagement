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
'* Personクラスのインスタンスに初期値を設定する
'*
'* @param myRange {Range} 設定するレコードのセル範囲
'*
Public Sub Initialize(ByVal myRange As Range)
    Id = myRange(eId).Value
    Name = myRange(eName).Value
    Gender = myRange(eGender).Value
    Birthday = myRange(eBirthday).Value
    Active = myRange(eActive).Value
End Sub

'**
'* Personクラスの年齢を取得する
'*
'* @return {Long} 年齢
'*
Public Property Get Age() As Long
    Dim myAge As Long
    myAge = DateDiff("yyyy", Me.Birthday, Date)
    
    If Date < DateSerial(Year(Now), Month(Me.Birthday), Day(Me.Birthday)) Then
        myAge = myAge - 1
    End If
    
    Age = myAge
End Property
