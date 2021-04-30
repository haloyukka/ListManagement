Attribute VB_Name = "Module1"
Option Explicit

Sub MySub()
'    'インスタンスの生成と初期設定の動作確認
'    Dim p As Person: Set p = New Person
'    p.Initialize Sheet1.ListObjects(1).ListRows(1).Range
'
'    Stop

'    'Ageプロパティの動作確認
'    Dim p As Person: Set p = New Person
'    p.Initialize Sheet1.ListObjects(1).ListRows(1).Range
'
'    Debug.Print p.Age
    
'    'データの読み書き確認
'    With Sheet1
'        .LoadData
'
'        With .Persons(1)
'            .Name = "横尾　勤"
'            .Birthday = #3/30/1988#
'            .Active = True
'        End With
'
'        .ApplyData
'    End With

'    'レコードの更新と追加を確認する
'    With Sheet1
'        .LoadData
'
'        Dim p As Person: Set p = New Person
'        With p
'            .Id = 1
'            .Name = "横尾　勤"
'            .Gender = "男"
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

    'テーブル更新の実行速度を測定する
    Dim start As Date: start = Time
    Call ShowUserForm
    Call Sheet1.ApplyData
    Unload UserForm1
    
    Dim finish As Date: finish = Time
    Debug.Print Minute(finish - start) * 60 + Second(finish - start)
        
End Sub

'**
'* ユーザーフォームUserForm1を表示する
'* (Sheet1のコントロールボタン「名簿管理」にマクロ登録）
'*
Sub ShowUserForm()
    UserForm1.Show vbModeless
End Sub
