VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   4575
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   4740
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Button1_Click()　　'登録Button Click するとUserFormに記載されている情報をセルに書き込む
    
    Dim lastRow As Long
    With Worksheets("sheet1")
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row + 1
        .Cells(lastRow, 1).Value = Me.TextBox1.Text
        .Cells(lastRow, 2).Value = Me.TextBox2.Text
        .Cells(lastRow, 3).Value = Me.TextBox3.Value
        .Cells(lastRow, 4).Value = Me.TextBox4.Value
        .Cells(lastRow, 5).Value = Me.TextBox5.Value
        .Cells(lastRow, 6).Value = Me.TextBox6.Value
    End With
End Sub
Private Sub ExInputCls()
    TextBox1.Text = ""
    TextBox2.Text = ""
    TextBox3.Text = ""
    TextBox4.Text = ""
    TextBox5.Text = ""
    TextBox6.Text = ""
    '日付テキストボックスにフォーカスを移す
    TextBox1.SetFocus
End Sub
Private Sub TextBox1_Change()
    
End Sub
Private Sub TextBox2_Change()
    
End Sub
Private Sub TextBox3_Change()
    
End Sub
Private Sub TextBox4_Change()
    
End Sub
Private Sub TextBox5_Change()
    
End Sub
Private Sub TextBox6_Change()

End Sub
Private Sub Button2_Click()
    Unload UserForm1
End Sub
Private Sub Button3_Click()
    ExInputCls
End Sub
