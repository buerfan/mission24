VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Data Input Form"
   ClientHeight    =   4490
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   9600.001
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Sheet2.Activate

Cells(Rows.Count, "A").End(xlUp).Offset(1, 0).Value = TextBox1.Value
Cells(Rows.Count, "B").End(xlUp).Offset(1, 0).Value = TextBox4.Value
Cells(Rows.Count, "C").End(xlUp).Offset(1, 0).Value = TextBox3.Value

Cells(Rows.Count, "F").End(xlUp).Offset(1, 0).Value = ComboBox1.Value





End Sub
