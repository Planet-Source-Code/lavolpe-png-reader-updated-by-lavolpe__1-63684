VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Canvas"
   ClientHeight    =   855
   ClientLeft      =   1320
   ClientTop       =   1395
   ClientWidth     =   1965
   LinkTopic       =   "Form1"
   ScaleHeight     =   855
   ScaleWidth      =   1965
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Show
    frmPNGmenu.Show 0, Me
End Sub
