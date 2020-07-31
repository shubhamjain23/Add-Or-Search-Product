VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddOrSearch_Ufm 
   Caption         =   "Add or Search"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "AddOrSearch_Ufm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddOrSearch_Ufm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AddAProduct_btn_Click()

    AddProduct_Ufm.Show

End Sub

Private Sub searchAProduct_btn_Click()

    Search_ufm.Show

End Sub
