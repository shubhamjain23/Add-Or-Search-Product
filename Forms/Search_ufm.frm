VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Search_ufm 
   Caption         =   "Search a Product"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8805.001
   OleObjectBlob   =   "Search_ufm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Search_ufm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SearchProduct_btn_Click()
    
    If (type_cmbx.Value = "" Or model_cmbx.Value = "" Or item_cmbx.Value = "" Or diameter_cmbx.Value = "" Or length_cmbx.Value = "") Then
        MsgBox "One or more argument is missing."
        Exit Sub
    End If
    
    Dim code As String
    code = GenerateCode_Mod.GenerateCode(Search_ufm)
    rowNum = AddProduct_Mod.SearchProduct(code)
    If (rowNum = 0) Then
        MsgBox ("Product not present" & vbNewLine & "Try adding product instead.")
    Else
        MsgBox ("Product already present in Row " & rowNum & " in " & """Products""" & " Sheet." & vbNewLine & "Code : " & code)
    End If
    
    Search_ufm.flanch_cbx.Value = False
    
End Sub

Private Sub UserForm_Initialize()

    flanch_cbx.Value = False
    
    Call Update_Mod.UpdateItemCmbx(Search_ufm)
    Call Update_Mod.UpdateModelCmbx(Search_ufm)
    Call Update_Mod.UpdateTypeCmbx(Search_ufm)
    Call Update_Mod.UpdateDiameterCmbx(Search_ufm)
    Call Update_Mod.UpdateLengthCmbx(Search_ufm)
    
End Sub
