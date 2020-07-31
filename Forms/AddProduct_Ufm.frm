VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddProduct_Ufm 
   Caption         =   "Add Product"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8790.001
   OleObjectBlob   =   "AddProduct_Ufm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddProduct_Ufm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub generateCode_btn_Click()

    If (type_cmbx.Value = "" Or model_cmbx.Value = "" Or item_cmbx.Value = "" Or diameter_cmbx.Value = "" Or length_cmbx.Value = "") Then
        MsgBox "One or more argument is missing."
        Exit Sub
    End If

    Dim code As String
    code = GenerateCode_Mod.GenerateCode(AddProduct_Ufm)
    Call AddProduct_Mod.AddProduct(code)
    
    AddProduct_Ufm.flanch_cbx.Value = False

End Sub

Private Sub UserForm_Initialize()

    AddProduct_Ufm.flanch_cbx.Value = False
    
    Call Update_Mod.UpdateItemCmbx(AddProduct_Ufm)
    Call Update_Mod.UpdateModelCmbx(AddProduct_Ufm)
    Call Update_Mod.UpdateTypeCmbx(AddProduct_Ufm)
    Call Update_Mod.UpdateDiameterCmbx(AddProduct_Ufm)
    Call Update_Mod.UpdateLengthCmbx(AddProduct_Ufm)

End Sub
