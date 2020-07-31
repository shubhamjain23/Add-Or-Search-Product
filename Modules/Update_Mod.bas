Attribute VB_Name = "Update_Mod"
Option Explicit
Sub UpdateItemCmbx(formName As UserForm)
    
    Sheets("Data").Select
    
'    Updating values to comboBox
    With ActiveSheet
        formName.item_cmbx.RowSource = .Range(.Range("A2"), .Cells(Rows.Count, "A").End(xlUp)).Address(, , , True)
    End With
    
End Sub

Sub UpdateModelCmbx(formName As UserForm)

    Sheets("Data").Select
    
'    Updating values to comboBox
    With ActiveSheet
        formName.model_cmbx.RowSource = .Range(.Range("G2"), .Cells(Rows.Count, "G").End(xlUp)).Address(, , , True)
    End With
    
End Sub

Sub UpdateTypeCmbx(formName As UserForm)

    Sheets("Data").Select
    
'    Updating values to comboBox
    With ActiveSheet
        formName.type_cmbx.RowSource = .Range(.Range("C2"), .Cells(Rows.Count, "C").End(xlUp)).Address(, , , True)
    End With
    
    
End Sub

Sub UpdateDiameterCmbx(formName As UserForm)

    Sheets("Data").Select
    
'    Updating values to comboBox
    With ActiveSheet
        formName.diameter_cmbx.RowSource = .Range(.Range("I2"), .Cells(Rows.Count, "I").End(xlUp)).Address(, , , True)
    End With
    
End Sub

Sub UpdateLengthCmbx(formName As UserForm)
    
    Sheets("Data").Select
    
'    Updating values to comboBox
    With ActiveSheet
        formName.length_cmbx.RowSource = .Range(.Range("K2"), .Cells(Rows.Count, "K").End(xlUp)).Address(, , , True)
    End With
    
End Sub

