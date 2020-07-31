Attribute VB_Name = "AddProduct_Mod"
Function AddProduct(code As String)
    
    If (AddProduct_Mod.SearchProduct(code) <> 0) Then
        MsgBox ("Product already present" & vbNewLine & "Code: " & code)
        Exit Function
    End If
        
    Sheets("Products").Select
    
    lastRow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
    
    With ActiveSheet
        .Cells(1, 1).Value = "Code"
        .Cells(1, 2).Value = "Item"
        .Cells(1, 3).Value = "Type"
        .Cells(1, 4).Value = "Flanch"
        .Cells(1, 5).Value = "Model"
        .Cells(1, 6).Value = "Diameter"
        .Cells(1, 7).Value = "Length"
        
        .Cells(lastRow + 1, 1).Value = code
        .Cells(lastRow + 1, 2).Value = AddProduct_Ufm.item_cmbx.Value
        .Cells(lastRow + 1, 3).Value = AddProduct_Ufm.type_cmbx.Value
        .Cells(lastRow + 1, 4).Value = AddProduct_Ufm.flanch_cbx.Value
        .Cells(lastRow + 1, 5).Value = AddProduct_Ufm.model_cmbx.Value
        .Cells(lastRow + 1, 6).Value = AddProduct_Ufm.diameter_cmbx.Value
        .Cells(lastRow + 1, 7).Value = AddProduct_Ufm.length_cmbx.Value
    End With
    
    Call fitAndFormat_Mod.FitAndFormat("Products")
    
    MsgBox ("Product added succesfully" & vbNewLine & "Generated Code: " & code)
    
End Function

Function SearchProduct(code As String)

    Sheets("Products").Select
    
    lastRowAddress = Range("A10000").End(xlUp).Address
    lastRow = Range("A10000").End(xlUp).Row
    
'   If there is only one row present
    If (lastRow = 1) Then
        rowNum = 0
    Else
'       If record doesn't exist
        If (IsError(Application.Match(code, Range("$A$1 : " & lastRowAddress), 0))) Then
            rownNum = 0
        Else
'       If record exists
            rowNum = Application.Match(code, Range("$A$1 : " & lastRowAddress), 0)
        End If
    End If
    
    SearchProduct = rowNum

End Function
