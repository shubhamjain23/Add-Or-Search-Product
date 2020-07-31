Attribute VB_Name = "GenerateCode_Mod"
Function GenerateCode(formName As UserForm)
    
'   If Error, then proceed as default
    On Error GoTo 0
    
    Sheets("Data").Select

    Range("A1").Select
    Selection.CurrentRegion.Select

'   Splitting string (address) by current region
    Address = Split(Selection.Address, ":")
    
'   assigning range to a variable
    data_table = Range(Address(0), Address(1))

    Dim lastColumnAddress As String
'   Finding last column address
    lastColumnAddress = Range("XFD1").End(xlToLeft).Address
    
    With Application
    
'       Using Index(Match) by first finding the columnNumber (say 2), then columnName(B), then indexing
    
'       Item
        columnNum_item = .Match("Item", Range("$A$1 : " & lastColumnAddress), 0)
        columnName_item = Split(Cells(1, (.Match("Item", Range("$A$1 : " & lastColumnAddress), 0))).Address, "$")(1)
        item = .Index(data_table, .Match(formName.item_cmbx.Value, Range("$" & columnName_item & "$1 : $" & columnName_item & "$1000"), 0), (columnNum_item + 1))
        
'       Model
        columnNum_model = .Match("Model", Range("$A$1 : " & lastColumnAddress), 0)
        columnName_model = Split(Cells(1, (.Match("Model", Range("$A$1 : " & lastColumnAddress), 0))).Address, "$")(1)
        model = .Index(data_table, .Match(formName.model_cmbx.Value, Range("$" & columnName_model & "$1 : $" & columnName_model & "$1000"), 0), (columnNum_model + 1))

'       Type
        columnNum_type = .Match("Type", Range("$A$1 : " & lastColumnAddress), 0)
        columnName_type = Split(Cells(1, (.Match("Type", Range("$A$1 : " & lastColumnAddress), 0))).Address, "$")(1)
        typ = .Index(data_table, .Match(formName.type_cmbx.Value, Range("$" & columnName_type & "$1 : $" & columnName_type & "$1000"), 0), (columnNum_type + 1))
        
'       Length
        columnNum_length = .Match("Length", Range("$A$1 : " & lastColumnAddress), 0)
        columnName_length = Split(Cells(1, (.Match("Length", Range("$A$1 : " & lastColumnAddress), 0))).Address, "$")(1)
        length = .Index(data_table, .Match(formName.length_cmbx.Value, Range("$" & columnName_length & "$1 : $" & columnName_length & "$1000"), 0), (columnNum_length + 1))
        
'       Flanch
        columnNum_flanch = .Match("Flanch", Range("$A$1 : " & lastColumnAddress), 0)
        columnName_flanch = Split(Cells(1, (.Match("Flanch", Range("$A$1 : " & lastColumnAddress), 0))).Address, "$")(1)
        flanch = .Index(data_table, .Match(formName.flanch_cbx.Value, Range("$" & columnName_flanch & "$1 : $" & columnName_flanch & "$1000"), 0), (columnNum_flanch + 1))
        
'       Diameter
        columnNum_diameter = .Match("Diameter", Range("$A$1 : " & lastColumnAddress), 0)
        columnName_diameter = Split(Cells(1, (.Match("Diameter", Range("$A$1 : " & lastColumnAddress), 0))).Address, "$")(1)
        diameter = .Index(data_table, .Match(formName.diameter_cmbx.Value, Range("$" & columnName_diameter & "$1 : $" & columnName_diameter & "$1000"), 0), (columnNum_diameter + 1))
        
    End With
    
'   Generating code
    code = item & typ & flanch & "-" & model & "-" & diameter & length
    
    GenerateCode = code
    
End Function
