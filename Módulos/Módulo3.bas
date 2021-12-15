Attribute VB_Name = "Módulo3"

Sub Cadastro()
Attribute Cadastro.VB_Description = "Botão que realiza o cadastro dos clientes"
Attribute Cadastro.VB_ProcData.VB_Invoke_Func = "C\n14"
'
' Cadastro Macro
' Realiza um cadastro de Cliente
'
' Atalho do teclado: Ctrl+j
'
    Sheets("Lista de Clientes").Select
    Rows("3:3").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
    
    Sheets("Cadastro").Select
    Range("G7").Select
    Selection.Copy
    Sheets("Lista de Clientes").Select
    Range("A3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Sheets("Cadastro").Select
    Range("J7").Select
    Selection.Copy
    Sheets("Lista de Clientes").Select
    Range("B3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Sheets("Cadastro").Select
    Range("G9").Select
    Selection.Copy
    Sheets("Lista de Clientes").Select
    Range("C3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Sheets("Cadastro").Select
    Range("M9").Select
    Selection.Copy
    Sheets("Lista de Clientes").Select
    Range("D3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Sheets("Cadastro").Select
    Range("G11").Select
    Selection.Copy
    Sheets("Lista de Clientes").Select
    Range("E3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Sheets("Cadastro").Select
    Range("J11").Select
    Selection.Copy
    Sheets("Lista de Clientes").Select
    Range("F3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Sheets("Cadastro").Select
    Range("L11").Select
    Selection.Copy
    Sheets("Lista de Clientes").Select
    Range("G3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Sheets("Cadastro").Select
    Range("N11").Select
    Selection.Copy
    Sheets("Lista de Clientes").Select
    Range("H3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Sheets("Cadastro").Select
    Range("G13").Select
    Selection.Copy
    Sheets("Lista de Clientes").Select
    Range("I3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Sheets("Cadastro").Select
    Range("J13").Select
    Selection.Copy
    Sheets("Lista de Clientes").Select
    Range("J3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Sheets("Cadastro").Select
    Range("L13").Select
    Selection.Copy
    Sheets("Lista de Clientes").Select
    Range("K3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Sheets("Cadastro").Select
    Range("G7").Select
    Selection.ClearContents
    Range("J7").Select
    Selection.ClearContents
    Range("G9").Select
    Selection.ClearContents
    Range("M9").Select
    Selection.ClearContents
    Range("G11").Select
    Selection.ClearContents
    Range("J11").Select
    Selection.ClearContents
    Range("L11").Select
    Selection.ClearContents
    Range("N11").Select
    Selection.ClearContents
    Range("G13").Select
    Selection.ClearContents
    Range("J13").Select
    Selection.ClearContents
    Range("L13").Select
    Selection.ClearContents
    Range("G7").Select
End Sub
Sub Classificar()
Attribute Classificar.VB_Description = "Botão que classifica os clientes de A-Z"
Attribute Classificar.VB_ProcData.VB_Invoke_Func = "R\n14"
'
' Classificar Macro
' Classifica os Clientes de A-Z
'
' Atalho do teclado: Ctrl+r
'
    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("Lista de Clientes").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lista de Clientes").Sort.SortFields.Add2 Key:= _
        Range("A3:A6"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Lista de Clientes").Sort
        .SetRange Range("A2:J6")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("A2").Select
End Sub
Sub Filtrar()
Attribute Filtrar.VB_Description = "Botão que aplica filtros na tabela com os dados dos clientes"
Attribute Filtrar.VB_ProcData.VB_Invoke_Func = "F\n14"
'
' Filtrar Macro
' Aplica filtros na tabela
'
' Atalho do teclado: Ctrl+f
'
    Range("A2").Select
    Selection.AutoFilter
End Sub
Sub Voltar()
Attribute Voltar.VB_Description = "Botão que volta para a página inicial"
Attribute Voltar.VB_ProcData.VB_Invoke_Func = "V\n14"
'
' Voltar Macro
' Volta para a página inicial
'
' Atalho do teclado: Ctrl+v
'
    Sheets("Cadastro").Select
    Range("G7:H7").Select
End Sub
