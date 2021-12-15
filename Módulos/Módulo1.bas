Attribute VB_Name = "Módulo1"
Sub Limpar()
Attribute Limpar.VB_Description = "Botão que limpa as informações digitadas"
Attribute Limpar.VB_ProcData.VB_Invoke_Func = "L\n14"
'
' Limpar Macro
' Limpa o cadastro
'
' Atalho do teclado: Ctrl+l
'
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
