Attribute VB_Name = "Módulo2"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    ActiveSheet.PivotTables("Tabela dinâmica4").ChangePivotCache ActiveWorkbook. _
        PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "C:\Users\Bruno.marques\Desktop\AVALIAÇÃO DO PEDRÃO\[Relatorio de Avaliação.xlsm]DADOS - SERVICOS!C1:C20" _
        , Version:=7)
    Range("E43").Select
    ActiveSheet.PivotTables("Tabela dinâmica4").PivotCache.Refresh
    Range("D49").Select
    ActiveSheet.PivotTables("Tabela dinâmica4").PivotCache.Refresh
    ActiveWorkbook.Save
    ActiveWindow.SmallScroll Down:=-48
    Range("C7").Select
    ActiveSheet.PivotTables("Tabela dinâmica1").PivotCache.Refresh
    Range("D12").Select
    ActiveWorkbook.Save
    ActiveWorkbook.Save
    Range("C17").Select
End Sub
