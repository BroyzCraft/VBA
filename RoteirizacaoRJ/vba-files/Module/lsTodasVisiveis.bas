Attribute VB_Name = "lsTodasVisiveis"

Global lPlanilhas()    As String
Global lArquivo        As String

'Faz com que todas as planilhas fiquem vis�veis
Public Sub lsTodasVisiveis()

Dim lWorksheet      As Worksheet
Dim lWorkbook       As Workbooks
Dim lCont           As Integer
Dim lActiveSheet    As Worksheet

    Set lworkbooks = ActiveWorkbook
    lCont = 0

    lArquivo = lworkbooks.Name
    Set lActiveSheet = ActiveSheet

    'Identifica as planilhas que n�o est�o vis�veis e as reexibe
    For Each lworksheets In lworkbooks.Worksheets

        If lworksheets.Visible <> xlSheetVisible Then
            'Redimensiona e mant�m as informas�es anteriores
            ReDim Preserve lPlanilhas(lCont)

            'Passa os dados para a vari�vel lPlanilhas e a deixa vis�vel
            lPlanilhas(lCont) = lworksheets.Name
            lworksheets.Visible = xlSheetVisible
            lCont = lCont + 1
        End If

    Next lworksheets

    lActiveSheet.Select

End Sub
