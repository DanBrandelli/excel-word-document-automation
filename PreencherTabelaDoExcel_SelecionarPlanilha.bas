Option Explicit

' Macro para preencher a primeira tabela do documento Word
' selecionando a planilha Excel em uma janela.
'
' Estrutura esperada no Excel:
' Linha 1: título
' Linha 2: cabeçalhos
' Linha 3 em diante: dados
' Colunas A-D: Código, Produto, Quantidade e Lote

Sub PreencherTabelaDoExcel()

    Dim seletor As Object
    Dim caminhoExcel As String
    Dim xlApp As Object
    Dim xlWb As Object
    Dim xlWs As Object
    Dim tbl As Table
    Dim ultimaLinha As Long
    Dim linhaExcel As Long
    Dim linhaWord As Long
    Dim coluna As Long
    Dim valor As String

    On Error GoTo TrataErro

    ' 1. Selecionar a planilha Excel
    Set seletor = Application.FileDialog(3) ' 3 = seletor de arquivos

    With seletor
        .Title = "Selecione a planilha Excel com os dados"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Planilhas Excel", "*.xlsx; *.xlsm; *.xls"
        If .Show <> -1 Then
            MsgBox "Operação cancelada. Nenhuma planilha foi selecionada.", vbInformation
            Exit Sub
        End If
        caminhoExcel = .SelectedItems(1)
    End With

    ' 2. Verificar se o documento possui uma tabela
    If ActiveDocument.Tables.Count = 0 Then
        MsgBox "O documento Word não possui tabela para preencher.", vbExclamation
        Exit Sub
    End If

    Set tbl = ActiveDocument.Tables(1)

    ' É necessário ter uma linha de cabeçalho e uma linha-modelo
    If tbl.Rows.Count < 2 Then
        MsgBox "A primeira tabela precisa ter cabeçalho na linha 1 e uma linha-modelo na linha 2.", vbExclamation
        Exit Sub
    End If

    ' 3. Abrir o Excel em segundo plano
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = False

    Set xlWb = xlApp.Workbooks.Open(caminhoExcel, 0, True)
    Set xlWs = xlWb.Worksheets(1)

    ' 4. Encontrar a última linha preenchida na coluna A
    ultimaLinha = xlWs.Cells(xlWs.Rows.Count, 1).End(-4162).Row ' -4162 = xlUp

    If ultimaLinha < 3 Then
        MsgBox "Não foram encontrados dados a partir da linha 3 da planilha.", vbExclamation
        GoTo Finalizar
    End If

    ' 5. Excluir linhas antigas do Word, mantendo cabeçalho e linha-modelo
    For linhaWord = tbl.Rows.Count To 3 Step -1
        tbl.Rows(linhaWord).Delete
    Next linhaWord

    ' 6. Copiar os dados do Excel para o Word
    linhaWord = 2

    For linhaExcel = 3 To ultimaLinha

        If linhaExcel > 3 Then
            tbl.Rows.Add
            linhaWord = linhaWord + 1
        End If

        For coluna = 1 To 4
            valor = CStr(xlWs.Cells(linhaExcel, coluna).Text)
            tbl.Cell(linhaWord, coluna).Range.Text = valor & Chr(13) & Chr(7)
        Next coluna

    Next linhaExcel

    MsgBox "Tabela atualizada com sucesso!" & vbCrLf & _
           "Registros importados: " & (ultimaLinha - 2), vbInformation

Finalizar:
    On Error Resume Next
    If Not xlWb Is Nothing Then xlWb.Close SaveChanges:=False
    If Not xlApp Is Nothing Then xlApp.Quit
    Set xlWs = Nothing
    Set xlWb = Nothing
    Set xlApp = Nothing
    Set seletor = Nothing
    Exit Sub

TrataErro:
    MsgBox "Não foi possível concluir a atualização." & vbCrLf & _
           "Detalhe: " & Err.Description, vbCritical
    Resume Finalizar

End Sub
