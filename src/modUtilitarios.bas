Attribute VB_Name = "modUtilitarios"
' ============================================================
' meuAddin - Macros Utilitárias
' Versão: 2.1
' Autor: Lobão
' Descrição: Macros de uso frequente com validação de entrada,
'            tratamento de erros e boas práticas de VBA.
'            Pensado para rodar a partir de um suplemento (.xlam):
'            todas as macros agem sobre o ARQUIVO ATIVO.
' ============================================================
Option Explicit

' ============================================================
' Constantes de cores (valor Long = R + G*256 + B*65536)
' Centraliza todas as cores para facilitar manutenção
' ============================================================
Private Const COR_HEADER_VERDE      As Long = 29210       ' RGB(26,114,0)  - verde do cabeçalho
Private Const COR_LINHA_ACENTO      As Long = 57344       ' RGB(0,224,0)   - linha de acento (linha 5)
Private Const COR_FUNDO_PAINEL      As Long = 1846041     ' RGB(25,43,28)  - fundo escuro do painel
Private Const COR_BORDA_VERDE       As Long = -16748006   ' RGB(26,114,0)  - bordas de LinhasVerdes
Private Const COR_BORDA_LATERAL     As Long = -14931175   ' RGB(25,43,28)  - bordas de BarraLateral
Private Const COR_SHAPE_GRAFICO     As Long = 1183744     ' RGB(0,16,18)   - caixa do gráfico (Layout2)

' Fonte padrão usada nos layouts. Se a fonte não estiver instalada,
' o Excel usa uma substituta. Troque aqui para "Calibri", "Aptos" etc.
Private Const FONTE_PADRAO          As String = "Simplon BP Regular"

' ============================================================
' Utilitários (funções auxiliares compartilhadas)
' ============================================================

' Verifica se uma planilha existe no workbook informado
Private Function WorksheetExists(sheetName As String, wb As Workbook) As Boolean
    Dim ws As Object
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    WorksheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

' Verifica se um Shape com determinado nome existe na planilha
Private Function ShapeExists(ws As Worksheet, shapeName As String) As Boolean
    Dim shp As Shape
    On Error Resume Next
    Set shp = ws.Shapes(shapeName)
    ShapeExists = Not shp Is Nothing
    On Error GoTo 0
End Function

' Exibe barra de progresso na StatusBar e atualiza a tela
Private Sub AtualizarProgresso(atual As Long, total As Long, descricao As String)
    Dim pct As Long
    pct = CLng((atual / total) * 100)
    Application.StatusBar = descricao & " - " & pct & "% concluído (" & atual & "/" & total & ")"
    DoEvents
End Sub

' Limpa a StatusBar e restaura o padrão
Private Sub LimparProgresso()
    Application.StatusBar = False
End Sub

' Garante que existe uma planilha ativa (e não um gráfico ou nada)
Private Function TemPlanilhaAtiva() As Boolean
    If ActiveWorkbook Is Nothing Then
        MsgBox "Abra um arquivo antes de executar.", vbExclamation, "Sem arquivo"
        Exit Function
    End If
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "Ative uma planilha (não um gráfico) antes de executar.", vbExclamation, "Aba inválida"
        Exit Function
    End If
    TemPlanilhaAtiva = True
End Function

' Aplica bordas padronizadas a um range com cor configurável
Private Sub AplicarBordas(rng As Range, corBorda As Long, _
                           Optional mostrarBordaEsquerda As Boolean = True)
    With rng.Borders
        .Item(xlDiagonalDown).LineStyle = xlNone
        .Item(xlDiagonalUp).LineStyle = xlNone

        ' Borda esquerda (opcional)
        If mostrarBordaEsquerda Then
            .Item(xlEdgeLeft).LineStyle = xlContinuous
            .Item(xlEdgeLeft).Color = corBorda
            .Item(xlEdgeLeft).TintAndShade = 0
            .Item(xlEdgeLeft).Weight = xlMedium
        Else
            .Item(xlEdgeLeft).LineStyle = xlNone
        End If

        ' Bordas fixas
        .Item(xlEdgeTop).LineStyle = xlContinuous
        .Item(xlEdgeTop).Color = corBorda
        .Item(xlEdgeTop).TintAndShade = 0
        .Item(xlEdgeTop).Weight = xlMedium

        .Item(xlEdgeBottom).LineStyle = xlContinuous
        .Item(xlEdgeBottom).Color = corBorda
        .Item(xlEdgeBottom).TintAndShade = 0
        .Item(xlEdgeBottom).Weight = xlMedium

        .Item(xlEdgeRight).LineStyle = xlContinuous
        .Item(xlEdgeRight).Color = corBorda
        .Item(xlEdgeRight).TintAndShade = 0
        .Item(xlEdgeRight).Weight = xlMedium

        ' Sem separadores verticais internos
        .Item(xlInsideVertical).LineStyle = xlNone

        ' Separador horizontal interno fino
        .Item(xlInsideHorizontal).LineStyle = xlContinuous
        .Item(xlInsideHorizontal).Color = corBorda
        .Item(xlInsideHorizontal).TintAndShade = 0
        .Item(xlInsideHorizontal).Weight = xlThin
    End With
End Sub

' ============================================================
' MACRO: Layout1
' Aplica layout base: sem grade, cabeçalho colorido, fonte padrão
' ============================================================
Sub Layout1()
    Dim ws As Worksheet
    If Not TemPlanilhaAtiva() Then Exit Sub
    Set ws = ActiveSheet

    On Error GoTo ErroLayout1
    Application.ScreenUpdating = False

    ' Esconde linhas de grade
    ActiveWindow.DisplayGridlines = False

    ' Cabeçalho A1:Z5
    With ws.Range("A1:Z5")
        .RowHeight = 15.75
        .Interior.Color = COR_HEADER_VERDE
    End With

    ' Fonte padrão de toda a planilha
    With ws.Cells.Font
        .Name = FONTE_PADRAO
        .Size = 12
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
    End With

    ' Coluna A estreita (separador visual)
    ws.Columns("A").ColumnWidth = 0.85

    ' Linha de acento A5:Z5
    With ws.Range("A5:Z5")
        .Interior.Color = COR_LINHA_ACENTO
        .RowHeight = 5
    End With

    ws.Range("A2").Select
    Application.StatusBar = "Layout1 aplicado em '" & ws.Name & "'"
    GoTo Fim

ErroLayout1:
    MsgBox "Erro em Layout1: " & Err.Description, vbCritical, "Erro"
Fim:
    Application.ScreenUpdating = True
End Sub

' ============================================================
' MACRO: Layout2
' Aplica layout completo com painel escuro e shapes de gráfico
' ============================================================
Sub Layout2()
    Dim ws As Worksheet
    Dim txtHeader As Shape
    Dim shpGrafico As Shape
    Dim txtGrafico As Shape

    If Not TemPlanilhaAtiva() Then Exit Sub
    Set ws = ActiveSheet

    On Error GoTo ErroLayout2
    Application.ScreenUpdating = False

    ' Cabeçalho A1:Z5
    With ws.Range("A1:Z5")
        .RowHeight = 15.75
        .Interior.Color = COR_HEADER_VERDE
    End With

    ' Linha de acento A5:Z5
    With ws.Range("A5:Z5")
        .Interior.Color = COR_LINHA_ACENTO
        .RowHeight = 5
    End With

    ' Coluna A estreita
    ws.Columns("A").ColumnWidth = 0.85

    ' Fundo painel A6:Z39
    ws.Range("A6:Z39").Interior.Color = COR_FUNDO_PAINEL

    ' --- Shape: Título do Header ---
    ' Remove se já existir (evita empilhamento em execuções repetidas)
    If ShapeExists(ws, "txtHeader") Then ws.Shapes("txtHeader").Delete
    Set txtHeader = ws.Shapes.AddTextbox( _
        msoTextOrientationHorizontal, 20.25, 2.25, 475.5, 30.75)
    With txtHeader
        .Name = "txtHeader"
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.Text = "Título Aqui"
        With .TextFrame2.TextRange.Font
            .Name = FONTE_PADRAO
            .Size = 24
            .Fill.ForeColor.ObjectThemeColor = xlThemeColorLight1
            .Fill.Solid
        End With
        .Line.Visible = msoFalse
        .Fill.Visible = msoFalse
    End With

    ' --- Shape: Caixa do gráfico ---
    If ShapeExists(ws, "shpGrafico") Then ws.Shapes("shpGrafico").Delete
    Set shpGrafico = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, 16.5, 126, 287.25, 151.5)
    With shpGrafico
        .Name = "shpGrafico"
        .Adjustments.Item(1) = 0.05776
        With .Fill
            .ForeColor.RGB = COR_SHAPE_GRAFICO
            .Solid
        End With
        .Line.Visible = msoFalse
    End With

    ' --- Shape: Título do gráfico ---
    If ShapeExists(ws, "txtGrafico") Then ws.Shapes("txtGrafico").Delete
    Set txtGrafico = ws.Shapes.AddTextbox( _
        msoTextOrientationHorizontal, 21.75, 71.25, 273.75, 20.25)
    With txtGrafico
        .Name = "txtGrafico"
        .TextFrame2.TextRange.Text = "Título do Gráfico Aqui"
        With .TextFrame2.TextRange.Font
            .Name = FONTE_PADRAO
            .Size = 14
            .Bold = msoTrue
            .Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
            .Fill.Solid
        End With
        .Line.Visible = msoFalse
        .Fill.Visible = msoFalse
    End With

    ws.Range("A2").Select
    Application.StatusBar = "Layout2 aplicado em '" & ws.Name & "'"
    GoTo Fim

ErroLayout2:
    MsgBox "Erro em Layout2: " & Err.Description, vbCritical, "Erro"
Fim:
    Application.ScreenUpdating = True
End Sub

' ============================================================
' MACRO: ExibirOcultarAbasComApoio
' Alterna visibilidade das abas que contêm "apoio" no nome
' Lógica: se TODAS estão visíveis -> oculta todas;
'         se qualquer uma está oculta -> exibe todas
' ============================================================
Sub ExibirOcultarAbasComApoio()
    Dim wb As Workbook
    Dim sh As Object
    Dim abasApoio() As String
    Dim count As Long
    Dim visiveisFora As Long
    Dim todasVisiveis As Boolean
    Dim i As Long

    If ActiveWorkbook Is Nothing Then
        MsgBox "Abra um arquivo antes de executar.", vbExclamation, "Sem arquivo"
        Exit Sub
    End If
    Set wb = ActiveWorkbook

    ' Coleta abas com "apoio" no nome e conta as visíveis que NÃO são de apoio
    For Each sh In wb.Sheets
        If InStr(1, sh.Name, "apoio", vbTextCompare) > 0 Then
            count = count + 1
            ReDim Preserve abasApoio(1 To count)
            abasApoio(count) = sh.Name
        ElseIf sh.Visible = xlSheetVisible Then
            visiveisFora = visiveisFora + 1
        End If
    Next sh

    ' Nenhuma aba encontrada
    If count = 0 Then
        MsgBox "Nenhuma aba com 'apoio' no nome foi encontrada.", _
               vbInformation, "Sem resultado"
        Exit Sub
    End If

    ' Verifica se TODAS estão visíveis
    todasVisiveis = True
    For i = 1 To count
        If wb.Sheets(abasApoio(i)).Visible <> xlSheetVisible Then
            todasVisiveis = False
            Exit For
        End If
    Next i

    ' O Excel exige pelo menos uma aba visível
    If todasVisiveis And visiveisFora = 0 Then
        MsgBox "Não é possível ocultar: todas as abas visíveis são de apoio." & vbCrLf & _
               "O Excel exige pelo menos uma aba visível.", vbExclamation, "Aviso"
        Exit Sub
    End If

    ' Aplica alternância
    For i = 1 To count
        wb.Sheets(abasApoio(i)).Visible = _
            IIf(todasVisiveis, xlSheetHidden, xlSheetVisible)
    Next i

    ' Feedback na StatusBar (sem interromper o fluxo com MsgBox)
    Application.StatusBar = count & " aba(s) com 'apoio' " & _
        IIf(todasVisiveis, "ocultadas", "exibidas") & " com sucesso."
End Sub

' ============================================================
' MACRO: LinhasVerdes
' Aplica bordas verdes à seleção (sem borda esquerda)
' ============================================================
Sub LinhasVerdes()
    If TypeName(Selection) <> "Range" Then
        MsgBox "Selecione um intervalo de células antes de executar.", _
               vbExclamation, "Seleção inválida"
        Exit Sub
    End If

    AplicarBordas Selection, COR_BORDA_VERDE, mostrarBordaEsquerda:=False
End Sub

' ============================================================
' MACRO: BarraLateral
' Aplica bordas escuras à seleção (sem borda direita)
' ============================================================
Sub BarraLateral()
    If TypeName(Selection) <> "Range" Then
        MsgBox "Selecione um intervalo de células antes de executar.", _
               vbExclamation, "Seleção inválida"
        Exit Sub
    End If

    AplicarBordas Selection, COR_BORDA_LATERAL, mostrarBordaEsquerda:=True
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
End Sub

' ============================================================
' MACRO: FundoVerde
' Aplica fundo verde + fonte branca centralizada à seleção
' ============================================================
Sub FundoVerde()
    If TypeName(Selection) <> "Range" Then
        MsgBox "Selecione um intervalo de células antes de executar.", _
               vbExclamation, "Seleção inválida"
        Exit Sub
    End If

    With Selection
        .Interior.Color = COR_HEADER_VERDE
        With .Font
            .Name = FONTE_PADRAO
            .Size = 10
            .Underline = xlUnderlineStyleNone
            .ThemeColor = xlThemeColorDark1
        End With
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
    End With
End Sub

' ============================================================
' MACRO: FormatoBD
' Transpõe tabela pivotada (colunas = períodos) para formato
' banco de dados (linhas = registros com coluna DIA e VALOR)
' ============================================================
Sub FormatoBD()
    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim wb As Workbook
    Dim colFim As Variant
    Dim colTituloFim As Long
    Dim colValorInicio As Long, colValorFim As Long
    Dim linhaFim As Long
    Dim novaLinha As Long
    Dim totalColunas As Long
    Dim i As Long
    Dim nomeDestinoBase As String, nomeDestino As String
    Dim sufixo As Long

    If Not TemPlanilhaAtiva() Then Exit Sub
    Set wb = ActiveWorkbook
    Set wsOrigem = ActiveSheet

    On Error GoTo ErroFormatoBD

    ' --- Quantidade de colunas de identificação ---
    colFim = Application.InputBox( _
        "Digite o número da última coluna de identificação (ex: se tiver 3 colunas de título, digite 3):", _
        "FormatoBD - Configuração", Type:=1)

    ' Cancelamento ou valor inválido
    If VarType(colFim) = vbBoolean Then
        GoTo Fim
    End If
    If CLng(colFim) < 1 Then
        MsgBox "Valor inválido.", vbInformation, "Cancelado"
        GoTo Fim
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    colTituloFim = CLng(colFim)
    colValorInicio = colTituloFim + 1
    colValorFim = wsOrigem.Cells(1, wsOrigem.Columns.count).End(xlToLeft).Column
    linhaFim = wsOrigem.Cells(wsOrigem.Rows.count, 1).End(xlUp).Row
    totalColunas = colValorFim - colValorInicio + 1

    ' --- Validação: estrutura mínima ---
    If colValorInicio > colValorFim Then
        MsgBox "Não foram encontradas colunas de valor após a coluna " & colTituloFim & "." & vbCrLf & _
               "Verifique o número informado.", vbExclamation, "Estrutura inválida"
        GoTo Fim
    End If

    If linhaFim < 2 Then
        MsgBox "Não há dados na planilha (apenas cabeçalho).", vbExclamation, "Sem dados"
        GoTo Fim
    End If

    ' --- Nome da aba destino (evita duplicatas) ---
    nomeDestinoBase = Left(wsOrigem.Name, 20) & "_BD"
    nomeDestino = nomeDestinoBase
    sufixo = 1
    Do While WorksheetExists(nomeDestino, wb)
        nomeDestino = nomeDestinoBase & "_" & sufixo
        sufixo = sufixo + 1
    Loop

    ' --- Cria aba destino ---
    wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.count)).Name = nomeDestino
    Set wsDestino = wb.Sheets(nomeDestino)

    ' --- Cabeçalho da aba destino ---
    wsDestino.Range( _
        wsDestino.Cells(1, 1), _
        wsDestino.Cells(1, colTituloFim)).Value = _
        wsOrigem.Range( _
            wsOrigem.Cells(1, 1), _
            wsOrigem.Cells(1, colTituloFim)).Value
    wsDestino.Cells(1, colTituloFim + 1).Value = "DIA"
    wsDestino.Cells(1, colTituloFim + 2).Value = "VALOR"

    novaLinha = 2

    ' --- Loop principal com barra de progresso ---
    For i = colValorInicio To colValorFim
        AtualizarProgresso i - colValorInicio + 1, totalColunas, _
            "Processando coluna '" & wsOrigem.Cells(1, i).Value & "'"

        ' Colunas de identificação (títulos)
        wsDestino.Range( _
            wsDestino.Cells(novaLinha, 1), _
            wsDestino.Cells(novaLinha + linhaFim - 2, colTituloFim)).Value = _
            wsOrigem.Range( _
                wsOrigem.Cells(2, 1), _
                wsOrigem.Cells(linhaFim, colTituloFim)).Value

        ' Coluna DIA (nome do período)
        wsDestino.Range( _
            wsDestino.Cells(novaLinha, colTituloFim + 1), _
            wsDestino.Cells(novaLinha + linhaFim - 2, colTituloFim + 1)).Value = _
            wsOrigem.Cells(1, i).Value

        ' Coluna VALOR
        wsDestino.Range( _
            wsDestino.Cells(novaLinha, colTituloFim + 2), _
            wsDestino.Cells(novaLinha + linhaFim - 2, colTituloFim + 2)).Value = _
            wsOrigem.Range( _
                wsOrigem.Cells(2, i), _
                wsOrigem.Cells(linhaFim, i)).Value

        novaLinha = novaLinha + linhaFim - 1
    Next i

    MsgBox "Concluído! Dados exportados para a aba '" & nomeDestino & "'." & vbCrLf & _
           totalColunas & " coluna(s) processada(s) | " & _
           (novaLinha - 2) & " linha(s) gerada(s).", _
           vbInformation, "FormatoBD"
    GoTo Fim

ErroFormatoBD:
    MsgBox "Erro em FormatoBD: " & Err.Description, vbCritical, "Erro"
Fim:
    LimparProgresso
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

' ============================================================
' MACRO: Copia_Imagem
' Copia a seleção como imagem (bitmap) para a área de transferência
' Atalho: Ctrl+I (substitui o atalho de Itálico enquanto o add-in estiver ativo)
' ============================================================
Sub Copia_Imagem()
Attribute Copia_Imagem.VB_ProcData.VB_Invoke_Func = "i\n14"
    If TypeName(Selection) = "Nothing" Then
        MsgBox "Nada selecionado para copiar.", vbExclamation, "Seleção vazia"
        Exit Sub
    End If

    On Error GoTo ErroCopiaImagem
    Selection.CopyPicture Appearance:=xlScreen, Format:=xlBitmap
    Application.StatusBar = "Imagem copiada para a área de transferência."
    Exit Sub

ErroCopiaImagem:
    MsgBox "Erro ao copiar imagem: " & Err.Description, vbCritical, "Erro"
End Sub

' ============================================================
' MACRO: SnapshotDados
' Grava uma "foto" (só valores) da aba ativa no fim da aba
' HISTORICO do mesmo arquivo, com data/hora ao lado
' ============================================================
Sub SnapshotDados()
    Dim wb As Workbook
    Dim wsOrigem As Worksheet, wsHist As Worksheet
    Dim dados As Range
    Dim ultimaLinha As Long

    If Not TemPlanilhaAtiva() Then Exit Sub
    Set wb = ActiveWorkbook
    Set wsOrigem = ActiveSheet

    If StrComp(wsOrigem.Name, "HISTORICO", vbTextCompare) = 0 Then
        MsgBox "Ative a aba que você quer fotografar (não a HISTORICO).", vbExclamation, "Snapshot"
        Exit Sub
    End If

    On Error GoTo ErroSnapshot
    Application.ScreenUpdating = False

    ' Cria a HISTORICO no fim do arquivo, se ainda não existir
    If Not WorksheetExists("HISTORICO", wb) Then
        wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.count)).Name = "HISTORICO"
        wsOrigem.Activate
    End If
    Set wsHist = wb.Worksheets("HISTORICO")

    ' Primeira linha livre
    If Application.WorksheetFunction.CountA(wsHist.Cells) = 0 Then
        ultimaLinha = 1
    Else
        ultimaLinha = wsHist.Cells(wsHist.Rows.count, 1).End(xlUp).Row + 1
    End If

    ' Copia só os valores, sem passar pela área de transferência
    Set dados = wsOrigem.UsedRange
    wsHist.Cells(ultimaLinha, 1).Resize(dados.Rows.count, dados.Columns.count).Value = dados.Value

    ' Carimbo de data/hora duas colunas à direita do bloco
    With wsHist.Cells(ultimaLinha, dados.Columns.count + 2)
        .Value = Now
        .NumberFormat = "dd/mm/yyyy hh:mm:ss"
    End With

    Application.StatusBar = "Snapshot de '" & wsOrigem.Name & "' salvo na HISTORICO (linha " & ultimaLinha & ")."
    GoTo Fim

ErroSnapshot:
    MsgBox "Erro em SnapshotDados: " & Err.Description, vbCritical, "Erro"
Fim:
    Application.ScreenUpdating = True
End Sub
