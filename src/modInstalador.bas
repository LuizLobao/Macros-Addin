Attribute VB_Name = "modInstalador"
' ============================================================
' modInstalador
' Cria o "Sistema de Alertas" no arquivo ativo:
'   1) copia o código de modAlertasModelo para um módulo modAlertas no arquivo;
'   2) cria a aba ALERTAS com layout, formatação condicional e botão;
'   3) roda a primeira atualização;
'   4) pede para salvar como .xlsm, se necessário.
' Requer: "Confiar no acesso ao modelo de objeto do projeto do VBA".
' ============================================================
Option Explicit
Private Const MODELO  As String = "modAlertasModelo"   ' módulo-fonte no add-in
Private Const DESTINO As String = "modAlertas"         ' nome do módulo no arquivo novo
Private Const ABA     As String = "ALERTAS"

'=== Botão do sistema ===
Public Sub CriarSistemaAlertas()
    Dim wb As Workbook
    If ActiveWorkbook Is Nothing Then
        MsgBox "Abra ou crie um arquivo antes.", vbExclamation: Exit Sub
    End If
    Set wb = ActiveWorkbook
    If wb Is ThisWorkbook Or wb.IsAddin Then
        MsgBox "Ative o arquivo onde o sistema deve ser criado.", vbExclamation: Exit Sub
    End If
    If ExisteAba(wb, ABA) Then
        MsgBox "Este arquivo já possui a aba " & ABA & ".", vbInformation: Exit Sub
    End If
    If Not AcessoVBA(wb) Then
        MsgBox "Ative em: Arquivo > Opções > Central de Confiabilidade > " & _
               "Configurações da Central de Confiabilidade > Configurações de Macro > " & _
               """Confiar no acesso ao modelo de objeto do projeto do VBA"".", vbExclamation
        Exit Sub
    End If
    If ExisteModulo(wb, DESTINO) Then
        MsgBox "O módulo " & DESTINO & " já existe neste arquivo.", vbExclamation: Exit Sub
    End If
    Application.ScreenUpdating = False
    InjetarModulo wb
    MontarAba wb
    Application.ScreenUpdating = True
    ' Primeira atualização (roda a macro que agora está DENTRO do arquivo novo)
    Application.Run "'" & wb.Name & "'!" & DESTINO & ".AtualizarAlertas"
    SalvarComoXlsm wb
End Sub

'=== Callback para botão na Faixa de Opções ===
Public Sub BtnCriarAlertas(control As IRibbonControl)
    CriarSistemaAlertas
End Sub

'--- Copia o código do módulo modelo para um módulo novo no arquivo ---
Private Sub InjetarModulo(wb As Workbook)
    Dim src As Object, comp As Object, cod As String
    Set src = ThisWorkbook.VBProject.VBComponents(MODELO).CodeModule
    cod = src.Lines(1, src.CountOfLines)
    Set comp = wb.VBProject.VBComponents.Add(1)   ' 1 = módulo padrão
    comp.Name = DESTINO
    With comp.CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines  ' evita "Option Explicit" duplicado
        .AddFromString cod
    End With
End Sub

'--- Cria a aba ALERTAS com o layout esperado pela macro ---
Private Sub MontarAba(wb As Workbook)
    Dim ws As Worksheet, btn As Shape
    Set ws = wb.Worksheets.Add(Before:=wb.Worksheets(1))
    ws.Name = ABA
    With ws
        .Range("A1").Value = "ALERTAS - Erros de fórmula e comentários"
        .Range("A1").Font.Size = 14
        .Range("A1").Font.Bold = True
        .Range("A3").Value = "Última atualização:"
        .Range("A3").Font.Bold = True
        .Range("B3").NumberFormat = "dd/mm/yyyy hh:mm:ss"
        .Range("B3").HorizontalAlignment = xlLeft
        .Range("A4").Value = "Clique em 'Atualizar alertas'. Passe o mouse nas colunas B e C para ver os detalhes."
        .Range("A4").Font.Italic = True
        .Range("A4").Font.Color = RGB(89, 89, 89)
        .Range("A6:C6").Value = Array("Aba", "Erros de fórmula", "Comentários")
        With .Range("A6:C6")
            .Font.Bold = True
            .Font.Color = vbWhite
            .Interior.Color = RGB(31, 78, 121)
            .HorizontalAlignment = xlCenter
        End With
        .Columns("A").ColumnWidth = 32
        .Columns("B:C").ColumnWidth = 18
        .Range("B7:C200").HorizontalAlignment = xlCenter
        ' Destaque: erros em vermelho, comentários em amarelo
        With .Range("B7:B200").FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0")
            .Interior.Color = RGB(255, 199, 206): .Font.Color = RGB(156, 0, 6)
        End With
        With .Range("C7:C200").FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0")
            .Interior.Color = RGB(255, 235, 156): .Font.Color = RGB(156, 87, 0)
        End With
        ' Botão que chama a macro do próprio arquivo
        Set btn = .Shapes.AddShape(msoShapeRoundedRectangle, .Range("E2").Left, .Range("E2").Top, 150, 32)
        btn.Name = "btnAtualizarAlertas"
        btn.Fill.ForeColor.RGB = RGB(31, 78, 121)
        btn.Line.Visible = msoFalse
        With btn.TextFrame2
            .TextRange.Text = "Atualizar alertas"
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = vbWhite
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .VerticalAnchor = msoAnchorMiddle
        End With
        btn.OnAction = "'" & wb.Name & "'!" & DESTINO & ".AtualizarAlertas"
    End With
    ws.Activate
    ActiveWindow.DisplayGridlines = False
End Sub

'--- Garante formato com macros ---
Private Sub SalvarComoXlsm(wb As Workbook)
    Dim f As Variant, nome As String
    Select Case wb.FileFormat
        Case xlOpenXMLWorkbookMacroEnabled, xlExcel12, xlExcel8   ' xlsm, xlsb, xls
            If Len(wb.Path) > 0 Then Exit Sub
    End Select
    nome = wb.Name
    If InStr(nome, ".") > 0 Then nome = Left(nome, InStrRev(nome, ".") - 1)
    MsgBox "Para manter as macros, salve o arquivo como .xlsm.", vbInformation
    f = Application.GetSaveAsFilename(InitialFileName:=nome & ".xlsm", _
        FileFilter:="Pasta de Trabalho Habilitada para Macro (*.xlsm), *.xlsm")
    If f = False Then
        MsgBox "Atenção: se salvar como .xlsx, as macros serão perdidas.", vbExclamation
        Exit Sub
    End If
    wb.SaveAs Filename:=f, FileFormat:=xlOpenXMLWorkbookMacroEnabled
End Sub

'--- Auxiliares ---
Private Function ExisteAba(wb As Workbook, nome As String) As Boolean
    Dim s As Object
    On Error Resume Next
    Set s = wb.Sheets(nome)
    ExisteAba = Not s Is Nothing
End Function

Private Function ExisteModulo(wb As Workbook, nome As String) As Boolean
    Dim c As Object
    On Error Resume Next
    Set c = wb.VBProject.VBComponents(nome)
    ExisteModulo = Not c Is Nothing
End Function

Private Function AcessoVBA(wb As Workbook) As Boolean
    Dim n As Long
    On Error Resume Next
    n = wb.VBProject.VBComponents.count
    AcessoVBA = (Err.Number = 0)
End Function
