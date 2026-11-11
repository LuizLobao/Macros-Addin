Attribute VB_Name = "modAlertasModelo"
' ============================================================
' modAlertasModelo
' ATENÇÃO: este módulo NÃO é executado a partir do add-in.
' Ele é um MODELO: o modInstalador copia este código para dentro
' do arquivo do usuário (como "modAlertas"). Por isso ele usa
' ThisWorkbook - depois de copiado, ThisWorkbook é o arquivo do usuário.
' Não renomeie este módulo (o instalador procura por este nome).
' ============================================================
Option Explicit
Private Const MAX_ITENS As Long = 60     ' máximo de itens listados por nota

Public Sub AtualizarAlertas()
    Dim ws As Worksheet, sh As Worksheet, r As Long
    Dim txtErr As String, txtCom As String
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("ALERTAS")
    On Error GoTo 0
    If ws Is Nothing Then MsgBox "Aba ALERTAS não encontrada.", vbExclamation: Exit Sub
    Application.ScreenUpdating = False
    With ws.Range("A7:C200")
        .ClearContents
        .ClearComments
        .Borders.LineStyle = xlNone
    End With
    r = 7
    For Each sh In ThisWorkbook.Worksheets
        If sh.Name <> ws.Name Then
            ws.Cells(r, 1).Value = sh.Name
            ws.Cells(r, 2).Value = ListarErros(sh, txtErr)
            ws.Cells(r, 3).Value = ListarComentarios(sh, txtCom)
            If Len(txtErr) > 0 Then CriarNota ws.Cells(r, 2), "Células com erro:" & vbLf & txtErr
            If Len(txtCom) > 0 Then CriarNota ws.Cells(r, 3), "Células com comentário:" & vbLf & txtCom
            r = r + 1
        End If
    Next sh
    With ws.Range("A6:C" & r).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = vbBlack
    End With
    ws.Range("B3").Value = Now
    Application.ScreenUpdating = True
    MsgBox "ALERTAS atualizada." & vbCrLf & _
           "Erros de fórmula: " & Application.Sum(ws.Range("B7:B" & r)) & vbCrLf & _
           "Comentários: " & Application.Sum(ws.Range("C7:C" & r)) & vbCrLf & vbCrLf & _
           "Passe o mouse nas células das colunas B e C para ver os detalhes.", vbInformation
End Sub

'--- Erros: devolve a quantidade e monta a lista em txt ---
Private Function ListarErros(sh As Worksheet, ByRef txt As String) As Long
    Dim rng As Range, c As Range, n As Long
    txt = ""
    On Error Resume Next
    Set rng = sh.UsedRange.SpecialCells(xlCellTypeFormulas, xlErrors)
    On Error GoTo 0
    If rng Is Nothing Then Exit Function
    ListarErros = rng.CountLarge
    For Each c In rng.Cells
        n = n + 1
        If n > MAX_ITENS Then Exit For
        txt = txt & c.Address(False, False) & "  " & NomeErro(c) & vbLf
    Next c
    If ListarErros > MAX_ITENS Then txt = txt & "... e mais " & (ListarErros - MAX_ITENS) & " célula(s)" & vbLf
End Function

Private Function NomeErro(c As Range) As String
    Select Case c.Value
        Case CVErr(xlErrDiv0):  NomeErro = "#DIV/0!"
        Case CVErr(xlErrNA):    NomeErro = "#N/D"
        Case CVErr(xlErrName):  NomeErro = "#NOME?"
        Case CVErr(xlErrNull):  NomeErro = "#NULO!"
        Case CVErr(xlErrNum):   NomeErro = "#NÚM!"
        Case CVErr(xlErrRef):   NomeErro = "#REF!"
        Case CVErr(xlErrValue): NomeErro = "#VALOR!"
        Case Else:              NomeErro = c.Text
    End Select
End Function

'--- Comentários + Anotações: devolve a quantidade e monta a lista em txt ---
Private Function ListarComentarios(sh As Worksheet, ByRef txt As String) As Long
    Dim shObj As Object, colCT As Object, ct As Object, cm As Comment
    Dim vistos As Object, n As Long, addr As String
    
    Set vistos = CreateObject("Scripting.Dictionary")
    txt = ""
    Set shObj = sh
    
    ' 1) Comentários (modernos, com respostas) - Excel 365 / 2021+
    On Error Resume Next
    Set colCT = shObj.CommentsThreaded
    On Error GoTo 0
    If Not colCT Is Nothing Then
        For Each ct In colCT
            addr = ct.Parent.Address(False, False)
            If Not vistos.Exists(addr) Then
                vistos.Add addr, True
                n = n + 1
                If n <= MAX_ITENS Then txt = txt & addr & " [Comentário]: " & TextoThreaded(ct) & vbLf
            End If
        Next ct
    End If
    
    ' 2) Anotações (notas amarelas)
    For Each cm In sh.Comments
        addr = cm.Parent.Address(False, False)
        If Not vistos.Exists(addr) Then
            vistos.Add addr, True
            n = n + 1
            If n <= MAX_ITENS Then txt = txt & addr & " [Anotação]: " & TextoNota(cm) & vbLf
        End If
    Next cm
    
    If n > MAX_ITENS Then txt = txt & "... e mais " & (n - MAX_ITENS) & " item(ns)" & vbLf
    ListarComentarios = n
End Function

Private Function TextoThreaded(ct As Object) As String
    Dim t As String
    t = ct.Text
    On Error Resume Next
    t = ct.Author.Name & ": " & t
    If ct.Replies.count > 0 Then t = t & "  (+" & ct.Replies.count & " resposta(s))"
    On Error GoTo 0
    TextoThreaded = Limpar(t)
End Function

Private Function TextoNota(cm As Comment) As String
    Dim t As String
    t = cm.Text
    If Len(cm.Author) > 0 And Left(t, Len(cm.Author) + 1) = cm.Author & ":" Then t = Mid(t, Len(cm.Author) + 2)
    TextoNota = Limpar(t)
End Function

Private Function Limpar(ByVal t As String) As String
    t = Trim(Replace(Replace(t, vbCr, " "), vbLf, " "))
    If Len(t) > 120 Then t = Left(t, 117) & "..."
    Limpar = t
End Function

'--- Cria a nota na célula (em partes, para textos longos) ---
Private Sub CriarNota(cel As Range, ByVal txt As String)
    Dim cm As Comment, i As Long
    If Right(txt, 1) = vbLf Then txt = Left(txt, Len(txt) - 1)
    Set cm = cel.AddComment("")
    For i = 1 To Len(txt) Step 250
        cm.Text Text:=Mid(txt, i, 250), Start:=i, Overwrite:=False
    Next i
    With cm.Shape
        .TextFrame.Characters.Font.Size = 9
        .TextFrame.AutoSize = True
    End With
End Sub
