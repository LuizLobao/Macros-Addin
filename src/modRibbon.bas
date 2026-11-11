Attribute VB_Name = "modRibbon"
' ============================================================
' modRibbon
' Callbacks da guia "Macros Lobão" na Faixa de Opções.
' Só é usado se o arquivo ribbon/customUI14.xml for embutido
' no .xlam (veja o README). Sem o XML, este módulo é inofensivo.
' O callback do Sistema de Alertas (BtnCriarAlertas) fica no modInstalador.
' ============================================================
Option Explicit

Public Sub rbLayout1(control As IRibbonControl)
    Layout1
End Sub

Public Sub rbLayout2(control As IRibbonControl)
    Layout2
End Sub

Public Sub rbFundoVerde(control As IRibbonControl)
    FundoVerde
End Sub

Public Sub rbLinhasVerdes(control As IRibbonControl)
    LinhasVerdes
End Sub

Public Sub rbBarraLateral(control As IRibbonControl)
    BarraLateral
End Sub

Public Sub rbFormatoBD(control As IRibbonControl)
    FormatoBD
End Sub

Public Sub rbSnapshot(control As IRibbonControl)
    SnapshotDados
End Sub

Public Sub rbCopiaImagem(control As IRibbonControl)
    Copia_Imagem
End Sub

Public Sub rbAbasApoio(control As IRibbonControl)
    ExibirOcultarAbasComApoio
End Sub
