Attribute VB_Name = "Module1"
Option Explicit

' =====================================================================
' ArquivarViagemENovaViagem
'
' Salva o resumo da viagem atual (aba Config + totais de Orcamento e
' Despesas) como uma nova linha na aba Historico, depois limpa os
' campos de entrada de Config, Orcamento, Despesas, Compras e
' Planejamento para que a planilha fique pronta para a proxima viagem.
' As formulas de cada aba NAO sao apagadas, apenas os valores digitados.
' =====================================================================
Sub ArquivarViagemENovaViagem()
    Dim wsConfig As Worksheet, wsOrc As Worksheet, wsDesp As Worksheet
    Dim wsCompras As Worksheet, wsPlan As Worksheet, wsHist As Worksheet
    Dim histRow As Long, r As Long
    Dim resp As VbMsgBoxResult
    Dim nomeViagem As String

    On Error GoTo TratarErro

    Set wsConfig = ThisWorkbook.Sheets("Config")
    Set wsOrc = ThisWorkbook.Sheets("Orcamento")
    Set wsDesp = ThisWorkbook.Sheets("Despesas")
    Set wsCompras = ThisWorkbook.Sheets("Compras")
    Set wsPlan = ThisWorkbook.Sheets("Planejamento")
    Set wsHist = ThisWorkbook.Sheets("Historico")

    nomeViagem = Trim(wsConfig.Range("B5").Value)
    If nomeViagem = "" Then
        MsgBox "Preencha o nome da viagem (Config!B5) antes de arquivar.", vbExclamation, "Arquivar Viagem"
        Exit Sub
    End If

    resp = MsgBox("Arquivar a viagem """ & nomeViagem & """ na aba Historico e " & _
        "limpar Config / Orcamento / Despesas / Compras / Planejamento para a proxima viagem?" & _
        vbCrLf & vbCrLf & "Esta acao nao pode ser desfeita com Ctrl+Z depois de concluida.", _
        vbYesNo + vbQuestion, "Arquivar Viagem")
    If resp <> vbYes Then Exit Sub

    ' ---- 1) Encontrar a primeira linha vazia na aba Historico (a partir da linha 4) ----
    histRow = 4
    Do While wsHist.Cells(histRow, 1).Value <> "" And histRow < 5000
        histRow = histRow + 1
    Loop

    ' Copia a formatacao de uma linha ja pronta da tabela (linha 7 = primeira linha em branco
    ' do modelo original) para a nova linha, garantindo bordas/cores consistentes mesmo se a
    ' tabela pre-formatada (ate a linha 33) ja tiver sido totalmente preenchida.
    wsHist.Rows(7).Copy
    wsHist.Rows(histRow).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    ' ---- 2) Gravar o resumo da viagem atual na aba Historico ----
    With wsHist
        .Cells(histRow, 1).Value = nomeViagem                              ' Viagem
        .Cells(histRow, 2).Value = wsConfig.Range("B6").Value               ' Destino
        .Cells(histRow, 3).Value = wsConfig.Range("B7").Value               ' Data Inicio
        .Cells(histRow, 4).Value = wsConfig.Range("B8").Value               ' Data Fim
        .Cells(histRow, 3).NumberFormat = "dd/mm/yyyy"
        .Cells(histRow, 4).NumberFormat = "dd/mm/yyyy"
        .Cells(histRow, 5).Formula = "=IF(OR(C" & histRow & "="""",D" & histRow & "=""""),"""",D" & histRow & "-C" & histRow & "+1)"
        .Cells(histRow, 6).Value = wsConfig.Range("B10").Value              ' N. Pessoas
        .Cells(histRow, 7).Value = wsConfig.Range("B11").Value              ' Moeda Local
        .Cells(histRow, 8).Value = ThisWorkbook.Names("OrcTotal").RefersToRange.Value   ' Orcamento (Moeda Base)
        .Cells(histRow, 9).Value = ThisWorkbook.Names("DespTotal").RefersToRange.Value  ' Gasto Real (Moeda Base)
        .Cells(histRow, 8).NumberFormat = "#,##0.00 ""R$"""
        .Cells(histRow, 9).NumberFormat = "#,##0.00 ""R$"""
        .Cells(histRow, 10).Formula = "=IF(OR(H" & histRow & "="""",I" & histRow & "=""""),"""",H" & histRow & "-I" & histRow & ")"
        .Cells(histRow, 11).Formula = "=IF(OR(H" & histRow & "="""",H" & histRow & "=0),"""",I" & histRow & "/H" & histRow & ")"
        .Cells(histRow, 12).Formula = "=IF(OR(I" & histRow & "="""",E" & histRow & "="""",E" & histRow & "=0),"""",I" & histRow & "/E" & histRow & ")"
        .Cells(histRow, 13).Formula = "=IF(OR(L" & histRow & "="""",F" & histRow & "="""",F" & histRow & "=0),"""",L" & histRow & "/F" & histRow & ")"
        .Cells(histRow, 10).NumberFormat = "#,##0.00 ""R$"""
        .Cells(histRow, 11).NumberFormat = "0.0%"
        .Cells(histRow, 12).NumberFormat = "#,##0.00 ""R$"""
        .Cells(histRow, 13).NumberFormat = "#,##0.00 ""R$"""
        .Cells(histRow, 14).Value = ""   ' Nota (1-5) - preencha manualmente se quiser
        .Cells(histRow, 15).Value = ""   ' Notas - preencha manualmente se quiser
    End With

    ' ---- 3) Limpar Config (mantem Moeda Base) ----
    wsConfig.Range("B5:B8").ClearContents
    wsConfig.Range("B10:B11").ClearContents

    ' ---- 4) Limpar Orcamento (mantem a coluna de categorias) ----
    wsOrc.Range("B4:C13").ClearContents
    wsOrc.Range("G4:G13").ClearContents

    ' ---- 5) Limpar Despesas (mantem as formulas de Taxa de Cambio e Valor Base) ----
    For r = 4 To 38
        wsDesp.Range("A" & r & ":F" & r).ClearContents
        wsDesp.Range("I" & r & ":J" & r).ClearContents
    Next r

    ' ---- 6) Limpar Compras (mantem as formulas de Taxa de Cambio e Preco Base) ----
    For r = 4 To 24
        wsCompras.Range("A" & r & ":E" & r).ClearContents
        wsCompras.Range("H" & r & ":J" & r).ClearContents
    Next r

    ' ---- 7) Limpar Itinerario (mantem as formulas de Dia da Semana, Taxa e Custo Base) ----
    For r = 4 To 26
        wsPlan.Range("A" & r).ClearContents
        wsPlan.Range("C" & r & ":G" & r).ClearContents
        wsPlan.Range("J" & r & ":K" & r).ClearContents
    Next r

    ' ---- 8) Checklist: mantem os itens e so reseta o Status para "Pendente" ----
    For r = 31 To 43
        If Trim(wsPlan.Range("A" & r).Value) <> "" Then
            wsPlan.Range("D" & r).Value = "Pendente"
        End If
    Next r

    wsConfig.Activate
    wsConfig.Range("B5").Select

    MsgBox "Viagem """ & nomeViagem & """ arquivada na aba Historico (linha " & histRow & ")." & _
        vbCrLf & "Tudo pronto para planejar a proxima viagem!", vbInformation, "Arquivar Viagem"
    Exit Sub

TratarErro:
    MsgBox "Ocorreu um erro ao arquivar a viagem:" & vbCrLf & Err.Description, vbCritical, "Arquivar Viagem"
End Sub
