Attribute VB_Name = "sfComGerePlanilhas"
Sub RelatarProcessosArmazenados(Planilhas() As Excel.Worksheet)

    Dim arrCont() As Variant
    Dim lngRegistros() As Long, lngTotal As Long
    Dim strRelatorio As String
    Dim btCont As Byte
    
    ReDim lngRegistros(1 To UBound(Planilhas))
    
    ' Confere as planilhas
    For btCont = 1 To UBound(Planilhas)
        arrCont = ContaRegistrosPlanilha(Planilhas(btCont), 4)
        strRelatorio = strRelatorio & arrCont(1)
        lngRegistros(btCont) = arrCont(2)
        lngTotal = lngTotal + lngRegistros(btCont)
    Next btCont
    
    ' Aviso final
    strRelatorio = strRelatorio & vbCrLf & "Total de registros na mem�ria a exportar: " & lngTotal & vbCrLf
    
    ' Aviso se estiverem todas vazias
    If lngTotal = 0 Then
        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", todas as planilhas para inclus�o no Espaider est�o vazias.", vbInformation + vbOKOnly, "S�sifo - Planilhas de exporta��o vazias"
        Exit Sub
    Else
        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", segue o relat�rio das planilhas:" & vbCrLf & vbCrLf & strRelatorio, vbInformation + vbOKOnly, "S�sifo - Relat�rio de registros a exportar"
        Exit Sub
    End If

End Sub
Function ContaRegistrosPlanilha(plan As Excel.Worksheet, lngQtdLinhasCabecalho As Long) As Variant()
''
'' Na planilha passada como par�metro, conta quantos registros existem na primeira coluna (exceto as X primeiras linhas, quantidade passada como par�metro)
''
    Dim lngUltimaLinhaCont As Long, lngQuantidadeRegistros As Long
    Dim rngCont As Excel.Range
    Dim arrResposta(1 To 2) As Variant
    
    With plan
        lngUltimaLinhaCont = .UsedRange.Rows(.UsedRange.Rows.Count).Row
        If lngUltimaLinhaCont = lngQtdLinhasCabecalho Then
            lngQuantidadeRegistros = 0
        Else
            Set rngCont = .Range(.Cells(5, 1), .Cells(lngUltimaLinhaCont, 1))
            lngQuantidadeRegistros = Application.WorksheetFunction.CountA(rngCont)
        End If
        arrResposta(1) = "Registros em """ & .Name & """: " & lngQuantidadeRegistros & vbCrLf
        arrResposta(2) = lngQuantidadeRegistros
    End With
    
    ContaRegistrosPlanilha = arrResposta
    
End Function
