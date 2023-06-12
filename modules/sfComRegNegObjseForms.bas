Attribute VB_Name = "sfComRegNegObjseForms"
Option Explicit

Public Function New_GereEventoExitBotao() As GereEventoExitBotao
    Set New_GereEventoExitBotao = New GereEventoExitBotao
End Function

Public Function New_GereEventoExitCombo() As GereEventoExitCombo
    Set New_GereEventoExitCombo = New GereEventoExitCombo
End Function

Public Function New_GereEventoExitCxTexto() As GereEventoExitCxTexto
    Set New_GereEventoExitCxTexto = New GereEventoExitCxTexto
End Function

Public Function New_Providencia() As Providencia
    Set New_Providencia = New Providencia
End Function

Public Sub AdicionarLinhaAjustarLegenda(ByRef Controle As MSForms.Control, sngVelocidade As Single)

    Dim lblLabel As MSForms.Label, lblLinha As MSForms.Label, ctrLabel As MSForms.Control
    Dim sngTimerInicio As Single, sngCont As Single, sngContCor As Single
    Dim sngTopDestino As Single, sngLeftDestino As Single, sngCorDestino As Single, btTamanhoFonteDestino As Byte
    Dim strCategoriaControle As String
    
    strCategoriaControle = Right(Controle.Name, Len(Controle.Name) - 3)
    Set lblLinha = Controle.Parent.Controls("Linha" & strCategoriaControle)
    Set ctrLabel = Controle.Parent.Controls("Label" & strCategoriaControle)
    
    ' Configura a linha azul
    lblLinha.Visible = True
    lblLinha.Width = 1
    
    If Controle.Text = "" Then
        ' Configurações do label (se estiver sem texto)
        Set lblLabel = ctrLabel
        sngTopDestino = ctrLabel.Top - 12
        sngLeftDestino = ctrLabel.Left - 6
        btTamanhoFonteDestino = 8
        sngCorDestino = 100
    End If
    
    ' Transição
    On Error Resume Next
    sngTimerInicio = timer
    While timer <= sngTimerInicio + (1 / sngVelocidade)
        sngCont = IIf(timer - sngTimerInicio = 0, 0.01, timer - sngTimerInicio)
        lblLinha.Width = Controle.Width * sngCont * sngVelocidade
        If Controle.Text = "" Then
            ctrLabel.Top = sngTopDestino + ctrLabel.Top * (0.01 / (sngCont * sngVelocidade / 2)) ' Expressão decrescente que tende a 0
            ctrLabel.Left = sngLeftDestino + ctrLabel.Left * (0.01 / (sngCont * sngVelocidade / 2)) ' Expressão decrescente que tende a 0
            lblLabel.Font.Size = btTamanhoFonteDestino + lblLabel.Font.Size * (0.01 / (sngCont * sngVelocidade / 2)) ' Expressão decrescente que tende a 0
            sngContCor = sngCorDestino * sngCont * sngVelocidade
            lblLabel.ForeColor = RGB(sngContCor, sngContCor, sngContCor)
        End If
        DoEvents
    Wend
    On Error GoTo 0
    
    ' Valores finais
    lblLinha.Width = Controle.Width
    If Controle.Text = "" Then
        ctrLabel.Top = sngTopDestino
        ctrLabel.Left = sngLeftDestino
        lblLabel.Font.Size = btTamanhoFonteDestino
        lblLabel.ForeColor = RGB(sngCorDestino, sngCorDestino, sngCorDestino)
    End If
    
End Sub

Public Sub AjustarLegendaSemTransicao(ByRef ctrLabel As MSForms.Control)

    Dim lblLabel As MSForms.Label
    Dim sngTopDestino As Single, sngLeftDestino As Single, sngCorDestino As Single, btTamanhoFonteDestino As Byte
    
    ' Configurações do label (se estiver sem texto)
    Set lblLabel = ctrLabel
    sngTopDestino = ctrLabel.Top - 12
    sngLeftDestino = ctrLabel.Left - 6
    btTamanhoFonteDestino = 8
    sngCorDestino = 100
    
    ctrLabel.Top = sngTopDestino
    ctrLabel.Left = sngLeftDestino
    lblLabel.Font.Size = btTamanhoFonteDestino
    lblLabel.ForeColor = RGB(sngCorDestino, sngCorDestino, sngCorDestino)
    
End Sub

Public Sub RetornarFormato(ByRef Controle As MSForms.Control, sngVelocidade As Single)

    Dim lblLabel As MSForms.Label, ctrLabel As MSForms.Control, lblLinha As MSForms.Label
    Dim sngTimerInicio As Single, sngCont As Single, sngContCor As Single
    Dim sngTopDestino As Single, sngLeftDestino As Single, sngCorDestino As Single, btTamanhoFonteDestino As Byte
    Dim strCategoriaControle As String
    
    strCategoriaControle = Right(Controle.Name, Len(Controle.Name) - 3)
    Set lblLinha = Controle.Parent.Controls("Linha" & strCategoriaControle)
    Set ctrLabel = Controle.Parent.Controls("Label" & strCategoriaControle)
    
    ' Configura a linha azul
    lblLinha.Visible = False
    
    ' Configurações
    If Controle.Text = "" Then
        Set lblLabel = ctrLabel
        sngTopDestino = ctrLabel.Top + 12
        sngLeftDestino = ctrLabel.Left + 6
        btTamanhoFonteDestino = 11.25
        sngCorDestino = 0
        
        ' Transição
        On Error Resume Next
        sngTimerInicio = timer
        While timer <= sngTimerInicio + (1 / sngVelocidade)
            sngCont = IIf(timer - sngTimerInicio = 0, 0.01, timer - sngTimerInicio)
            ctrLabel.Top = sngTopDestino * sngCont * sngVelocidade
            ctrLabel.Left = sngLeftDestino * sngCont * sngVelocidade
            lblLabel.Font.Size = btTamanhoFonteDestino * sngCont * sngVelocidade
            DoEvents
        Wend
        On Error GoTo 0
        
        ' Valores finais
        ctrLabel.Top = sngTopDestino
        ctrLabel.Left = sngLeftDestino
        lblLabel.Font.Size = btTamanhoFonteDestino
        lblLabel.ForeColor = RGB(sngCorDestino, sngCorDestino, sngCorDestino)
    End If
    
End Sub

Function SomarColecoes(colecao1 As Collection, colecao2 As Collection) As Collection
    
    Dim colecaoTotal As Collection
    Dim i As Variant
    
    Set colecaoTotal = New Collection
    
    If Not colecao1 Is Nothing Then
        For Each i In colecao1
            colecaoTotal.Add i
        Next i
    End If
    
    If Not colecao2 Is Nothing Then
        For Each i In colecao2
            colecaoTotal.Add i
        Next i
    End If
    
    Set SomarColecoes = colecaoTotal
    
End Function

Sub PreencherTextboxChromedriver(Textbox As Selenium.WebElement, ValorAPreencher As String)
    Do
        Textbox.Clear
        Textbox.SendKeys ValorAPreencher
    Loop Until Textbox.Value = ValorAPreencher
End Sub

Sub PreencherTextboxSimulandoDigitacaoChromedriver(Textbox As Selenium.WebElement, ValorAPreencher As String)
    Dim valor() As String
    Dim i As Integer
    Dim tempoInicial As Single, momentoHumano As Single
     
    valor = Split(StrConv(ValorAPreencher, vbUnicode), Chr$(0))
    ReDim Preserve valor(UBound(valor) - 1)
    
    Do
        Textbox.Clear
        For i = 0 To UBound(valor) Step 1
            tempoInicial = timer
            Textbox.SendKeys valor(i)
            Randomize
            momentoHumano = 0.7 - Rnd(0.5)
            Do
            Loop Until timer >= tempoInicial + momentoHumano
        Next i
    Loop Until Textbox.Value = ValorAPreencher
    tempoInicial = timer
    Randomize
    momentoHumano = 1 - Rnd(0.3)
    Do
    Loop Until timer >= tempoInicial + momentoHumano
End Sub

Public Function ValorExisteNaColecao(valor As Variant, colecao As Collection) As Boolean
    Dim i As Long
    For i = 1 To colecao.Count
        If colecao(i) = valor Then
            ValorExisteNaColecao = True
            Exit For
        End If
    Next i
End Function

Public Function ConverterRangeColunaParaVetorString(rangeColuna As Excel.Range) As String()
    Dim valoresRange As Variant
    Dim i As Long
    Dim resposta() As String
    
    valoresRange = rangeColuna.Value
    ReDim resposta(1 To UBound(valoresRange))
    
    For i = 1 To UBound(resposta)
        resposta(i) = CStr(valoresRange(i, 1))
    Next i
    
    ConverterRangeColunaParaVetorString = resposta
End Function
