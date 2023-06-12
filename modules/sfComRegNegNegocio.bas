Attribute VB_Name = "sfComRegNegNegocio"
Function InferirTribunalPeloNumero(plan As Excel.Worksheet, strNumero As String) As sfTribunal
    Dim strNumJustica As String, strNumTibunal As String
    Dim resposta As sfTribunal
    
    If EhCNJ(strNumero) Then
        strNumJustica = Mid(strNumero, 17, 1)
        strNumTibunal = Mid(strNumero, 19, 2)
        
        If strNumJustica = "8" And strNumTibunal = "05" Then
            resposta = sfTribunal.Tjba
        ElseIf strNumJustica = "5" And strNumTibunal = "05" Then
            resposta = sfTribunal.trt5
        Else
            resposta = sfTribunal.Erro
        End If
        
    Else
        resposta = sfTribunal.Erro
    End If
    
    InferirTribunalPeloNumero = resposta
End Function

Function InferirSistemaPeloNumero(plan As Excel.Worksheet, strNumero As String, tribunal As sfTribunal, instancia As sfInstancia) As sfSistema
    Dim strCont As String
    Dim resposta As sfSistema
    
    If Not EhCNJ(strNumero) Then
        resposta = sfSistema.Erro
        
    Else
        Select Case tribunal
        Case Tjba
            
            If Left(strNumero, 1) <> "0" Or Left(strNumero, 2) = "03" Or Left(strNumero, 2) = "05" Then   ' Se come�ar com 03 ou 05, � eSaj (o S�sifo vai tratar como PJe, pois o TJ/BA descontinuar� o eSaj e j� migrou alguns processos)
                If instancia = PrimeiroGrau Then
                    resposta = sfSistema.pje1g
                Else
                    resposta = sfSistema.pje2g
                End If
            Else
                resposta = sfSistema.projudi
            End If
            
        Case trt5
                If instancia = PrimeiroGrau Then
                    resposta = sfSistema.pje1g
                Else
                    resposta = sfSistema.pje2g
                End If
                
        Case Else
            resposta = sfSistema.Erro
        End Select
    End If
    
    InferirSistemaPeloNumero = resposta
End Function

Function PegaNumeroProcessoDeCelula(rngRange As Excel.Range) As String
''
'' Retorna o n�mero do processo contido na primeira c�lula da range passada como par�metro -- ou, se n�o for padr�o CNJ, pergunta.
'' Em caso de erro, retorna a mensagem de erro.
''
    Dim strNumeroProcesso As String
    Dim intTentarDeNovo As Integer
    Dim rngCelula As Range
    
    Set rngCelula = rngRange(1, 1)
    strNumeroProcesso = Trim(rngCelula.Text)
    
    ' Se n�o houver c�lula no espa�o enviado, ou se estiver vazia, ou se contiver algo em formato n�o CNJ, pergunta o n�mero do processo.
    If rngCelula Is Nothing Or strNumeroProcesso = "" Or Not EhCNJ(strNumeroProcesso) Then
PerguntaNumero:
        strNumeroProcesso = InputBox(SisifoEmbasaFuncoes.DeterminarTratamento & ", informe o n�mero do processo do Projudi a cadastrar no formato CNJ " & _
                "(""0000000-00.0000.0.00.0000""):", "S�sifo - Cadastrar processo")
            
AvisoNaoCNJ:
        If Not EhCNJ(strNumeroProcesso) Then
            intTentarDeNovo = MsgBox(SisifoEmbasaFuncoes.DeterminarTratamento & ", o n�mero informado (" & strNumeroProcesso & ") n�o est� no padr�o do CNJ. " & _
                "Deseja tentar novamente com um n�mero no padr�o ""0000000-00.0000.0.00.0000""?", vbYesNo + vbCritical + vbDefaultButton1, _
                "S�sifo - Erro no cadastro")
                
            If intTentarDeNovo = vbYes Then
                GoTo PerguntaNumero
            Else
                PegarNumeroProcesso = "N�mero n�o � CNJ"
                Exit Function
            End If
        End If
    End If
    
    ' Se o conte�do de strNumeroProcesso for um n�mero de processo CNJ, aceita.
    PegaNumeroProcessoDeCelula = strNumeroProcesso
    
End Function

Function EhCNJ(strNumero As String) As Boolean
''
'' Confere se o n�mero passado corresponde a um n�mero de processo no padr�o CNJ.
'' S� retorna VERDADEIRO se for padr�o CNJ com pontos e tra�os.
'' O padr�o � 0000000-00.0000.0.00.0000 - zeros significam qualquer n�mero; h�fens e pontos s�o h�fens e pontos mesmo.
''

    Dim strCont As String
    Dim btCont As Byte
    Dim bolEhCNJ As Boolean
    
    bolEhCNJ = True
    
    ' Se n�o tiver 25 caracteres, n�o � CNJ
    If Len(strNumero) <> 25 Then bolEhCNJ = False
    
    For btCont = 1 To 25 Step 1 ' Itera caractere a caractere, verificando:
        strCont = Mid(strNumero, btCont, 1)
        
        Select Case btCont
        Case 8
            If strCont <> "-" Then bolEhCNJ = False ' Se o h�fen est� no lugar
        Case 11, 16, 18, 21
            If strCont <> "." Then bolEhCNJ = False ' Se os pontos est�o no lugar
        Case Else
            If Not IsNumeric(strCont) Then bolEhCNJ = False ' Se os demais s�o n�meros
        End Select
    Next btCont
    
    EhCNJ = bolEhCNJ
    
End Function

Function ValidaData(strStringAValidar As String) As String

    Dim strCont As String

    ' Valida��o da data
    strCont = Replace(strStringAValidar, " ", "")
    strCont = Replace(strCont, ":", "")
    strCont = Replace(strCont, "/", "")
    
    If IsNumeric(strCont) Then ' Se forem s� n�meros
        If Len(strCont) = 6 Then 'Ano com dois d�gitos
            strCont = Left(strCont, 4) & "20" & Mid(strCont, 5)
            strCont = Format(strCont, "00/00/0000")
        ElseIf Len(strCont) = 8 Then 'Ano com quatro d�gitos
            strCont = Format(strCont, "00/00/0000")
        ElseIf Len(strCont) = 10 Then 'Ano com dois d�gitos e hora
            strCont = Left(strCont, 4) & "20" & Mid(strCont, 5)
            strCont = Format(strCont, "00/00/0000 00:00")
        ElseIf Len(strCont) = 12 Then 'Ano com quatro d�gitos e hora
            strCont = Format(strCont, "00/00/0000 00:00")
        ElseIf Len(strCont) = 14 Then 'Ano com quatro d�gitos e hora com segundos
            strCont = Left(strCont, 12)
            strCont = Format(strCont, "00/00/0000 00:00")
        End If
    Else ' Se n�o forem s� n�meros (n�o imagino como seria poss�vel, mas...)
        strCont = Trim(strStringAValidar)
    End If
    
    If strCont <> "" And Not IsDate(strCont) Then
        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", o valor """ & strCont & """ n�o parece ser uma data. " & _
            "O programa rodar� assim mesmo, mas � melhor corrigir, para evitar erros no Espaider.", vbOKOnly, "S�sifo - Data n�o reconhecida"
    End If
            
    ValidaData = strCont
    
End Function

Function ValidaNumeros(ChaveAscii As MSForms.ReturnInteger, Optional arrayStrOutrosCaractPermitidos As Variant) As Boolean
''
'' Faz uma valida��o front end, s� permitindo n�meros e, caso informados, os dois caracteres passados como par�metros.
''
    Dim i As Variant
    
    Select Case ChaveAscii
    Case Asc("0") To Asc("9") 'N�meros s�o sempre permitidos
        ValidaNumeros = True
    Case Else
        ValidaNumeros = False
    End Select
    
    If Not IsMissing(arrayStrOutrosCaractPermitidos) Then
        Select Case VarType(arrayStrOutrosCaractPermitidos)
        Case vbString
            If ChaveAscii = Asc(arrayStrOutrosCaractPermitidos) And arrayStrOutrosCaractPermitidos <> "" Then ValidaNumeros = True
        Case vbArray + vbVariant
            For Each i In arrayStrOutrosCaractPermitidos
                If i <> "" Then
                    If ChaveAscii = Asc(i) Then ValidaNumeros = True
                End If
            Next i
        End Select
    End If
End Function

Public Function PegarSequencialAndamento(numeroProcesso As String, dataAndamento As Date) As String
    Dim textoMarcador As String
    textoMarcador = Format(dataAndamento, "ddmmyyhhmm")
    PegarSequencialAndamento = PegarSequencial(numeroProcesso, textoMarcador)
End Function

Public Function PegarSequencialProvidencia(Providencia As Providencia, indice As Integer) As String
    Dim textoMarcador As String
    textoMarcador = Format(Providencia.DataFinal, "ddmmyy") & indice
    PegarSequencialProvidencia = PegarSequencial(Providencia.numeroProcesso, textoMarcador)
End Function

Public Function PegarSequencialPedido(numeroProcesso As String, codigoPedido As String) As String
    Dim textoMarcador As String
    textoMarcador = codigoPedido
    PegarSequencialPedido = PegarSequencial(numeroProcesso, textoMarcador)
End Function

Private Function PegarSequencial(numeroProcesso As String, textoAAdicionar As String) As String
    Dim sequencial As String
    
    sequencial = Replace(numeroProcesso, ".8.05.", "")
    sequencial = sequencial & textoAAdicionar
    sequencial = RetornarSomenteNumeros(sequencial)
    sequencial = RemoverZerosEsquerda(sequencial)
    PegarSequencial = sequencial
End Function

Private Function RetornarSomenteNumeros(textoEntrada As String) As String
    Dim sequencial() As String, caractere As String
    Dim i As Integer
    
    ReDim sequencial(Len(textoEntrada) - 1)
    For i = 1 To Len(textoEntrada)
        caractere = Mid(textoEntrada, i, 1)
        sequencial(i - 1) = IIf(IsNumeric(caractere), caractere, "")
    Next
    
    RetornarSomenteNumeros = Join(sequencial, "")
    
End Function

Private Function RemoverZerosEsquerda(textoEntrada As String) As String
    Dim texto As String
    Dim i As Integer
    Dim parar As Boolean
    
    texto = textoEntrada
    Do While Left(texto, 1) = "0"
        texto = Replace(texto, "0", "", 1, 1)
    Loop
    
    RemoverZerosEsquerda = texto
End Function

