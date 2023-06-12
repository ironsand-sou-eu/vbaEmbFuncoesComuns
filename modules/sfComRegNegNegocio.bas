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
            
            If Left(strNumero, 1) <> "0" Or Left(strNumero, 2) = "03" Or Left(strNumero, 2) = "05" Then   ' Se começar com 03 ou 05, é eSaj (o Sísifo vai tratar como PJe, pois o TJ/BA descontinuará o eSaj e já migrou alguns processos)
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
'' Retorna o número do processo contido na primeira célula da range passada como parâmetro -- ou, se não for padrão CNJ, pergunta.
'' Em caso de erro, retorna a mensagem de erro.
''
    Dim strNumeroProcesso As String
    Dim intTentarDeNovo As Integer
    Dim rngCelula As Range
    
    Set rngCelula = rngRange(1, 1)
    strNumeroProcesso = Trim(rngCelula.Text)
    
    ' Se não houver célula no espaço enviado, ou se estiver vazia, ou se contiver algo em formato não CNJ, pergunta o número do processo.
    If rngCelula Is Nothing Or strNumeroProcesso = "" Or Not EhCNJ(strNumeroProcesso) Then
PerguntaNumero:
        strNumeroProcesso = InputBox(SisifoEmbasaFuncoes.DeterminarTratamento & ", informe o número do processo do Projudi a cadastrar no formato CNJ " & _
                "(""0000000-00.0000.0.00.0000""):", "Sísifo - Cadastrar processo")
            
AvisoNaoCNJ:
        If Not EhCNJ(strNumeroProcesso) Then
            intTentarDeNovo = MsgBox(SisifoEmbasaFuncoes.DeterminarTratamento & ", o número informado (" & strNumeroProcesso & ") não está no padrão do CNJ. " & _
                "Deseja tentar novamente com um número no padrão ""0000000-00.0000.0.00.0000""?", vbYesNo + vbCritical + vbDefaultButton1, _
                "Sísifo - Erro no cadastro")
                
            If intTentarDeNovo = vbYes Then
                GoTo PerguntaNumero
            Else
                PegarNumeroProcesso = "Número não é CNJ"
                Exit Function
            End If
        End If
    End If
    
    ' Se o conteúdo de strNumeroProcesso for um número de processo CNJ, aceita.
    PegaNumeroProcessoDeCelula = strNumeroProcesso
    
End Function

Function EhCNJ(strNumero As String) As Boolean
''
'' Confere se o número passado corresponde a um número de processo no padrão CNJ.
'' Só retorna VERDADEIRO se for padrão CNJ com pontos e traços.
'' O padrão é 0000000-00.0000.0.00.0000 - zeros significam qualquer número; hífens e pontos são hífens e pontos mesmo.
''

    Dim strCont As String
    Dim btCont As Byte
    Dim bolEhCNJ As Boolean
    
    bolEhCNJ = True
    
    ' Se não tiver 25 caracteres, não é CNJ
    If Len(strNumero) <> 25 Then bolEhCNJ = False
    
    For btCont = 1 To 25 Step 1 ' Itera caractere a caractere, verificando:
        strCont = Mid(strNumero, btCont, 1)
        
        Select Case btCont
        Case 8
            If strCont <> "-" Then bolEhCNJ = False ' Se o hífen está no lugar
        Case 11, 16, 18, 21
            If strCont <> "." Then bolEhCNJ = False ' Se os pontos estão no lugar
        Case Else
            If Not IsNumeric(strCont) Then bolEhCNJ = False ' Se os demais são números
        End Select
    Next btCont
    
    EhCNJ = bolEhCNJ
    
End Function

Function ValidaData(strStringAValidar As String) As String

    Dim strCont As String

    ' Validação da data
    strCont = Replace(strStringAValidar, " ", "")
    strCont = Replace(strCont, ":", "")
    strCont = Replace(strCont, "/", "")
    
    If IsNumeric(strCont) Then ' Se forem só números
        If Len(strCont) = 6 Then 'Ano com dois dígitos
            strCont = Left(strCont, 4) & "20" & Mid(strCont, 5)
            strCont = Format(strCont, "00/00/0000")
        ElseIf Len(strCont) = 8 Then 'Ano com quatro dígitos
            strCont = Format(strCont, "00/00/0000")
        ElseIf Len(strCont) = 10 Then 'Ano com dois dígitos e hora
            strCont = Left(strCont, 4) & "20" & Mid(strCont, 5)
            strCont = Format(strCont, "00/00/0000 00:00")
        ElseIf Len(strCont) = 12 Then 'Ano com quatro dígitos e hora
            strCont = Format(strCont, "00/00/0000 00:00")
        ElseIf Len(strCont) = 14 Then 'Ano com quatro dígitos e hora com segundos
            strCont = Left(strCont, 12)
            strCont = Format(strCont, "00/00/0000 00:00")
        End If
    Else ' Se não forem só números (não imagino como seria possível, mas...)
        strCont = Trim(strStringAValidar)
    End If
    
    If strCont <> "" And Not IsDate(strCont) Then
        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", o valor """ & strCont & """ não parece ser uma data. " & _
            "O programa rodará assim mesmo, mas é melhor corrigir, para evitar erros no Espaider.", vbOKOnly, "Sísifo - Data não reconhecida"
    End If
            
    ValidaData = strCont
    
End Function

Function ValidaNumeros(ChaveAscii As MSForms.ReturnInteger, Optional arrayStrOutrosCaractPermitidos As Variant) As Boolean
''
'' Faz uma validação front end, só permitindo números e, caso informados, os dois caracteres passados como parâmetros.
''
    Dim i As Variant
    
    Select Case ChaveAscii
    Case Asc("0") To Asc("9") 'Números são sempre permitidos
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

