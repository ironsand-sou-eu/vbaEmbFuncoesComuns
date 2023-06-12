Attribute VB_Name = "sfComColDadProjudi"
Option Explicit

Function DescobrirPerfilLogadoProjudi(DocHTML As HTMLDocument) As String
''
'' Descobre o perfil do documento aberto e, conforme o caso, retorna "Parte" ou "Advogado"
''
    Dim frFrame As HTMLFrameElement
    Dim frForm As HTMLFormElement
    
    Set frFrame = DocHTML.getElementsByName("mainFrame")(0)
    
    On Error Resume Next
    Set frForm = frFrame.contentDocument.getElementsByName("formLogin")(0)
    On Error GoTo 0
    
    If Not frForm Is Nothing Then
        DescobrirPerfilLogadoProjudi = "Não logado"
    Else
        If InStr(1, frFrame.contentDocument.getElementById("Stm0p0i0eHR").href, "Parte") <> 0 Then 'É parte
            DescobrirPerfilLogadoProjudi = "Parte"
        ElseIf InStr(1, frFrame.contentDocument.getElementById("Stm0p0i0eHR").href, "Advogado") <> 0 Then ' É Advogado
            DescobrirPerfilLogadoProjudi = "Advogado"
        ElseIf InStr(1, frFrame.contentDocument.getElementById("Stm0p0i0eHR").href, "Representante") <> 0 Then ' É Representante
            DescobrirPerfilLogadoProjudi = "Representante"
        Else 'É outra coisa
            DescobrirPerfilLogadoProjudi = "Outro"
        End If
    End If
    
End Function

Function DescobrirPerfilLogadoProjudiChrome(oChrome As Selenium.ChromeDriver) As String
''
'' Descobre o perfil do documento aberto e, conforme o caso, retorna "Parte" ou "Advogado"
''
    Dim frForm As Selenium.WebElement
    Dim oFrame As Selenium.WebElement
    Dim strLink As String
    
    Select Case oChrome.Url
    Case SisifoEmbasaFuncoes.sfUrlProjudiHomeAdvogado
        If oChrome.FindElementsByTag("a").Count = 0 Then
            DescobrirPerfilLogadoProjudiChrome = "Representante"
        Else
            DescobrirPerfilLogadoProjudiChrome = "Advogado"
        End If
        
    Case SisifoEmbasaFuncoes.sfUrlProjudiBuscaAdvogado1g, SisifoEmbasaFuncoes.sfUrlProjudiBuscaAdvogado2g
        DescobrirPerfilLogadoProjudiChrome = "Advogado"
        
    Case SisifoEmbasaFuncoes.sfUrlProjudiBuscaParte1g, SisifoEmbasaFuncoes.sfUrlProjudiBuscaParte2g
        DescobrirPerfilLogadoProjudiChrome = "Representante"
        
    Case "https://projudi.tjba.jus.br/projudi/"
        oChrome.SwitchToFrame 1
        
        On Error Resume Next
        Set frForm = oChrome.FindElementByName("formLogin")
        Set oFrame = oChrome.FindElementByTag("iframe")
        On Error GoTo 0
        
        If Not frForm Is Nothing Then
            DescobrirPerfilLogadoProjudiChrome = "Não logado"
        ElseIf Not oFrame Is Nothing Then
            If InStr(1, LCase(oFrame.Attribute("src")), "representante") <> 0 Then
                DescobrirPerfilLogadoProjudiChrome = "Representante"
                oChrome.SwitchToFrame 0
                oChrome.ExecuteScript "submitForm(1);"
                'oChrome.SwitchToParentFrame
            Else: GoTo OutrosCasos
            End If
        Else
OutrosCasos:
            strLink = oChrome.FindElementById("Stm0p0i0eHR").Attribute("href")
            If InStr(1, strLink, "Parte") <> 0 Then 'É parte
                DescobrirPerfilLogadoProjudiChrome = "Parte"
            ElseIf InStr(1, strLink, "Advogado") <> 0 Then ' É Advogado
                DescobrirPerfilLogadoProjudiChrome = "Advogado"
            ElseIf InStr(1, strLink, "Representante") <> 0 Then ' É Representante
            Else 'É outra coisa
                DescobrirPerfilLogadoProjudiChrome = "Outro"
            End If
        End If
    End Select
End Function

Function CarregarPaginaBuscaProjudi(strPerfilLogado As String) As InternetExplorer
''
'' Abre nova janela do Internet Explorer na página de buscas, conforme perfil logado
''
    
    Dim IE As InternetExplorer
    Dim DocHTML As HTMLDocument
    Dim strCont As String
    
    If strPerfilLogado = "Outro" Then
        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", é necessário estar logado num perfil de parte, advogado ou representante. Faça login no Projudi de " & _
            "um desses perfis e rode a função novamente.", vbCritical + vbOKOnly, "Sísifo - Internet Explorer não encontrado"
        Set CarregarPaginaBuscaProjudi = Nothing
        Exit Function
    End If
    
    ' Carrega página de busca, conforme o perfil
    Set IE = New InternetExplorer
    IE.Visible = True
    IE.navigate IIf(strPerfilLogado = "Advogado", sfUrlProjudiBuscaAdvogado1g, sfUrlProjudiBuscaParte1g)
    Set IE = SisifoEmbasaFuncoes.RecuperarIE(IIf(strPerfilLogado = "Advogado", sfUrlProjudiBuscaAdvogado1g, sfUrlProjudiBuscaParte1g))
    
    ' Aguarda carregar
    Do
        DoEvents
    Loop Until IE.readyState = 4
    
    Do
        DoEvents
        strCont = IE.document.Url
    Loop Until strCont = IIf(strPerfilLogado = "Advogado", sfUrlProjudiBuscaAdvogado1g, sfUrlProjudiBuscaParte1g)
    
    Set DocHTML = IE.document
    'Set DocHTML = DocHTML.getElementsByName("mainFrame")(0).contentDocument.getElementsByName("userMainFrame")(0).contentDocument
    
    Set CarregarPaginaBuscaProjudi = IE
    
End Function

Function PegarLinkProcessoProjudi(ByVal strNumeroCNJ As String, ByVal strPerfilLogado As String, ByRef IE As InternetExplorer, ByRef DocHTML As HTMLDocument) As String
''
'' Retorna o link da página principal do processo strNumeroCNJ.
'' DEVO LIDAR COM O ERRO DE NÃO ESTAR LOGADO!!!!!!!
''

    Dim strContNumeroProcesso As String, strCont As String
    Dim frmProcessos As HTMLFormElement
    Dim intCont As Integer

'    ADICIONAR (NO LOCAL ADEQUADO) TRATAMENTO PARA:
'    Case "Não abriu por demora"
'        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", o processo não abriu por demora. Provavelmente, a conexão está muito lenta. Tente novamente daqui a pouco.", vbCritical + vbOKOnly, "Sísifo - Tempo de espera expirado"
'        GoTo FinalizaFechaIE
'    Case "Mais de um processo encontrado"
'        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", foi encontrado mais de um processo para o número " & rngProcesso.Formula & ". Isso é completamente inesperado! " & _
'            "Suplico que confira o número e tente novamente.", vbCritical + vbOKOnly, "Sísifo - Mais de um processo encontrado"
'        GoTo FinalizaFechaIE

    If DocHTML.Title = "Sistema CNJ - A sessão expirou" Then
        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", a sessão expirou. Faça login no Projudi NA MESMA JANELA EM QUE ESTÁ EXPIRADA, então " & _
        "clique OK e tente novamente.", vbCritical + vbOKOnly, "Sísifo - Sessão do Projudi expirada"
        PegarLinkProcessoProjudi = "Erro: sessão expirada"
        Exit Function
    End If
    
    DocHTML.getElementById("numeroProcesso").Value = strNumeroCNJ
    DocHTML.forms("busca").submit
    
    'SisifoEmbasaFuncoes.Esperar 1
    On Error GoTo Volta2
Volta2:
    Do
        DoEvents
    Loop Until IE.readyState = 4
    
    'Do
    '    DoEvents
    'Loop Until DocHTML.getElementsByTagName("body")(0).Children(0).Children(0).innerText = "Processos Obtidos Por Busca"
    
    Do
        strCont = IIf(strPerfilLogado = "Advogado", "form1", "formProcessos")
        Set frmProcessos = DocHTML.getElementById(strCont)
        'COLOCAR UM TIMEOUT AQUI (TRATAMENTO DO ERRO ESTÁ COMENTARIZADO ALI EM CIMA)
    Loop While frmProcessos Is Nothing

    intCont = frmProcessos.getElementsByTagName("a").length - 1
    For intCont = 0 To intCont Step 1
        If frmProcessos.getElementsByTagName("a")(intCont).innerText = strNumeroCNJ Then Exit For
    Next intCont
    On Error GoTo 0
    
    If intCont = frmProcessos.getElementsByTagName("a").length Then 'Correu todos os links e não achou
        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", lastimo informar o processo não foi encontrado. Verifique se o número está " & _
            "correto e tente novamente, ou talvez o processo não seja acessível para o usuário logado no Projudi (por exemplo, em segredo " & _
            "de justiça.", vbCritical + vbOKOnly, "Sísifo - Processo não encontrado"
        PegarLinkProcessoProjudi = "Erro: processo não encontrado"
    Else 'Achou
        PegarLinkProcessoProjudi = frmProcessos.getElementsByTagName("a")(intCont)
    End If
    
End Function

Sub ExpandirBotoesProcesso(ByRef IE As InternetExplorer, ByRef DocHTML As HTMLDocument, Optional ByVal intQuantidadeAExpandir As Integer)
''
'' Expande os "intQuantidadeAExpandir" primeiros botões de arquivos para download e informações de andamentos.
'' Se "intQuantidadeAExpandir" não tiver sido passada, abre tudo.
'' DEVO LIDAR COM O ERRO DE NÃO SER PASSADA UMA PÁGINA!!!!!!!
''
    Dim elCont As IHTMLElement
    Dim intCont As Integer, intContAbertos As Integer
    
    If intQuantidadeAExpandir <> 0 Then intContAbertos = 0
    
    For intCont = 0 To DocHTML.getElementsByTagName("img").length - 1
        Set elCont = DocHTML.getElementsByTagName("img")(intCont)
        If (InStr(1, elCont.outerHTML, "src=""/projudi/imagens/observacao.png""") <> 0) Or (InStr(1, elCont.outerHTML, "src=""/projudi/imagens/arquivos.png""") <> 0) Then
            elCont.parentElement.Click
            If intQuantidadeAExpandir <> 0 Then
                intContAbertos = intContAbertos + 1
                If intContAbertos > intQuantidadeAExpandir - 1 Then Exit For
            End If
        End If
    Next intCont
    
End Sub

Function ConverteDataProjudiParaDate(strData As String) As Date
''
'' Pega uma string no formato de data do projudi (por extenso) e converte em data.
''

    ' Retira início e final
    strData = Replace(strData, "(Agendada para ", "")
    strData = Replace(strData, "(Para ", "")
    strData = Replace(strData, "Incluído em pauta para ", "")
    strData = Replace(strData, " h)", "")
    strData = Replace(strData, " h )", "")
    strData = Replace(strData, " h", "")
    strData = Replace(strData, ")", "")
    
    ' Substitui " de " por barras
    strData = Replace(strData, " de ", "/")
    
    ' Substitui " às " por espaço
    strData = Replace(strData, " às ", " ")
    
    ' Substitui mês extenso por mês numérico
    If InStr(1, strData, "Janeiro") Then
        strData = Replace(strData, "Janeiro", "01")
    ElseIf InStr(1, strData, "Fevereiro") Then
        strData = Replace(strData, "Fevereiro", "02")
    ElseIf InStr(1, strData, "Março") Then
        strData = Replace(strData, "Março", "03")
    ElseIf InStr(1, strData, "Abril") Then
        strData = Replace(strData, "Abril", "04")
    ElseIf InStr(1, strData, "Maio") Then
        strData = Replace(strData, "Maio", "05")
    ElseIf InStr(1, strData, "Junho") Then
        strData = Replace(strData, "Junho", "06")
    ElseIf InStr(1, strData, "Julho") Then
        strData = Replace(strData, "Julho", "07")
    ElseIf InStr(1, strData, "Agosto") Then
        strData = Replace(strData, "Agosto", "08")
    ElseIf InStr(1, strData, "Setembro") Then
        strData = Replace(strData, "Setembro", "09")
    ElseIf InStr(1, strData, "Outubro") Then
        strData = Replace(strData, "Outubro", "10")
    ElseIf InStr(1, strData, "Novembro") Then
        strData = Replace(strData, "Novembro", "11")
    ElseIf InStr(1, strData, "Dezembro") Then
        strData = Replace(strData, "Dezembro", "12")
    End If
    
    ConverteDataProjudiParaDate = strData
    
End Function
