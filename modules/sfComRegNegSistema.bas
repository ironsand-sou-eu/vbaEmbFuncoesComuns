Attribute VB_Name = "sfComRegNegSistema"
#If Win64 Then
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal length As Long)

    Sub AoCarregarRibbon(ByVal plPlan As Worksheet, Ribbon As IRibbonUI)
        ' Guarda a refer�ncia ao objeto Ribbon
        Dim lngRibbon As LongPtr
        lngRibbon = ObjPtr(Ribbon)
        plPlan.Cells().Find(What:="Ponteiro do Ribbon", LookAt:=xlWhole).Offset(0, 1).Formula = "'" & lngRibbon
    End Sub
    
    Function RecuperarObjetoPorReferencia(arq As Excel.Workbook, planConfig As Excel.Worksheet) As IRibbonUI
        ' Recupera o objeto pela refer�ncia
        Dim lngRefObjeto As LongPtr, rbObjeto As IRibbonUI
        
        lngRefObjeto = planConfig.Cells().Find(What:="Ponteiro do Ribbon", LookAt:=xlWhole).Offset(0, 1).Text
        CopyMemory rbObjeto, lngRefObjeto, 6
        Set RecuperarObjetoPorReferencia = rbObjeto
    End Function

#Else
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal length As Long)
    
    Sub AoCarregarRibbon(ByVal plPlan As Worksheet, Ribbon As IRibbonUI)
        ' Guarda a refer�ncia ao objeto Ribbon
        Dim lngRibbon As Long
        lngRibbon = ObjPtr(Ribbon)
        plPlan.Cells().Find(What:="Ponteiro do Ribbon", LookAt:=xlWhole).Offset(0, 1).Formula = "'" & lngRibbon
    End Sub
    
    Function RecuperarObjetoPorReferencia(arq As Excel.Workbook, planConfig As Excel.Worksheet) As IRibbonUI
        ' Recupera o objeto pela refer�ncia
        Dim lngRefObjeto As Long, rbObjeto As IRibbonUI
        
        lngRefObjeto = planConfig.Cells().Find(What:="Ponteiro do Ribbon", LookAt:=xlWhole).Offset(0, 1).Text
        CopyMemory rbObjeto, lngRefObjeto, 4
        Set RecuperarObjetoPorReferencia = rbObjeto
    End Function
    
#End If

Function CaminhoDesktop() As String
    Dim caminho As String
    
    caminho = CreateObject("WScript.Shell").specialfolders("desktop")
    If Right(caminho, 1) <> "\" Then
        caminho = caminho & "\"
    End If
    CaminhoDesktop = caminho
    
End Function

Sub LiberarEdicao(arq As Workbook, planConfig As Excel.Worksheet)
' Apresentar dados para as planilhas de configura��es serem alterados pelo usu�rio
    Dim rib As IRibbonUI
    
    arq.IsAddin = False
    Set rib = RecuperarObjetoPorReferencia(arq, planConfig)
    
    Select Case Right(arq.CodeName, Len(arq.CodeName) - 3)
    Case "Citacoes"
        rib.InvalidateControl "btCitFechaConfig"
    Case "Intimacoes"
        rib.InvalidateControl "btIntFechaConfig"
    Case "Prazos"
        rib.InvalidateControl "btPrzFechaConfig"
    End Select
        
    MsgBox "Planilhas de configura��o liberadas para edi��o. Tenha cuidado, e s� realize altera��es conforme as " & _
            "instru��es fornecidas em cada planilha.", vbInformation + vbOKOnly, "S�sifo - Liberando altera��es"
End Sub

Sub RestringirEdicaoRibbon(arq As Workbook, planConfig As Worksheet, Optional ByVal Controle As IRibbonControl)
    arq.IsAddin = True
    Application.DisplayAlerts = False
    arq.Save
    'arq.SaveAs Filename:=arq.FullName, FileFormat:=xlOpenXMLAddIn
    Application.DisplayAlerts = True
    
    If Not Controle Is Nothing Then
        Select Case Right(arq.CodeName, Len(arq.CodeName) - 3)
        Case "Citacoes"
            FechaConfigVisivel arq, planConfig, Controle
        Case "Intimacoes"
            FechaConfigVisivel arq, planConfig, Controle
        Case "Prazos"
            FechaConfigVisivel arq, planConfig, Controle
        End Select
    End If
    
End Sub

Sub FechaConfigVisivel(arqSuplemento As Workbook, planConfig As Worksheet, Optional ByVal Controle As IRibbonControl, Optional ByRef returnedVal)
    Dim rib As IRibbonUI
    
    Set rib = RecuperarObjetoPorReferencia(arqSuplemento, planConfig)
    returnedVal = Not arqSuplemento.IsAddin
    rib.InvalidateControl Controle.ID
    
End Sub

Public Sub FechaConfiguracoesAoSalvar(ByRef Arquivo As Workbook, ByRef planConfig As Worksheet)
    'Se as planilhas de configura��o est�o abertas, avisa e oculta.
    If Arquivo.IsAddin = False Then
        MsgBox DeterminarTratamento & ", para salvar as altera��es, as planilhas de configura��o precisam ser ocultas novamente. ", _
                vbInformation + vbOKOnly, "S�sifo - Salvar"
        RestringirEdicaoRibbon Arquivo, planConfig
    End If
End Sub

Public Function ConferePlanilhasPerguntaSeFecha(arrPlans() As Excel.Worksheet, planConfig As Worksheet) As Boolean
' Se as planilhas n�o estiverem vazias, alerta usu�rio da exist�ncia de processos, permitindo cancelar
' o fechamento.
    Dim btCont As Byte
    Dim strNaoVazias As String
    
    For btCont = 1 To UBound(arrPlans) Step 1
        If ConfereSePlanilhaEstaVazia(arrPlans(btCont), 4) = False Then strNaoVazias = strNaoVazias & arrPlans(btCont).Name & " ,"
    Next btCont
    
    If strNaoVazias <> "" Then
        strNaoVazias = Left(strNaoVazias, Len(strNaoVazias) - 2)
    End If
    
    ' Se n�o estiverem todas vazias, avisa e permite interromper o fechamento.
    If strNaoVazias <> "" Then
        ConferePlanilhasPerguntaSeFecha = IIf(MsgBox(DeterminarTratamento & ", h� registros ainda n�o exportados para o Espaider na planilha """ & _
            strNaoVazias & """. Se voc� sair, os " & _
            "processos que ainda n�o foram exportados ser�o guardados, e podem ser exportados " & _
            "posteriormente." & Chr(13) & "Deseja mesmo sair?", vbQuestion + vbYesNo, "S�sifo - Confirmar " & _
            "sa�da com processos pendentes") = vbYes, False, True)
        Exit Function
    End If

    ' Se as planilhas de configura��o estiverem abertas para edi��o, esconde-as de novo.
    ConferePlanilhasPerguntaSeFecha = False
    RestringirEdicaoRibbon arrPlans(1).Parent, planConfig
End Function

Function ConfereSePlanilhaEstaVazia(plan As Excel.Worksheet, lngQtdRegistrosCabecalho As Long) As Boolean
' Retorna True se a planilha estiver vazia (isto �, apenas a quantidade de registros do cabe�alho)
    Dim lngUltimaLinhaPlan As Long
    
    lngUltimaLinhaPlan = plan.UsedRange.Rows(plan.UsedRange.Rows.Count).Row
    If lngUltimaLinhaPlan = lngQtdRegistrosCabecalho Then
        ConfereSePlanilhaEstaVazia = True
    Else
        ConfereSePlanilhaEstaVazia = False
    End If
End Function

Function DeterminarTratamento() As String
''
'' Vai � planilha cfTratamentos, pega um adjetivo ou superlativo e um pronome/substantivo de tratamento.
''

    Dim rngCont As Range
    Dim intCont As Integer
    Dim strTratamento As String
    
    ' Define se ser� superlativo ou normal. intChance 1 = normal; 2 = superlativo
    Randomize
    If CInt(100 * Rnd + 1) <= 60 Then intChance = 1 Else intChance = 2
    
    ' Conta os adjetivos e escolhe um aleatoriamente
    Set rngCont = cfTratamentos.Cells().Find(IIf(intChance = 1, "Adjetivos", "Superlativos"), LookAt:=xlWhole).Offset(1, 0)
    Set rngCont = Range(rngCont, rngCont.End(xlDown))
    Randomize
    intCont = CInt((rngCont.Cells().Count) * Rnd + 1)
    strTratamento = rngCont.Cells(intCont).Text
        
    ' Conta os substantivos e escolhe um aleatoriamente
    Set rngCont = cfTratamentos.Cells().Find("Substantivos", LookAt:=xlWhole).Offset(1, 0)
    Set rngCont = Range(rngCont, rngCont.End(xlDown))
    Randomize
    intCont = CInt((rngCont.Cells().Count) * Rnd + 1)
    strTratamento = strTratamento & " " & rngCont.Cells(intCont).Text
    
    DeterminarTratamento = strTratamento
    
End Function

Function RecuperarIE(strTrechoURLProcurada As String) As InternetExplorer
''
'' Reatribui o objeto InternetExplorer para a vari�vel IE, perdida por causa da sa�da da intranet.
''

    Dim Shell As Shell32.Shell
    Dim CadaIE As Variant
    Dim snInicioTimer As Single
    
    snInicioTimer = timer
IEVazio:
    'Do
    Set Shell = New Shell32.Shell
    For Each CadaIE In Shell.Windows
        If InStr(1, CadaIE.LocationURL, strTrechoURLProcurada) <> 0 Then Exit For
    Next CadaIE
    'Loop
    
    If timer >= snInicioTimer + 10 Then GoTo TempoEsgotado
    If CadaIE = Empty Then GoTo IEVazio

    Set RecuperarIE = CadaIE
    Exit Function
    
TempoEsgotado:
    Set RecuperarIE = Nothing
    
End Function

Sub Esperar(snSegundos As Single)
''
'' Espera a quantidade de segundos passada como par�metro.
''

    Dim snInicioTimer As Single
    
    snInicioTimer = timer
    Do
        DoEvents
    Loop Until timer >= snInicioTimer + snSegundos
    
End Sub

Function ControleExiste(oForm As UserForm, strNome As String) As Boolean

    On Error Resume Next
    ControleExiste = Not oForm.Controls(strNome) Is Nothing
    On Error GoTo 0
    
End Function
