Attribute VB_Name = "sfComRegNegSistema"
#If Win64 Then
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal length As Long)

    Sub AoCarregarRibbon(ByVal plPlan As Worksheet, Ribbon As IRibbonUI)
        ' Guarda a referência ao objeto Ribbon
        Dim lngRibbon As LongPtr
        lngRibbon = ObjPtr(Ribbon)
        plPlan.Cells().Find(What:="Ponteiro do Ribbon", LookAt:=xlWhole).Offset(0, 1).Formula = "'" & lngRibbon
    End Sub
    
    Function RecuperarObjetoPorReferencia(arq As Excel.Workbook, planConfig As Excel.Worksheet) As IRibbonUI
        ' Recupera o objeto pela referência
        Dim lngRefObjeto As LongPtr, rbObjeto As IRibbonUI
        
        lngRefObjeto = planConfig.Cells().Find(What:="Ponteiro do Ribbon", LookAt:=xlWhole).Offset(0, 1).Text
        CopyMemory rbObjeto, lngRefObjeto, 6
        Set RecuperarObjetoPorReferencia = rbObjeto
    End Function

#Else
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal length As Long)
    
    Sub AoCarregarRibbon(ByVal plPlan As Worksheet, Ribbon As IRibbonUI)
        ' Guarda a referência ao objeto Ribbon
        Dim lngRibbon As Long
        lngRibbon = ObjPtr(Ribbon)
        plPlan.Cells().Find(What:="Ponteiro do Ribbon", LookAt:=xlWhole).Offset(0, 1).Formula = "'" & lngRibbon
    End Sub
    
    Function RecuperarObjetoPorReferencia(arq As Excel.Workbook, planConfig As Excel.Worksheet) As IRibbonUI
        ' Recupera o objeto pela referência
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
' Apresentar dados para as planilhas de configurações serem alterados pelo usuário
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
        
    MsgBox "Planilhas de configuração liberadas para edição. Tenha cuidado, e só realize alterações conforme as " & _
            "instruções fornecidas em cada planilha.", vbInformation + vbOKOnly, "Sísifo - Liberando alterações"
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
    'Se as planilhas de configuração estão abertas, avisa e oculta.
    If Arquivo.IsAddin = False Then
        MsgBox DeterminarTratamento & ", para salvar as alterações, as planilhas de configuração precisam ser ocultas novamente. ", _
                vbInformation + vbOKOnly, "Sísifo - Salvar"
        RestringirEdicaoRibbon Arquivo, planConfig
    End If
End Sub

Public Function ConferePlanilhasPerguntaSeFecha(arrPlans() As Excel.Worksheet, planConfig As Worksheet) As Boolean
' Se as planilhas não estiverem vazias, alerta usuário da existência de processos, permitindo cancelar
' o fechamento.
    Dim btCont As Byte
    Dim strNaoVazias As String
    
    For btCont = 1 To UBound(arrPlans) Step 1
        If ConfereSePlanilhaEstaVazia(arrPlans(btCont), 4) = False Then strNaoVazias = strNaoVazias & arrPlans(btCont).Name & " ,"
    Next btCont
    
    If strNaoVazias <> "" Then
        strNaoVazias = Left(strNaoVazias, Len(strNaoVazias) - 2)
    End If
    
    ' Se não estiverem todas vazias, avisa e permite interromper o fechamento.
    If strNaoVazias <> "" Then
        ConferePlanilhasPerguntaSeFecha = IIf(MsgBox(DeterminarTratamento & ", há registros ainda não exportados para o Espaider na planilha """ & _
            strNaoVazias & """. Se você sair, os " & _
            "processos que ainda não foram exportados serão guardados, e podem ser exportados " & _
            "posteriormente." & Chr(13) & "Deseja mesmo sair?", vbQuestion + vbYesNo, "Sísifo - Confirmar " & _
            "saída com processos pendentes") = vbYes, False, True)
        Exit Function
    End If

    ' Se as planilhas de configuração estiverem abertas para edição, esconde-as de novo.
    ConferePlanilhasPerguntaSeFecha = False
    RestringirEdicaoRibbon arrPlans(1).Parent, planConfig
End Function

Function ConfereSePlanilhaEstaVazia(plan As Excel.Worksheet, lngQtdRegistrosCabecalho As Long) As Boolean
' Retorna True se a planilha estiver vazia (isto é, apenas a quantidade de registros do cabeçalho)
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
'' Vai à planilha cfTratamentos, pega um adjetivo ou superlativo e um pronome/substantivo de tratamento.
''

    Dim rngCont As Range
    Dim intCont As Integer
    Dim strTratamento As String
    
    ' Define se será superlativo ou normal. intChance 1 = normal; 2 = superlativo
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
'' Reatribui o objeto InternetExplorer para a variável IE, perdida por causa da saída da intranet.
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
'' Espera a quantidade de segundos passada como parâmetro.
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
