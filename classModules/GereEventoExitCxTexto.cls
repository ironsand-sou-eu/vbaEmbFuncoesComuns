VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GereEventoExitCxTexto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' Gerenciador de evento exclusivo para gerenciar o controle atual e o anterior

Private WithEvents moCxTexto As MSForms.Textbox
Attribute moCxTexto.VB_VarHelpID = -1

Public Property Set CxTexto(ByVal oCxTexto As MSForms.Textbox)
    Set moCxTexto = oCxTexto
End Property

Private Sub moCxTexto_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 9 Or KeyCode = 16 Then
        If Not oControleAtual Is Nothing Then Set oControleAnterior = oControleAtual
        Set oControleAtual = moCxTexto
        
        If Not oControleAtual Is oControleAnterior Then
            ' Equivale ao evento Enter do controle clicado:
            AdicionarLinhaAjustarLegenda moCxTexto, 4
            
            ' Equivale ao evento Exit do controle anterior
            If Not oControleAnterior Is Nothing Then
                On Error Resume Next ' Porque o controle anterior pode ser um botão
                If InStr(1, LCase(oControleAnterior.Tag), "data") <> 0 Then ' Se for data
                    oControleAnterior.Text = SisifoEmbasaFuncoes.ValidaData(oControleAnterior.Text)
                Else
                    oControleAnterior.Text = Trim(oControleAnterior.Text)
                End If
                On Error GoTo 0
                If Not TypeOf oControleAnterior Is CommandButton Then RetornarFormato oControleAnterior, 4
            End If
        End If
    End If
End Sub

Private Sub moCxTexto_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not oControleAtual Is Nothing Then Set oControleAnterior = oControleAtual
    Set oControleAtual = moCxTexto
    
    If Not oControleAtual Is oControleAnterior Then
        ' Equivale ao evento Enter do controle clicado:
        AdicionarLinhaAjustarLegenda moCxTexto, 4
        
        ' Equivale ao evento Exit do controle anterior
        If Not oControleAnterior Is Nothing Then
            On Error Resume Next ' Porque o controle anterior pode ser um botão
            If InStr(1, LCase(oControleAnterior.Tag), "data") <> 0 Then ' Se for data
                oControleAnterior.Text = SisifoEmbasaFuncoes.ValidaData(oControleAnterior.Text)
            Else
                oControleAnterior.Text = Trim(oControleAnterior.Text)
            End If
            On Error GoTo 0
            If Not TypeOf oControleAnterior Is CommandButton Then RetornarFormato oControleAnterior, 4
        End If
    End If
End Sub
