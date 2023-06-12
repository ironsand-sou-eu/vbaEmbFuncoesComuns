VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GereEventoExitCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' Gerenciador de evento exclusivo para gerenciar o controle atual e o anterior

Private WithEvents moCombo As MSForms.ComboBox
Attribute moCombo.VB_VarHelpID = -1

Public Property Set Combo(ByVal oCombo As MSForms.ComboBox)
    Set moCombo = oCombo
End Property

Private Sub moCombo_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 9 Or KeyCode = 16 Then
        If Not oControleAtual Is Nothing Then Set oControleAnterior = oControleAtual
        Set oControleAtual = moCombo
        
        If Not oControleAtual Is oControleAnterior Then
            ' Equivale ao evento Enter do controle clicado:
            AdicionarLinhaAjustarLegenda moCombo, 4
            
            ' Equivale ao evento Exit do controle anterior
            If Not oControleAnterior Is Nothing Then
                On Error Resume Next ' O controle anterior pode ser um botão
                oControleAnterior.Text = Trim(oControleAnterior.Text)
                On Error GoTo 0
                RetornarFormato oControleAnterior, 4
            End If
        End If
    End If
End Sub

Private Sub moCombo_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not oControleAtual Is Nothing Then Set oControleAnterior = oControleAtual
    Set oControleAtual = moCombo
    
    If Not oControleAtual Is oControleAnterior Then
        ' Equivale ao evento Enter do controle clicado:
        AdicionarLinhaAjustarLegenda moCombo, 4
        
        ' Equivale ao evento Exit do controle anterior
        If Not oControleAnterior Is Nothing Then
            If oControleAnterior <> False Then
                On Error Resume Next ' O controle anterior pode ser um botão
                oControleAnterior.Text = Trim(oControleAnterior.Text)
                On Error GoTo 0
                RetornarFormato oControleAnterior, 4
            End If
        End If
    End If
End Sub
