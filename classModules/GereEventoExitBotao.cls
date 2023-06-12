VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GereEventoExitBotao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' Gerenciador de evento exclusivo para gerenciar o controle atual e o anterior

Private WithEvents moBotao As MSForms.CommandButton
Attribute moBotao.VB_VarHelpID = -1

Public Property Set Botao(ByVal oBotao As MSForms.CommandButton)
    Set moBotao = oBotao
End Property

Private Sub moBotao_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 9 Or KeyCode = 16 Then
        If oControleAtual Is Nothing Then Set oControleAnterior = oControleAtual
        Set oControleAtual = moBotao
    End If
End Sub

Private Sub moBotao_Click()
    If oControleAtual Is Nothing Then Set oControleAnterior = oControleAtual
    Set oControleAtual = moBotao
End Sub
