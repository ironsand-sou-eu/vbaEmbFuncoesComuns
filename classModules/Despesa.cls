VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Despesa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strNumProc As String, strTipoDespesa As String, strComarca As String, strCodigoDAJE As String
Private dtVencimento As Date
Private curValorDAJE As Currency

'Propriedade NumeroProcesso:
Public Property Get numeroProcesso() As String
numeroProcesso = strNumProc
End Property
Public Property Let numeroProcesso(str As String)
strNumProc = str
End Property

'Propriedade ValorDAJE:
Public Property Get ValorDAJE() As Currency
ValorDAJE = curValorDAJE
End Property
Public Property Let ValorDAJE(valor As Currency)
curValorDAJE = valor
End Property

'Propriedade TipoDespesa:
Public Property Get TipoDespesa() As String
TipoDespesa = strTipoDespesa
End Property
Public Property Let TipoDespesa(str As String)
strTipoDespesa = str
End Property

'Propriedade Comarca:
Public Property Get Comarca() As String
Comarca = strComarca
End Property
Public Property Let Comarca(str As String)
strComarca = str
End Property

'Propriedade CodigoDAJE:
Public Property Get CodigoDAJE() As String
CodigoDAJE = strCodigoDAJE
End Property
Public Property Let CodigoDAJE(str As String)
strCodigoDAJE = str
End Property

'Propriedade Vencimento:
Public Property Get Vencimento() As Date
Vencimento = dtVencimento
End Property
Public Property Let Vencimento(dt As Date)
dtVencimento = dt
End Property

'Método ExportaLinhasEspaider:
Public Sub ExportaLinhasDespesasEspaider()
''
'' Exporta as linhas. Joga direto em CadDespesasDAJEs.
''
    
    Dim lnPrimeiraLinhaVazia As Long
    
    ' Confirmados os dados, confere se o processo já está na planilha.
    If Not ThisWorkbook.Sheets("Sisifo CadDespesasDAJEs").Cells().Find(strCodigoDAJE) Is Nothing Then
        'arq.Close False
        Do
        Loop Until MsgBox("Magnânimo mestre, o DAJE já existe na planilha! A inclusão foi cancelada. Descartamos os dados." & vbCrLf & _
        "Processo: " & strNumProc & vbCrLf & _
        "Tipo de despesa: " & strTipoDespesa & vbCrLf & _
        "DAJE: " & strCodigoDAJE & vbCrLf & _
        "Clique em 'Cancelar' e insira o próximo processo.", vbCritical + vbOKCancel, "Sísifo - DAJE repetido") = vbCancel
        Exit Sub
    End If
    
    With ThisWorkbook.Sheets("Sisifo CadDespesasDAJEs")
        lnPrimeiraLinhaVazia = .UsedRange.Rows.Count + 1
        .Cells(lnPrimeiraLinhaVazia, 1) = strNumProc
        .Cells(lnPrimeiraLinhaVazia, 2) = "Custas"
        .Cells(lnPrimeiraLinhaVazia, 3) = strTipoDespesa
        .Cells(lnPrimeiraLinhaVazia, 4).Formula = curValorDAJE
        .Cells(lnPrimeiraLinhaVazia, 5) = "'" & strCodigoDAJE
        .Cells(lnPrimeiraLinhaVazia, 6) = strComarca
        .Cells(lnPrimeiraLinhaVazia, 7) = 2
        .Cells(lnPrimeiraLinhaVazia, 8) = "Embasa S/A"
        .Cells(lnPrimeiraLinhaVazia, 9) = 901
        .Cells(lnPrimeiraLinhaVazia, 10) = 7
        .Cells(lnPrimeiraLinhaVazia, 11) = dtVencimento
        .Cells(lnPrimeiraLinhaVazia, 12) = curValorDAJE
    End With
    
    'Após salvar, confirma a inclusão e fecha a pasta de trabalho de rascunho.
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    'ThisWorkbook.SaveAs Filename:=ThisWorkbook.FullName, FileFormat:=xlOpenXMLAddIn
    Application.DisplayAlerts = True
    
End Sub






