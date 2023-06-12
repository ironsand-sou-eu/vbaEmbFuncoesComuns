VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Providencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Esta Classe possiu muitas propriedades, representadas pelas seguintes variáveis:

Private dtDataAlerta As Date, dtDataFinal As Date
Private intPrazo As Integer, intSituacao As Integer, intAntecAlerta As Integer, intPeriodicidade As Integer, intTipoDataBaseAndamento As Integer
Private strNumeroProcesso As String, strNomeProv As String, strNucleo As String, strResponsavel As String, strObs As String, strAndamentoAGerar As String
Private strIdConversao As String, strIdAndamento As String
Private bolAlertar As Boolean, bolGerarAndamento As Boolean

Public Sub CarregarValoresPadrao(ByRef planConfiguracoes As Excel.Worksheet)
    ' Carrega os valores padrão
    intPrazo = planConfiguracoes.Cells().Find(What:="Prazo da Providência", LookAt:=xlWhole).Offset(0, 1).Formula
    intAntecAlerta = planConfiguracoes.Cells().Find(What:="Antecedência do alerta", LookAt:=xlWhole).Offset(0, 1).Formula
    intPeriodicidade = planConfiguracoes.Cells().Find(What:="Periodicidade do alerta", LookAt:=xlWhole).Offset(0, 1).Formula
    intSituacao = planConfiguracoes.Cells().Find(What:="Situação da Providência", LookAt:=xlWhole).Offset(0, 1).Formula
    'intTipoDataBaseAndamento = planConfiguracoes.Cells().Find(What:="Data Base Andamento", LookAt:=xlWhole).Offset(0, 1).Formula
    'strAndamento = planConfiguracoes.Cells().Find(What:="Andamento", LookAt:=xlWhole).Offset(0, 1).Formula
End Sub

'Propriedade NumeroProcesso:
Public Property Get numeroProcesso() As String
numeroProcesso = strNumeroProcesso
End Property
Public Property Let numeroProcesso(str As String)
strNumeroProcesso = str
End Property

'Propriedade Nome:
Public Property Get Nome() As String
Nome = strNomeProv
End Property
Public Property Let Nome(str As String)
strNomeProv = str
End Property

'Propriedade AndamentoAGerar:
Public Property Get AndamentoAGerar() As String
AndamentoAGerar = strAndamentoAGerar
End Property
Public Property Let AndamentoAGerar(str As String)
strAndamentoAGerar = str
End Property

'Propriedade GerarAndamento:
Public Property Get GerarAndamento() As Boolean
GerarAndamento = bolGerarAndamento
End Property
Public Property Let GerarAndamento(bol As Boolean)
bolGerarAndamento = bol
End Property

'Propriedade Núcleo:
Public Property Get Nucleo() As String
Nucleo = strNucleo
End Property
Public Property Let Nucleo(str As String)
strNucleo = str
End Property

'Propriedade Responsável:
Public Property Get Responsavel() As String
Responsavel = strResponsavel
End Property
Public Property Let Responsavel(str As String)
strResponsavel = str
End Property

'Propriedade Observação:
Public Property Get Observacao() As String
Observacao = strObs
End Property
Public Property Let Observacao(str As String)
strObs = str
End Property

'Propriedade IdConversao:
Public Property Get IdConversao() As String
IdConversao = strIdConversao
End Property
Public Property Let IdConversao(str As String)
strIdConversao = str
End Property

'Propriedade IdAndamento:
Public Property Get IdAndamento() As String
IdAndamento = strIdAndamento
End Property
Public Property Let IdAndamento(str As String)
strIdAndamento = str
End Property

'Propriedade Alertar:
Public Property Get Alertar() As Boolean
Alertar = bolAlertar
End Property
Public Property Let Alertar(bol As Boolean)
bolAlertar = bol
End Property

'Propriedade DataFinal:
Public Property Get DataFinal() As Date
DataFinal = dtDataFinal
End Property
Public Property Let DataFinal(dt As Date)
dtDataFinal = dt
End Property

'Propriedade DataAlerta:
Public Property Get DataAlerta() As Date
DataAlerta = dtDataAlerta
End Property
Public Property Let DataAlerta(dt As Date)
dtDataAlerta = dt
End Property

'Propriedade PrazoDias:
Public Property Get PrazoDias() As Integer
PrazoDias = intPrazo
End Property
Public Property Let PrazoDias(num As Integer)
intPrazo = num
End Property

'Propriedade Situacao:
Public Property Get Situacao() As Integer
Situacao = intSituacao
End Property
Public Property Let Situacao(num As Integer)
intSituacao = num
End Property

'Propriedade DiasAntecedenciaAlerta:
Public Property Get DiasAntecedenciaAlerta() As Integer
DiasAntecedenciaAlerta = intAntecAlerta
End Property
Public Property Let DiasAntecedenciaAlerta(num As Integer)
intAntecAlerta = num
End Property

'Propriedade PeriodicidadeAlerta:
Public Property Get PeriodicidadeAlerta() As Integer
PeriodicidadeAlerta = intPeriodicidade
End Property
Public Property Let PeriodicidadeAlerta(num As Integer)
intPeriodicidade = num
End Property

'Propriedade TipoDataBaseAndamento:
Public Property Get TipoDataBaseAndamento() As Integer
TipoDataBaseAndamento = intTipoDataBaseAndamento
End Property
Public Property Let TipoDataBaseAndamento(num As Integer)
intTipoDataBaseAndamento = num
End Property


