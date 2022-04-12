Imports Microsoft.VisualBasic
Imports System.Data

Public Class PeriodoMapeamento
	Implements IDisposable

	Private RH89_ID_PERIODO_MAPEAMENTO As Integer
	Private RH36_ID_LOTACAO As String
	Private CA04_ID_USUARIO As Integer
	Private RH89_DH_CADASTRO As String
	Private RH89_NR_ANO_REFERENCIA As String

	Public Property idPeriodoMapeamento() As Integer
		Get
			Return RH89_ID_PERIODO_MAPEAMENTO
		End Get
		Set(ByVal Value As Integer)
			RH89_ID_PERIODO_MAPEAMENTO = Value
		End Set
	End Property
	Public Property idLotacao() As String
		Get
			Return RH36_ID_LOTACAO
		End Get
		Set(ByVal Value As String)
			RH36_ID_LOTACAO = Value
		End Set
	End Property
	Public Property idUsuario() As Integer
		Get
			Return CA04_ID_USUARIO
		End Get
		Set(ByVal Value As Integer)
			CA04_ID_USUARIO = Value
		End Set
	End Property
	Public Property DataHoraCadastro() As String
		Get
			Return RH89_DH_CADASTRO
		End Get
		Set(ByVal Value As String)
			RH89_DH_CADASTRO = Value
		End Set
	End Property
	Public Property Ano() As String
		Get
			Return RH89_NR_ANO_REFERENCIA
		End Get
		Set(ByVal Value As String)
			RH89_NR_ANO_REFERENCIA = Value
		End Set
	End Property

	Public Sub New(Optional ByVal idPeriodoMapeamento As Integer = 0)
		If idPeriodoMapeamento > 0 Then
			Obter(idPeriodoMapeamento)
		End If
	End Sub

	Public Sub Salvar()
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder

		strSQL.Append(" Select * ")
		strSQL.Append(" from RH89_PERIODO_MAPEAMENTO")
		strSQL.Append(" where RH89_ID_PERIODO_MAPEAMENTO = " & idPeriodoMapeamento)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("RH36_ID_LOTACAO") = ProBanco(RH36_ID_LOTACAO, eTipoValor.NUMERO_DECIMAL)
		dr("CA04_ID_USUARIO") = ProBanco(CA04_ID_USUARIO, eTipoValor.CHAVE)
		dr("RH89_DH_CADASTRO") = ProBanco(RH89_DH_CADASTRO, eTipoValor.DATA)
		dr("RH89_NR_ANO_REFERENCIA") = ProBanco(RH89_NR_ANO_REFERENCIA, eTipoValor.NUMERO_DECIMAL)

		cnn.SalvarDataTable(dr)

		dt.Dispose()
		dt = Nothing

		cnn.FecharBanco()
	End Sub

	Public Sub Obter(ByVal idPeriodoMapeamento As String)
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder

		strSQL.Append(" Select * ")
		strSQL.Append(" from RH89_PERIODO_MAPEAMENTO")
		strSQL.Append(" where RH89_ID_PERIODO_MAPEAMENTO = " & idPeriodoMapeamento)

		dt = cnn.AbrirDataTable(strSQL.ToString)

		If dt.Rows.Count > 0 Then
			dr = dt.Rows(0)

			RH89_ID_PERIODO_MAPEAMENTO = DoBanco(dr("RH89_ID_PERIODO_MAPEAMENTO"), eTipoValor.CHAVE)
			RH36_ID_LOTACAO = DoBanco(dr("RH36_ID_LOTACAO"), eTipoValor.NUMERO_DECIMAL)
			CA04_ID_USUARIO = DoBanco(dr("CA04_ID_USUARIO"), eTipoValor.CHAVE)
			RH89_DH_CADASTRO = DoBanco(dr("RH89_DH_CADASTRO"), eTipoValor.DATA)
			RH89_NR_ANO_REFERENCIA = DoBanco(dr("RH89_NR_ANO_REFERENCIA"), eTipoValor.NUMERO_DECIMAL)
		End If

		cnn.FecharBanco()
	End Sub

	Public Function Pesquisar(Optional ByVal Sort As String = "", Optional idPeriodoMapeamento As Integer = 0, Optional idLotacao As String = "", Optional idUsuario As Integer = 0, Optional DataHoraCadastro As String = "", Optional Ano As String = "") As DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder

		strSQL.Append(" Select * ")
		strSQL.Append(" from RH89_PERIODO_MAPEAMENTO")
		'strSQL.Append(" left join tabela On coluna1 = coluna2 ")
		strSQL.Append(" where RH89_ID_PERIODO_MAPEAMENTO Is Not null")

		If idPeriodoMapeamento > 0 Then
			strSQL.Append(" And RH89_ID_PERIODO_MAPEAMENTO = " & idPeriodoMapeamento)
		End If

		If IsNumeric(idLotacao.Replace(".", "")) Then
			strSQL.Append(" And RH36_ID_LOTACAO = " & idLotacao.Replace(".", "").Replace(",", "."))
		End If

		If idUsuario > 0 Then
			strSQL.Append(" And CA04_ID_USUARIO = " & idUsuario)
		End If

		If IsDate(DataHoraCadastro) Then
			strSQL.Append(" And RH89_DH_CADASTRO = Convert(DateTime, '" & DataHoraCadastro & "', 103)")
		End If

		If IsNumeric(Ano.Replace(".", "")) Then
			strSQL.Append(" and RH89_NR_ANO_REFERENCIA = " & Ano.Replace(".", "").Replace(",", "."))
		End If

		strSQL.Append(" Order By " & IIf(Sort = "", "RH89_ID_PERIODO_MAPEAMENTO", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

	Public Function ObterTabela() As DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder

		strSQL.Append(" select RH89_ID_PERIODO_MAPEAMENTO as CODIGO, RH36_ID_LOTACAO as DESCRICAO")
		strSQL.Append(" from RH89_PERIODO_MAPEAMENTO")
		strSQL.Append(" order by 2 ")

		dt = cnn.AbrirDataTable(strSQL.ToString)

		cnn.FecharBanco()

		Return dt
	End Function

	Public Function ObterUltimo() As Integer
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		Dim CodigoUltimo As Integer

		strSQL.Append(" select max(RH89_ID_PERIODO_MAPEAMENTO) from RH89_PERIODO_MAPEAMENTO")

		With cnn.AbrirDataTable(strSQL.ToString)
			If Not IsDBNull(.Rows(0)(0)) Then
				CodigoUltimo = .Rows(0)(0)
			Else
				CodigoUltimo = 0
			End If
		End With

		cnn.FecharBanco()

		Return CodigoUltimo

	End Function
	Public Function Excluir(ByVal idLotacao As String) As Integer
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		Dim LinhasAfetadas As Integer

		strSQL.Append(" delete ")
		strSQL.Append(" from RH89_PERIODO_MAPEAMENTO")
		strSQL.Append(" where RH36_ID_LOTACAO = " & idLotacao)
		strSQL.Append(" and RH89_NR_ANO_REFERENCIA = year(getdate())")

		LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

		cnn.FecharBanco()

		Return LinhasAfetadas
	End Function

#Region "IDisposable Support"
	Private disposedValue As Boolean ' Para detectar chamadas redundantes

	' IDisposable
	Protected Overridable Sub Dispose(disposing As Boolean)
		If Not disposedValue Then
			If disposing Then
				' TODO: descartar estado gerenciado (objetos gerenciados).
			End If

			' TODO: liberar recursos não gerenciados (objetos não gerenciados) e substituir um Finalize() abaixo.
			' TODO: definir campos grandes como nulos.
		End If
		disposedValue = True
	End Sub

	' TODO: substituir Finalize() somente se Dispose(disposing As Boolean) acima tiver o código para liberar recursos não gerenciados.
	'Protected Overrides Sub Finalize()
	'    ' Não altere este código. Coloque o código de limpeza em Dispose(disposing As Boolean) acima.
	'    Dispose(False)
	'    MyBase.Finalize()
	'End Sub

	' Código adicionado pelo Visual Basic para implementar corretamente o padrão descartável.
	Public Sub Dispose() Implements IDisposable.Dispose
		' Não altere este código. Coloque o código de limpeza em Dispose(disposing As Boolean) acima.
		Dispose(True)
		' TODO: remover marca de comentário da linha a seguir se Finalize() for substituído acima.
		GC.SuppressFinalize(Me)
	End Sub
#End Region

End Class

'******************************************************************************
'*                                 27/01/2020                                 *
'*                                                                            *
'*          ESTE CÓDIGO FOI GERADO PELO GERA CODIGO VERSÃO 4.0                *
'*    SUPORTE PARA ASP.NET 2.0, AJAX, SQL SERVER COM ENTERPRISE LIBRARY       *
'*                                                                            *
'*  O Gera-Codigo gera um MODELO de código Página, Interface, Classe e Css    *
'*  cabe a cada programador fazer as adaptações quando NECESSÁRIAS.           *
'*                                                                            *
'*  Esta ferramenta é TOTALMENTE GRATUITA, por favor, não remova os créditos  *
'*                                                                            *
'*  O autor não se responsabiliza por qualquer evento acontecido com o uso    *
'*  desta ferramenta ou do sistema que ela vier a gerar.                      *
'*                                                                            *
'*          Desenvolvido por Nírondes Anglada Casanovas Tavares               *
'*                  E-Mail/MSN: nirondes@hotmail.com                          *
'*                                                                            *
'******************************************************************************

