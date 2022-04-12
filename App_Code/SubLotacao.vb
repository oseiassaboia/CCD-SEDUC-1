Imports Microsoft.VisualBasic
Imports System.Data

Public Class SubLotacao
	Implements IDisposable
	Private RH40_ID_SUBLOTACAO As Integer
	Private CA04_ID_USUARIO As Integer
	Private RH40_NM_SUBLOTACAO As String
	Private RH40_DH_CADASTRO As String

	Public Property IdSubLotacao() As Integer
		Get
			Return RH40_ID_SUBLOTACAO
		End Get
		Set(ByVal Value As Integer)
			RH40_ID_SUBLOTACAO = Value
		End Set
	End Property
	Public Property IdUsuario() As Integer
		Get
			Return CA04_ID_USUARIO
		End Get
		Set(ByVal Value As Integer)
			CA04_ID_USUARIO = Value
		End Set
	End Property
	Public Property Descricao() As String
		Get
			Return RH40_NM_SUBLOTACAO
		End Get
		Set(ByVal Value As String)
			RH40_NM_SUBLOTACAO = Value
		End Set
	End Property
	Public Property DataHoraCadastro() As String
		Get
			Return RH40_DH_CADASTRO
		End Get
		Set(ByVal Value As String)
			RH40_DH_CADASTRO = Value
		End Set
	End Property

	Public Sub New(Optional ByVal IdSubLotacao As Integer = 0)
		If IdSubLotacao > 0 Then
			Obter(IdSubLotacao)
		End If
	End Sub

	Public Sub Salvar()
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder

		strSQL.Append(" select * ")
		strSQL.Append(" from RH40_SUBLOTACAO")
		strSQL.Append(" where RH40_ID_SUBLOTACAO = " & IdSubLotacao)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("CA04_ID_USUARIO") = ProBanco(CA04_ID_USUARIO, eTipoValor.CHAVE)
		dr("RH40_NM_SUBLOTACAO") = ProBanco(RH40_NM_SUBLOTACAO, eTipoValor.TEXTO)
		dr("RH40_DH_CADASTRO") = ProBanco(RH40_DH_CADASTRO, eTipoValor.DATA)

		cnn.SalvarDataTable(dr)

		dt.Dispose()
		dt = Nothing

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Sub Obter(ByVal IdSubLotacao As Integer)
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder

		strSQL.Append(" select * ")
		strSQL.Append(" from RH40_SUBLOTACAO")
		strSQL.Append(" where RH40_ID_SUBLOTACAO = " & IdSubLotacao)

		dt = cnn.AbrirDataTable(strSQL.ToString)

		If dt.Rows.Count > 0 Then
			dr = dt.Rows(0)

			RH40_ID_SUBLOTACAO = DoBanco(dr("RH40_ID_SUBLOTACAO"), eTipoValor.CHAVE)
			CA04_ID_USUARIO = DoBanco(dr("CA04_ID_USUARIO"), eTipoValor.CHAVE)
			RH40_NM_SUBLOTACAO = DoBanco(dr("RH40_NM_SUBLOTACAO"), eTipoValor.TEXTO)
			RH40_DH_CADASTRO = DoBanco(dr("RH40_DH_CADASTRO"), eTipoValor.DATA)
		End If

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Function Pesquisar(Optional ByVal Sort As String = "", Optional IdSubLotacao As Integer = 0, Optional IdUsuario As Integer = 0, Optional Descricao As String = "", Optional DataHoraCadastro As String = "") As DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder

		strSQL.Append(" select * ")
		strSQL.Append(" from RH40_SUBLOTACAO")
		'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
		strSQL.Append(" where RH40_ID_SUBLOTACAO is not null")

		If IdSubLotacao > 0 Then
			strSQL.Append(" and RH40_ID_SUBLOTACAO = " & IdSubLotacao)
		End If

		If IdUsuario > 0 Then
			strSQL.Append(" and CA04_ID_USUARIO = " & IdUsuario)
		End If

		If Descricao <> "" Then
			strSQL.Append(" and upper(RH40_NM_SUBLOTACAO) like '%" & Descricao.ToUpper & "%'")
		End If

		If IsDate(DataHoraCadastro) Then
			strSQL.Append(" and RH40_DH_CADASTRO = Convert(DateTime, '" & DataHoraCadastro & "', 103)")
		End If

		strSQL.Append(" Order By " & IIf(Sort = "", "RH40_ID_SUBLOTACAO", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

	Public Function ObterTabela() As DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder

		strSQL.Append(" select RH40_ID_SUBLOTACAO as CODIGO, RH40_NM_SUBLOTACAO as DESCRICAO")
		strSQL.Append(" from RH40_SUBLOTACAO")
		strSQL.Append(" order by 2 ")

		dt = cnn.AbrirDataTable(strSQL.ToString)

		cnn.FecharBanco()
		cnn = Nothing

		Return dt
	End Function

	Public Function ObterUltimo() As Integer
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		Dim CodigoUltimo As Integer

		strSQL.Append(" select max(RH40_ID_SUBLOTACAO) from RH40_SUBLOTACAO")

		With cnn.AbrirDataTable(strSQL.ToString)
			If Not IsDBNull(.Rows(0)(0)) Then
				CodigoUltimo = .Rows(0)(0)
			Else
				CodigoUltimo = 0
			End If
		End With

		cnn.FecharBanco()
		cnn = Nothing

		Return CodigoUltimo

	End Function
	Public Function Excluir(ByVal IdSubLotacao As String) As Integer
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		Dim LinhasAfetadas As Integer

		strSQL.Append(" delete ")
		strSQL.Append(" from RH40_SUBLOTACAO")
		strSQL.Append(" where RH40_ID_SUBLOTACAO = " & IdSubLotacao)

		LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

		cnn.FecharBanco()
		cnn = Nothing

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
'*                                 06/05/2019                                 *
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

