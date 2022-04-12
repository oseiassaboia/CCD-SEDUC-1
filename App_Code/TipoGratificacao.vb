Imports Microsoft.VisualBasic
Imports System.Data

Public Class TipoGratificacao
	Implements IDisposable

	Private RH90_ID_TIPO_GRATIFICACAO As Integer
	Private RH90_NM_TIPO_GRATIFICACAO As String
	Private disposedValue As Boolean

	Public Property IdTipoGratificacao() As Integer
		Get
			Return RH90_ID_TIPO_GRATIFICACAO
		End Get
		Set(ByVal Value As Integer)
			RH90_ID_TIPO_GRATIFICACAO = Value
		End Set
	End Property
	Public Property Descricao() As String
		Get
			Return RH90_NM_TIPO_GRATIFICACAO
		End Get
		Set(ByVal Value As String)
			RH90_NM_TIPO_GRATIFICACAO = Value
		End Set
	End Property

	Public Sub New(Optional ByVal IdTipoGratificacao As Integer = 0)
		If IdTipoGratificacao > 0 Then
			Obter(IdTipoGratificacao)
		End If
	End Sub

	Public Sub Salvar()
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder

		strSQL.Append("Then select * ")
		strSQL.Append(" from RH99_TIPO_GRATIFICACAO")
		strSQL.Append(" where RH90_ID_TIPO_GRATIFICACAO = " & IdTipoGratificacao)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("RH90_NM_TIPO_GRATIFICACAO") = ProBanco(RH90_NM_TIPO_GRATIFICACAO, eTipoValor.TEXTO)

		cnn.SalvarDataTable(dr)

		dt.Dispose()
		dt = Nothing

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Sub Obter(ByVal IdTipoGratificacao As String)
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder

		strSQL.Append(" select * ")
		strSQL.Append(" from RH99_TIPO_GRATIFICACAO")
		strSQL.Append(" where RH90_ID_TIPO_GRATIFICACAO = " & IdTipoGratificacao)

		dt = cnn.AbrirDataTable(strSQL.ToString)

		If dt.Rows.Count > 0 Then
			dr = dt.Rows(0)

			RH90_ID_TIPO_GRATIFICACAO = DoBanco(dr("RH90_ID_TIPO_GRATIFICACAO"), eTipoValor.CHAVE)
			RH90_NM_TIPO_GRATIFICACAO = DoBanco(dr("RH90_NM_TIPO_GRATIFICACAO"), eTipoValor.TEXTO)
		End If

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Function Pesquisar(Optional ByVal Sort As String = "", Optional IdTipoGratificacao As Integer = 0, Optional Descricao As String = "") As DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder

		strSQL.Append(" select * ")
		strSQL.Append(" from RH99_TIPO_GRATIFICACAO")
		'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
		strSQL.Append(" where RH90_ID_TIPO_GRATIFICACAO is not null")

		If IdTipoGratificacao > 0 Then
			strSQL.Append(" and RH90_ID_TIPO_GRATIFICACAO = " & IdTipoGratificacao)
		End If

		If Descricao <> "" Then
			strSQL.Append(" and upper(RH90_NM_TIPO_GRATIFICACAO) like '%" & Descricao.ToUpper & "%'")
		End If

		strSQL.Append(" Order By " & IIf(Sort = "", "RH90_ID_TIPO_GRATIFICACAO", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

	Public Function ObterTabela() As DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder

		strSQL.Append(" select RH90_ID_TIPO_GRATIFICACAO as CODIGO, RH90_NM_TIPO_GRATIFICACAO as DESCRICAO")
		strSQL.Append(" from RH90_TIPO_GRATIFICACAO")
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

		strSQL.Append(" select max(RH90_ID_TIPO_GRATIFICACAO) from RH99_TIPO_GRATIFICACAO")

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
	Public Function Excluir(ByVal IdTipoGratificacao As String) As Integer
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		Dim LinhasAfetadas As Integer

		strSQL.Append(" delete ")
		strSQL.Append(" from RH99_TIPO_GRATIFICACAO")
		strSQL.Append(" where RH90_ID_TIPO_GRATIFICACAO = " & IdTipoGratificacao)

		LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

		cnn.FecharBanco()
		cnn = Nothing

		Return LinhasAfetadas
	End Function

	Protected Overridable Sub Dispose(disposing As Boolean)
		If Not disposedValue Then
			If disposing Then
				' Tarefa pendente: descartar o estado gerenciado (objetos gerenciados)
			End If

			' Tarefa pendente: liberar recursos não gerenciados (objetos não gerenciados) e substituir o finalizador
			' Tarefa pendente: definir campos grandes como nulos
			disposedValue = True
		End If
	End Sub

	' ' Tarefa pendente: substituir o finalizador somente se 'Dispose(disposing As Boolean)' tiver o código para liberar recursos não gerenciados
	' Protected Overrides Sub Finalize()
	'     ' Não altere este código. Coloque o código de limpeza no método 'Dispose(disposing As Boolean)'
	'     Dispose(disposing:=False)
	'     MyBase.Finalize()
	' End Sub

	Public Sub Dispose() Implements IDisposable.Dispose
		' Não altere este código. Coloque o código de limpeza no método 'Dispose(disposing As Boolean)'
		Dispose(disposing:=True)
		GC.SuppressFinalize(Me)
	End Sub
End Class

'******************************************************************************
'*                                 27/07/2020                                 *
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

