Imports Microsoft.VisualBasic
Imports System.Data

Public Class SituacaoServidor
	Private RH07_ID_SITUACAO_SERVIDOR as Integer
	Private RH07_NM_SITUACAO_SERVIDOR as String

	Public Property SituacaoId() as Integer
		Get
			Return RH07_ID_SITUACAO_SERVIDOR
		End Get
		Set(ByVal Value As Integer)
			RH07_ID_SITUACAO_SERVIDOR = Value
		End Set
	End Property
	Public Property Situacao() as String
		Get
			Return RH07_NM_SITUACAO_SERVIDOR
		End Get
		Set(ByVal Value As String)
			RH07_NM_SITUACAO_SERVIDOR = Value
		End Set
	End Property

    Public Sub New(Optional ByVal SituacaoId As Integer = 0)
        If SituacaoId > 0 Then
            Obter(SituacaoId)
        End If
    End Sub

	Public Sub Salvar()
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append("Then select * ")
		strSQL.Append(" from RH07_SITUACAO_SERVIDOR")
		strSQL.Append(" where RH07_ID_SITUACAO_SERVIDOR = " & SituacaoId)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("RH07_NM_SITUACAO_SERVIDOR") = ProBanco(RH07_NM_SITUACAO_SERVIDOR, eTipoValor.TEXTO)

		cnn.SalvarDataTable(dr)

		dt.Dispose()
		dt = Nothing

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Sub Obter(ByVal SituacaoId as String)
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH07_SITUACAO_SERVIDOR")
		strSQL.Append(" where RH07_ID_SITUACAO_SERVIDOR = " & SituacaoId)

		dt = cnn.AbrirDataTable(strSQL.ToString)

		If dt.Rows.Count > 0 Then
			dr = dt.Rows(0)
			
			RH07_ID_SITUACAO_SERVIDOR = DoBanco(dr("RH07_ID_SITUACAO_SERVIDOR"), eTipoValor.CHAVE)
			RH07_NM_SITUACAO_SERVIDOR = DoBanco(dr("RH07_NM_SITUACAO_SERVIDOR"), eTipoValor.TEXTO)
		End If

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Function Pesquisar(Optional ByVal Sort as String = "", Optional SituacaoId as Integer = 0, Optional Situacao as String = "") as DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH07_SITUACAO_SERVIDOR")
		'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
		strSQL.Append(" where RH07_ID_SITUACAO_SERVIDOR is not null")
		
		If SituacaoId > 0 then 
			strSQL.Append(" and RH07_ID_SITUACAO_SERVIDOR = " & SituacaoId)
		End If
		
		If Situacao <> "" then 
			strSQL.Append(" and upper(RH07_NM_SITUACAO_SERVIDOR) like '%" & Situacao.toUpper & "%'")
		End If
		
		strSQL.Append(" Order By " & IIf(Sort = "", "RH07_ID_SITUACAO_SERVIDOR", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

	Public Function ObterTabela() as DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select RH07_ID_SITUACAO_SERVIDOR as CODIGO, RH07_NM_SITUACAO_SERVIDOR as DESCRICAO")
		strSQL.Append(" from RH07_SITUACAO_SERVIDOR")
		strSQL.Append(" order by 2 ")

		dt = cnn.AbrirDataTable(strSQL.ToString)

		cnn.FecharBanco()
		cnn = Nothing

		Return dt
	End Function

	Public Function ObterUltimo() as Integer
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		Dim CodigoUltimo As Integer
		
		strSQL.Append(" select max(RH07_ID_SITUACAO_SERVIDOR) from RH07_SITUACAO_SERVIDOR")

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
	Public Function Excluir(ByVal SituacaoId as String) As Integer
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		Dim LinhasAfetadas As Integer
		
		strSQL.Append(" delete ")
		strSQL.Append(" from RH07_SITUACAO_SERVIDOR")
		strSQL.Append(" where RH07_ID_SITUACAO_SERVIDOR = " & SituacaoId)

		LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

		cnn.FecharBanco()
		cnn = Nothing

		Return LinhasAfetadas
	End Function

End Class

'******************************************************************************
'*                                 11/09/2018                                 *
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

