Imports Microsoft.VisualBasic
Imports System.Data

Public Class TipoLotacao
	Private RH32_ID_TIPO_LOTACAO as Integer
	Private RH32_NM_TIPO_LOTACAO as String

	Public Property TipoLotacaoId() as Integer
		Get
			Return RH32_ID_TIPO_LOTACAO
		End Get
		Set(ByVal Value As Integer)
			RH32_ID_TIPO_LOTACAO = Value
		End Set
	End Property
	Public Property TipoLotacao() as String
		Get
			Return RH32_NM_TIPO_LOTACAO
		End Get
		Set(ByVal Value As String)
			RH32_NM_TIPO_LOTACAO = Value
		End Set
	End Property

    Public Sub New(Optional ByVal TipoLotacaoId As Integer = 0)
        If TipoLotacaoId > 0 Then
            Obter(TipoLotacaoId)
        End If
    End Sub

	Public Sub Salvar()
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append("Then select * ")
		strSQL.Append(" from RH32_TIPO_LOTACAO")
		strSQL.Append(" where RH32_ID_TIPO_LOTACAO = " & TipoLotacaoId)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("RH32_NM_TIPO_LOTACAO") = ProBanco(RH32_NM_TIPO_LOTACAO, eTipoValor.TEXTO)

		cnn.SalvarDataTable(dr)

		dt.Dispose()
		dt = Nothing

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Sub Obter(ByVal TipoLotacaoId as String)
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH32_TIPO_LOTACAO")
		strSQL.Append(" where RH32_ID_TIPO_LOTACAO = " & TipoLotacaoId)

		dt = cnn.AbrirDataTable(strSQL.ToString)

		If dt.Rows.Count > 0 Then
			dr = dt.Rows(0)
			
			RH32_ID_TIPO_LOTACAO = DoBanco(dr("RH32_ID_TIPO_LOTACAO"), eTipoValor.CHAVE)
			RH32_NM_TIPO_LOTACAO = DoBanco(dr("RH32_NM_TIPO_LOTACAO"), eTipoValor.TEXTO)
		End If

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Function Pesquisar(Optional ByVal Sort as String = "", Optional TipoLotacaoId as Integer = 0, Optional TipoLotacao as String = "") as DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH32_TIPO_LOTACAO")
		'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
		strSQL.Append(" where RH32_ID_TIPO_LOTACAO is not null")
		
		If TipoLotacaoId > 0 then 
			strSQL.Append(" and RH32_ID_TIPO_LOTACAO = " & TipoLotacaoId)
		End If
		
		If TipoLotacao <> "" then 
			strSQL.Append(" and upper(RH32_NM_TIPO_LOTACAO) like '%" & TipoLotacao.toUpper & "%'")
		End If
		
		strSQL.Append(" Order By " & IIf(Sort = "", "RH32_ID_TIPO_LOTACAO", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

	Public Function ObterTabela() as DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select RH32_ID_TIPO_LOTACAO as CODIGO, RH32_NM_TIPO_LOTACAO as DESCRICAO")
		strSQL.Append(" from RH32_TIPO_LOTACAO")
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
		
		strSQL.Append(" select max(RH32_ID_TIPO_LOTACAO) from RH32_TIPO_LOTACAO")

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
	Public Function Excluir(ByVal TipoLotacaoId as String) As Integer
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		Dim LinhasAfetadas As Integer
		
		strSQL.Append(" delete ")
		strSQL.Append(" from RH32_TIPO_LOTACAO")
		strSQL.Append(" where RH32_ID_TIPO_LOTACAO = " & TipoLotacaoId)

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

