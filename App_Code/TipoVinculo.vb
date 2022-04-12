Imports Microsoft.VisualBasic
Imports System.Data

Public Class TipoVinculo
	Private RH05_ID_TIPO_VINCULO as Integer
	Private RH05_NM_TIPO_VINCULO as String

	Public Property TipoVinculoId() as Integer
		Get
			Return RH05_ID_TIPO_VINCULO
		End Get
		Set(ByVal Value As Integer)
			RH05_ID_TIPO_VINCULO = Value
		End Set
	End Property
	Public Property TipoVinculo() as String
		Get
			Return RH05_NM_TIPO_VINCULO
		End Get
		Set(ByVal Value As String)
			RH05_NM_TIPO_VINCULO = Value
		End Set
	End Property

    Public Sub New(Optional ByVal TipoVinculoId As Integer = 0)
        If TipoVinculoId > 0 Then
            Obter(TipoVinculoId)
        End If
    End Sub

	Public Sub Salvar()
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append("Then select * ")
		strSQL.Append(" from RH05_TIPO_VINCULO")
		strSQL.Append(" where RH05_ID_TIPO_VINCULO = " & TipoVinculoId)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("RH05_NM_TIPO_VINCULO") = ProBanco(RH05_NM_TIPO_VINCULO, eTipoValor.TEXTO)

		cnn.SalvarDataTable(dr)

		dt.Dispose()
		dt = Nothing

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Sub Obter(ByVal TipoVinculoId as String)
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH05_TIPO_VINCULO")
		strSQL.Append(" where RH05_ID_TIPO_VINCULO = " & TipoVinculoId)

		dt = cnn.AbrirDataTable(strSQL.ToString)

		If dt.Rows.Count > 0 Then
			dr = dt.Rows(0)
			
			RH05_ID_TIPO_VINCULO = DoBanco(dr("RH05_ID_TIPO_VINCULO"), eTipoValor.CHAVE)
			RH05_NM_TIPO_VINCULO = DoBanco(dr("RH05_NM_TIPO_VINCULO"), eTipoValor.TEXTO)
		End If

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Function Pesquisar(Optional ByVal Sort as String = "", Optional TipoVinculoId as Integer = 0, Optional TipoVinculo as String = "") as DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH05_TIPO_VINCULO")
		'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
		strSQL.Append(" where RH05_ID_TIPO_VINCULO is not null")
		
		If TipoVinculoId > 0 then 
			strSQL.Append(" and RH05_ID_TIPO_VINCULO = " & TipoVinculoId)
		End If
		
		If TipoVinculo <> "" then 
			strSQL.Append(" and upper(RH05_NM_TIPO_VINCULO) like '%" & TipoVinculo.toUpper & "%'")
		End If
		
		strSQL.Append(" Order By " & IIf(Sort = "", "RH05_ID_TIPO_VINCULO", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

	Public Function ObterTabela() as DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select RH05_ID_TIPO_VINCULO as CODIGO, RH05_NM_TIPO_VINCULO as DESCRICAO")
		strSQL.Append(" from RH05_TIPO_VINCULO")
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
		
		strSQL.Append(" select max(RH05_ID_TIPO_VINCULO) from RH05_TIPO_VINCULO")

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
	Public Function Excluir(ByVal TipoVinculoId as String) As Integer
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		Dim LinhasAfetadas As Integer
		
		strSQL.Append(" delete ")
		strSQL.Append(" from RH05_TIPO_VINCULO")
		strSQL.Append(" where RH05_ID_TIPO_VINCULO = " & TipoVinculoId)

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

