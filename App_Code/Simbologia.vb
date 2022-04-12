Imports Microsoft.VisualBasic
Imports System.Data

Public Class Simbologia
	Private RH35_ID_SIMBOLOGIA as Integer
	Private RH35_NM_SIMBOLOGIA as String
	Private RH35_SG_SIMBOLOGIA as String

	Public Property SimbologiaId() as Integer
		Get
			Return RH35_ID_SIMBOLOGIA
		End Get
		Set(ByVal Value As Integer)
			RH35_ID_SIMBOLOGIA = Value
		End Set
	End Property
	Public Property Simbologia() as String
		Get
			Return RH35_NM_SIMBOLOGIA
		End Get
		Set(ByVal Value As String)
			RH35_NM_SIMBOLOGIA = Value
		End Set
	End Property
	Public Property Sigla() as String
		Get
			Return RH35_SG_SIMBOLOGIA
		End Get
		Set(ByVal Value As String)
			RH35_SG_SIMBOLOGIA = Value
		End Set
	End Property

    Public Sub New(Optional ByVal SimbologiaId As Integer = 0)
        If SimbologiaId > 0 Then
            Obter(SimbologiaId)
        End If
    End Sub

	Public Sub Salvar()
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append("Then select * ")
		strSQL.Append(" from RH35_SIMBOLOGIA")
		strSQL.Append(" where RH35_ID_SIMBOLOGIA = " & SimbologiaId)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("RH35_NM_SIMBOLOGIA") = ProBanco(RH35_NM_SIMBOLOGIA, eTipoValor.TEXTO)
		dr("RH35_SG_SIMBOLOGIA") = ProBanco(RH35_SG_SIMBOLOGIA, eTipoValor.TEXTO)

		cnn.SalvarDataTable(dr)

		dt.Dispose()
		dt = Nothing

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Sub Obter(ByVal SimbologiaId as String)
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH35_SIMBOLOGIA")
		strSQL.Append(" where RH35_ID_SIMBOLOGIA = " & SimbologiaId)

		dt = cnn.AbrirDataTable(strSQL.ToString)

		If dt.Rows.Count > 0 Then
			dr = dt.Rows(0)
			
			RH35_ID_SIMBOLOGIA = DoBanco(dr("RH35_ID_SIMBOLOGIA"), eTipoValor.CHAVE)
			RH35_NM_SIMBOLOGIA = DoBanco(dr("RH35_NM_SIMBOLOGIA"), eTipoValor.TEXTO)
			RH35_SG_SIMBOLOGIA = DoBanco(dr("RH35_SG_SIMBOLOGIA"), eTipoValor.TEXTO)
		End If

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Function Pesquisar(Optional ByVal Sort as String = "", Optional SimbologiaId as Integer = 0, Optional Simbologia as String = "", Optional Sigla as String = "") as DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH35_SIMBOLOGIA")
		'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
		strSQL.Append(" where RH35_ID_SIMBOLOGIA is not null")
		
		If SimbologiaId > 0 then 
			strSQL.Append(" and RH35_ID_SIMBOLOGIA = " & SimbologiaId)
		End If
		
		If Simbologia <> "" then 
			strSQL.Append(" and upper(RH35_NM_SIMBOLOGIA) like '%" & Simbologia.toUpper & "%'")
		End If
		
		If Sigla <> "" then 
			strSQL.Append(" and upper(RH35_SG_SIMBOLOGIA) like '%" & Sigla.toUpper & "%'")
		End If
		
		strSQL.Append(" Order By " & IIf(Sort = "", "RH35_ID_SIMBOLOGIA", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

	Public Function ObterTabela() as DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select RH35_ID_SIMBOLOGIA as CODIGO, RH35_NM_SIMBOLOGIA as DESCRICAO")
		strSQL.Append(" from RH35_SIMBOLOGIA")
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
		
		strSQL.Append(" select max(RH35_ID_SIMBOLOGIA) from RH35_SIMBOLOGIA")

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
	Public Function Excluir(ByVal SimbologiaId as String) As Integer
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		Dim LinhasAfetadas As Integer
		
		strSQL.Append(" delete ")
		strSQL.Append(" from RH35_SIMBOLOGIA")
		strSQL.Append(" where RH35_ID_SIMBOLOGIA = " & SimbologiaId)

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

