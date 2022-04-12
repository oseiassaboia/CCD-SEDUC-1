Imports Microsoft.VisualBasic
Imports System.Data

Public Class TipoPredio
	Private TG57_ID_TIPO_PREDIO as Integer
	Private TG57_NM_TIPO_PREDIO as String

	Public Property IdPredio() as Integer
		Get
			Return TG57_ID_TIPO_PREDIO
		End Get
		Set(ByVal Value As Integer)
			TG57_ID_TIPO_PREDIO = Value
		End Set
	End Property
	Public Property Descricao() as String
		Get
			Return TG57_NM_TIPO_PREDIO
		End Get
		Set(ByVal Value As String)
			TG57_NM_TIPO_PREDIO = Value
		End Set
	End Property

    Public Sub New(Optional ByVal IdPredio As Integer = 0)
        If IdPredio > 0 Then
            Obter(IdPredio)
        End If
    End Sub

	Public Sub Salvar()
        Dim cnn As New ConexaoGeral
        Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append("Then select * ")
		strSQL.Append(" from TG57_TIPO_PREDIO")
		strSQL.Append(" where TG57_ID_TIPO_PREDIO = " & IdPredio)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("TG57_NM_TIPO_PREDIO") = ProBanco(TG57_NM_TIPO_PREDIO, eTipoValor.TEXTO)

		cnn.SalvarDataTable(dr)

		dt.Dispose()
		dt = Nothing


        cnn = Nothing
	End Sub

	Public Sub Obter(ByVal IdPredio as String)
        Dim cnn As New ConexaoGeral
        Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from TG57_TIPO_PREDIO")
		strSQL.Append(" where TG57_ID_TIPO_PREDIO = " & IdPredio)

		dt = cnn.AbrirDataTable(strSQL.ToString)

		If dt.Rows.Count > 0 Then
			dr = dt.Rows(0)
			
			TG57_ID_TIPO_PREDIO = DoBanco(dr("TG57_ID_TIPO_PREDIO"), eTipoValor.CHAVE)
			TG57_NM_TIPO_PREDIO = DoBanco(dr("TG57_NM_TIPO_PREDIO"), eTipoValor.TEXTO)
		End If


        cnn = Nothing
	End Sub

	Public Function Pesquisar(Optional ByVal Sort as String = "", Optional IdPredio as Integer = 0, Optional Descricao as String = "") as DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from TG57_TIPO_PREDIO")
		'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
		strSQL.Append(" where TG57_ID_TIPO_PREDIO is not null")
		
		If IdPredio > 0 then 
			strSQL.Append(" and TG57_ID_TIPO_PREDIO = " & IdPredio)
		End If
		
		If Descricao <> "" then 
			strSQL.Append(" and upper(TG57_NM_TIPO_PREDIO) like '%" & Descricao.toUpper & "%'")
		End If
		
		strSQL.Append(" Order By " & IIf(Sort = "", "TG57_ID_TIPO_PREDIO", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

	Public Function ObterTabela() as DataTable
        Dim cnn As New ConexaoGeral
        Dim dt As DataTable
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select TG57_ID_TIPO_PREDIO as CODIGO, TG57_NM_TIPO_PREDIO as DESCRICAO")
		strSQL.Append(" from TG57_TIPO_PREDIO")
		strSQL.Append(" order by 2 ")

		dt = cnn.AbrirDataTable(strSQL.ToString)


        cnn = Nothing

        Return dt
	End Function

	Public Function ObterUltimo() as Integer
        Dim cnn As New ConexaoGeral
        Dim strSQL As New StringBuilder
		Dim CodigoUltimo As Integer
		
		strSQL.Append(" select max(TG57_ID_TIPO_PREDIO) from TG57_TIPO_PREDIO")

		With cnn.AbrirDataTable(strSQL.ToString)
			If Not IsDBNull(.Rows(0)(0)) Then
				CodigoUltimo = .Rows(0)(0)
			Else
				CodigoUltimo = 0
			End If
		End With

        cnn = Nothing

        Return CodigoUltimo

	End Function
	Public Function Excluir(ByVal IdPredio as String) As Integer
        Dim cnn As New ConexaoGeral
        Dim strSQL As New StringBuilder
		Dim LinhasAfetadas As Integer
		
		strSQL.Append(" delete ")
		strSQL.Append(" from TG57_TIPO_PREDIO")
		strSQL.Append(" where TG57_ID_TIPO_PREDIO = " & IdPredio)

		LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

        cnn = Nothing

        Return LinhasAfetadas
	End Function

End Class

'******************************************************************************
'*                                 20/09/2019                                 *
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

