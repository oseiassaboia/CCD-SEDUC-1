Imports Microsoft.VisualBasic
Imports System.Data

Public Class TipoSala
	Private TG60_ID_TIPO_SALA as Integer
	Private TG60_NM_TIPO_SALA as String

	Public Property IdTipoSala() as Integer
		Get
			Return TG60_ID_TIPO_SALA
		End Get
		Set(ByVal Value As Integer)
			TG60_ID_TIPO_SALA = Value
		End Set
	End Property
	Public Property Descricao() as String
		Get
			Return TG60_NM_TIPO_SALA
		End Get
		Set(ByVal Value As String)
			TG60_NM_TIPO_SALA = Value
		End Set
	End Property

    Public Sub New(Optional ByVal IdTipoSala As Integer = 0)
        If IdTipoSala > 0 Then
            Obter(IdTipoSala)
        End If
    End Sub

	Public Sub Salvar()
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append("Then select * ")
		strSQL.Append(" from TG60_TIPO_SALA")
		strSQL.Append(" where TG60_ID_TIPO_SALA = " & IdTipoSala)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("TG60_NM_TIPO_SALA") = ProBanco(TG60_NM_TIPO_SALA, eTipoValor.TEXTO)

		cnn.SalvarDataTable(dr)

		dt.Dispose()
		dt = Nothing

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Sub Obter(ByVal IdTipoSala as String)
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from TG60_TIPO_SALA")
		strSQL.Append(" where TG60_ID_TIPO_SALA = " & IdTipoSala)

		dt = cnn.AbrirDataTable(strSQL.ToString)

		If dt.Rows.Count > 0 Then
			dr = dt.Rows(0)
			
			TG60_ID_TIPO_SALA = DoBanco(dr("TG60_ID_TIPO_SALA"), eTipoValor.CHAVE)
			TG60_NM_TIPO_SALA = DoBanco(dr("TG60_NM_TIPO_SALA"), eTipoValor.TEXTO)
		End If

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Function Pesquisar(Optional ByVal Sort as String = "", Optional IdTipoSala as Integer = 0, Optional Descricao as String = "") as DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from TG60_TIPO_SALA")
		'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
		strSQL.Append(" where TG60_ID_TIPO_SALA is not null")
		
		If IdTipoSala > 0 then 
			strSQL.Append(" and TG60_ID_TIPO_SALA = " & IdTipoSala)
		End If
		
		If Descricao <> "" then 
			strSQL.Append(" and upper(TG60_NM_TIPO_SALA) like '%" & Descricao.toUpper & "%'")
		End If
		
		strSQL.Append(" Order By " & IIf(Sort = "", "TG60_ID_TIPO_SALA", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

	Public Function ObterTabela() as DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select TG60_ID_TIPO_SALA as CODIGO, TG60_NM_TIPO_SALA as DESCRICAO")
		strSQL.Append(" from TG60_TIPO_SALA")
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
		
		strSQL.Append(" select max(TG60_ID_TIPO_SALA) from TG60_TIPO_SALA")

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
	Public Function Excluir(ByVal IdTipoSala as String) As Integer
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		Dim LinhasAfetadas As Integer
		
		strSQL.Append(" delete ")
		strSQL.Append(" from TG60_TIPO_SALA")
		strSQL.Append(" where TG60_ID_TIPO_SALA = " & IdTipoSala)

		LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

		cnn.FecharBanco()
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

