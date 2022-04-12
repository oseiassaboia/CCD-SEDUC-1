Imports Microsoft.VisualBasic
Imports System.Data

Public Class Sala
	Private TG61_ID_SALA as Integer
	Private TG59_ID_PREDIO as String
	Private TG60_ID_TIPO_SALA as String
	Private TG61_NU_SALA as String
	Private TG61_NM_SALA as String

	Public Property IdSala() as Integer
		Get
			Return TG61_ID_SALA
		End Get
		Set(ByVal Value As Integer)
			TG61_ID_SALA = Value
		End Set
	End Property
	Public Property IdPredio() as String
		Get
			Return TG59_ID_PREDIO
		End Get
		Set(ByVal Value As String)
			TG59_ID_PREDIO = Value
		End Set
	End Property
	Public Property idTipoSala() as String
		Get
			Return TG60_ID_TIPO_SALA
		End Get
		Set(ByVal Value As String)
			TG60_ID_TIPO_SALA = Value
		End Set
	End Property
	Public Property NumeroSala() as String
		Get
			Return TG61_NU_SALA
		End Get
		Set(ByVal Value As String)
			TG61_NU_SALA = Value
		End Set
	End Property
	Public Property Descricao() as String
		Get
			Return TG61_NM_SALA
		End Get
		Set(ByVal Value As String)
			TG61_NM_SALA = Value
		End Set
	End Property

    Public Sub New(Optional ByVal IdSala As Integer = 0)
        If IdSala > 0 Then
            Obter(IdSala)
        End If
    End Sub

	Public Sub Salvar()
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append("Then select * ")
		strSQL.Append(" from TG61_SALA")
		strSQL.Append(" where TG61_ID_SALA = " & IdSala)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("TG59_ID_PREDIO") = ProBanco(TG59_ID_PREDIO, eTipoValor.NUMERO_DECIMAL)
		dr("TG60_ID_TIPO_SALA") = ProBanco(TG60_ID_TIPO_SALA, eTipoValor.NUMERO_DECIMAL)
		dr("TG61_NU_SALA") = ProBanco(TG61_NU_SALA, eTipoValor.TEXTO)
		dr("TG61_NM_SALA") = ProBanco(TG61_NM_SALA, eTipoValor.TEXTO)

		cnn.SalvarDataTable(dr)

		dt.Dispose()
		dt = Nothing

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Sub Obter(ByVal IdSala as String)
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from TG61_SALA")
		strSQL.Append(" where TG61_ID_SALA = " & IdSala)

		dt = cnn.AbrirDataTable(strSQL.ToString)

		If dt.Rows.Count > 0 Then
			dr = dt.Rows(0)
			
			TG61_ID_SALA = DoBanco(dr("TG61_ID_SALA"), eTipoValor.CHAVE)
			TG59_ID_PREDIO = DoBanco(dr("TG59_ID_PREDIO"), eTipoValor.NUMERO_DECIMAL)
			TG60_ID_TIPO_SALA = DoBanco(dr("TG60_ID_TIPO_SALA"), eTipoValor.NUMERO_DECIMAL)
			TG61_NU_SALA = DoBanco(dr("TG61_NU_SALA"), eTipoValor.TEXTO)
			TG61_NM_SALA = DoBanco(dr("TG61_NM_SALA"), eTipoValor.TEXTO)
		End If

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Function Pesquisar(Optional ByVal Sort as String = "", Optional IdSala as Integer = 0, Optional IdPredio as String = "", Optional idTipoSala as String = "", Optional NumeroSala as String = "", Optional Descricao as String = "") as DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from TG61_SALA")
		'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
		strSQL.Append(" where TG61_ID_SALA is not null")
		
		If IdSala > 0 then 
			strSQL.Append(" and TG61_ID_SALA = " & IdSala)
		End If
		
		If IsNumeric(IdPredio.Replace(".", "")) then
			strSQL.Append(" and TG59_ID_PREDIO = " & IdPredio.Replace(".", "").Replace(",", "."))
		End If
		
		If IsNumeric(idTipoSala.Replace(".", "")) then
			strSQL.Append(" and TG60_ID_TIPO_SALA = " & idTipoSala.Replace(".", "").Replace(",", "."))
		End If
		
		If NumeroSala <> "" then 
			strSQL.Append(" and upper(TG61_NU_SALA) like '%" & NumeroSala.toUpper & "%'")
		End If
		
		If Descricao <> "" then 
			strSQL.Append(" and upper(TG61_NM_SALA) like '%" & Descricao.toUpper & "%'")
		End If
		
		strSQL.Append(" Order By " & IIf(Sort = "", "TG61_ID_SALA", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

	Public Function ObterTabela() as DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select TG61_ID_SALA as CODIGO, TG59_ID_PREDIO as DESCRICAO")
		strSQL.Append(" from TG61_SALA")
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
		
		strSQL.Append(" select max(TG61_ID_SALA) from TG61_SALA")

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
	Public Function Excluir(ByVal IdSala as String) As Integer
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		Dim LinhasAfetadas As Integer
		
		strSQL.Append(" delete ")
		strSQL.Append(" from TG61_SALA")
		strSQL.Append(" where TG61_ID_SALA = " & IdSala)

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

