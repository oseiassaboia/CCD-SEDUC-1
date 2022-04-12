
Imports System.Data

Public Class GrauInstrucao
	Private TG14_ID_GRAU_INSTRUCAO as Integer
	Private TG14_NM_GRAU_INSTRUCAO as String

	Public Property GrauInstrucaoId() as Integer
		Get
			Return TG14_ID_GRAU_INSTRUCAO
		End Get
		Set(ByVal Value As Integer)
			TG14_ID_GRAU_INSTRUCAO = Value
		End Set
	End Property
	Public Property GrauInstrucao() as String
		Get
			Return TG14_NM_GRAU_INSTRUCAO
		End Get
		Set(ByVal Value As String)
			TG14_NM_GRAU_INSTRUCAO = Value
		End Set
	End Property

    Public Sub New(Optional ByVal GrauInstrucaoId As Integer = 0)
        If GrauInstrucaoId > 0 Then
            Obter(GrauInstrucaoId)
        End If
    End Sub

	Public Sub Salvar()
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append("Then select * ")
		strSQL.Append(" from DBGERAL..TG14_GRAU_INSTRUCAO")
		strSQL.Append(" where TG14_ID_GRAU_INSTRUCAO = " & GrauInstrucaoId)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("TG14_NM_GRAU_INSTRUCAO") = ProBanco(TG14_NM_GRAU_INSTRUCAO, eTipoValor.TEXTO)

		cnn.SalvarDataTable(dr)

		dt.Dispose()
		dt = Nothing

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Sub Obter(ByVal GrauInstrucaoId as String)
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from DBGERAL..TG14_GRAU_INSTRUCAO")
		strSQL.Append(" where TG14_ID_GRAU_INSTRUCAO = " & GrauInstrucaoId)

		dt = cnn.AbrirDataTable(strSQL.ToString)

		If dt.Rows.Count > 0 Then
			dr = dt.Rows(0)
			
			TG14_ID_GRAU_INSTRUCAO = DoBanco(dr("TG14_ID_GRAU_INSTRUCAO"), eTipoValor.CHAVE)
			TG14_NM_GRAU_INSTRUCAO = DoBanco(dr("TG14_NM_GRAU_INSTRUCAO"), eTipoValor.TEXTO)
		End If

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Function Pesquisar(Optional ByVal Sort as String = "", Optional GrauInstrucaoId as Integer = 0, Optional GrauInstrucao as String = "") as DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from DBGERAL..TG14_GRAU_INSTRUCAO")
		'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
		strSQL.Append(" where TG14_ID_GRAU_INSTRUCAO is not null")
		
		If GrauInstrucaoId > 0 then 
			strSQL.Append(" and TG14_ID_GRAU_INSTRUCAO = " & GrauInstrucaoId)
		End If
		
		If GrauInstrucao <> "" then 
			strSQL.Append(" and upper(TG14_NM_GRAU_INSTRUCAO) like '%" & GrauInstrucao.toUpper & "%'")
		End If
		
		strSQL.Append(" Order By " & IIf(Sort = "", "TG14_ID_GRAU_INSTRUCAO", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

	Public Function ObterTabela() as DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select TG14_ID_GRAU_INSTRUCAO as CODIGO, TG14_NM_GRAU_INSTRUCAO as DESCRICAO")
		strSQL.Append(" from DBGERAL..TG14_GRAU_INSTRUCAO")
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
		
		strSQL.Append(" select max(TG14_ID_GRAU_INSTRUCAO) from DBGERAL..TG14_GRAU_INSTRUCAO")

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
	Public Function Excluir(ByVal GrauInstrucaoId as String) As Integer
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		Dim LinhasAfetadas As Integer
		
		strSQL.Append(" delete ")
		strSQL.Append(" from DBGERAL..TG14_GRAU_INSTRUCAO")
		strSQL.Append(" where TG14_ID_GRAU_INSTRUCAO = " & GrauInstrucaoId)

		LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

		cnn.FecharBanco()
		cnn = Nothing

		Return LinhasAfetadas
	End Function

End Class

'******************************************************************************
'*                                 10/09/2018                                 *
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

