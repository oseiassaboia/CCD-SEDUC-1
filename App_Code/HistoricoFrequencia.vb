Imports Microsoft.VisualBasic
Imports System.Data

Public Class HistoricoFrequencia
	Private RH50_ID_HISTORICO_PERIODO as Integer
	Private RH17_ID_PERIODO_FREQ as Integer
	Private RH49_ST_PERIODO as String
	Private RH50_DH_ST_PERIODO as String

	Public Property Codigo() as Integer
		Get
			Return RH50_ID_HISTORICO_PERIODO
		End Get
		Set(ByVal Value As Integer)
			RH50_ID_HISTORICO_PERIODO = Value
		End Set
	End Property
	Public Property IdPeriodoFrequencia() as Integer
		Get
			Return RH17_ID_PERIODO_FREQ
		End Get
		Set(ByVal Value As Integer)
			RH17_ID_PERIODO_FREQ = Value
		End Set
	End Property
	Public Property SituacaoPeriodo() as String
		Get
			Return RH49_ST_PERIODO
		End Get
		Set(ByVal Value As String)
			RH49_ST_PERIODO = Value
		End Set
	End Property
	Public Property DataHoraSituacaoPeriodo() as String
		Get
			Return RH50_DH_ST_PERIODO
		End Get
		Set(ByVal Value As String)
			RH50_DH_ST_PERIODO = Value
		End Set
	End Property

	Public Sub New(Optional ByVal Codigo as Integer = 0)
		If Codigo > 0 Then
			Obter(Codigo)
		End If
	End Sub

	Public Sub Salvar()
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH50_HISTORICO_PERIODO")
		strSQL.Append(" where RH50_ID_HISTORICO_PERIODO = " & Codigo)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("RH17_ID_PERIODO_FREQ") = ProBanco(RH17_ID_PERIODO_FREQ, eTipoValor.CHAVE)
		dr("RH49_ST_PERIODO") = ProBanco(RH49_ST_PERIODO, eTipoValor.TEXTO)
		dr("RH50_DH_ST_PERIODO") = ProBanco(RH50_DH_ST_PERIODO, eTipoValor.DATA)

		cnn.SalvarDataTable(dr)

		dt.Dispose()
		dt = Nothing

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Sub Obter(ByVal Codigo as Integer)
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH50_HISTORICO_PERIODO")
		strSQL.Append(" where RH50_ID_HISTORICO_PERIODO = " & Codigo)

		dt = cnn.AbrirDataTable(strSQL.ToString)

		If dt.Rows.Count > 0 Then
			dr = dt.Rows(0)
			
			RH50_ID_HISTORICO_PERIODO = DoBanco(dr("RH50_ID_HISTORICO_PERIODO"), eTipoValor.CHAVE)
			RH17_ID_PERIODO_FREQ = DoBanco(dr("RH17_ID_PERIODO_FREQ"), eTipoValor.CHAVE)
			RH49_ST_PERIODO = DoBanco(dr("RH49_ST_PERIODO"), eTipoValor.TEXTO)
			RH50_DH_ST_PERIODO = DoBanco(dr("RH50_DH_ST_PERIODO"), eTipoValor.DATA)
		End If

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Function Pesquisar(Optional ByVal Sort as String = "", Optional Codigo as Integer = 0, Optional IdPeriodoFrequencia as Integer = 0, Optional SituacaoPeriodo as String = "", Optional DataHoraSituacaoPeriodo as String = "") as DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH50_HISTORICO_PERIODO")
		'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
		strSQL.Append(" where RH50_ID_HISTORICO_PERIODO is not null")
		
		If Codigo > 0 then 
			strSQL.Append(" and RH50_ID_HISTORICO_PERIODO = " & Codigo)
		End If
		
		If IdPeriodoFrequencia > 0 then 
			strSQL.Append(" and RH17_ID_PERIODO_FREQ = " & IdPeriodoFrequencia)
		End If
		
		If SituacaoPeriodo <> "" then 
			strSQL.Append(" and upper(RH49_ST_PERIODO) like '%" & SituacaoPeriodo.toUpper & "%'")
		End If
		
		If isDate(DataHoraSituacaoPeriodo) then 
			strSQL.Append(" and RH50_DH_ST_PERIODO = Convert(DateTime, '" & DataHoraSituacaoPeriodo & "', 103)")
		End If
		
		strSQL.Append(" Order By " & IIf(Sort = "", "RH50_ID_HISTORICO_PERIODO", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

	Public Function ObterTabela() as DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select RH50_ID_HISTORICO_PERIODO as CODIGO, RH17_ID_PERIODO_FREQ as DESCRICAO")
		strSQL.Append(" from RH50_HISTORICO_PERIODO")
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
		
		strSQL.Append(" select max(RH50_ID_HISTORICO_PERIODO) from RH50_HISTORICO_PERIODO")

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
	Public Function Excluir(ByVal Codigo as Integer) As Integer
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		Dim LinhasAfetadas As Integer
		
		strSQL.Append(" delete ")
		strSQL.Append(" from RH50_HISTORICO_PERIODO")
		strSQL.Append(" where RH50_ID_HISTORICO_PERIODO = " & Codigo)

		LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

		cnn.FecharBanco()
		cnn = Nothing

		Return LinhasAfetadas
	End Function

End Class

'******************************************************************************
'*                                 20/05/2019                                 *
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

