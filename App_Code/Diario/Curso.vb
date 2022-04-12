Imports Microsoft.VisualBasic
Imports System.Data

Public Class Curso
	Private DE15_ID_CURSO as Integer
	Private DE12_ID_AREA as Integer
	Private DE15_NM_CURSO as String
	Private DE15_CD_CURSO_CENSO as String

	Public Property Codigo() as Integer
		Get
			Return DE15_ID_CURSO
		End Get
		Set(ByVal Value As Integer)
			DE15_ID_CURSO = Value
		End Set
	End Property
	Public Property Area() as Integer
		Get
			Return DE12_ID_AREA
		End Get
		Set(ByVal Value As Integer)
			DE12_ID_AREA = Value
		End Set
	End Property
	Public Property Nome() as String
		Get
			Return DE15_NM_CURSO
		End Get
		Set(ByVal Value As String)
			DE15_NM_CURSO = Value
		End Set
	End Property
	Public Property CodigoCenso() as String
		Get
			Return DE15_CD_CURSO_CENSO
		End Get
		Set(ByVal Value As String)
			DE15_CD_CURSO_CENSO = Value
		End Set
	End Property

    Public Sub New(Optional ByVal Codigo As Integer = 0)
        If Codigo > 0 Then
            Obter(Codigo)
        End If
    End Sub

    Public Sub Salvar()
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" Select * ")
        strSQL.Append(" from DBDIARIO..DE15_CURSO")
        strSQL.Append(" where DE15_ID_CURSO = " & Codigo)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("DE12_ID_AREA") = ProBanco(DE12_ID_AREA, eTipoValor.CHAVE)
		dr("DE15_NM_CURSO") = ProBanco(DE15_NM_CURSO, eTipoValor.TEXTO)
		dr("DE15_CD_CURSO_CENSO") = ProBanco(DE15_CD_CURSO_CENSO, eTipoValor.TEXTO)

		cnn.SalvarDataTable(dr)

		dt.Dispose()
		dt = Nothing

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Sub Obter(ByVal Codigo as String)
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" Select * ")
        strSQL.Append(" from DBDIARIO..DE15_CURSO")
        strSQL.Append(" where DE15_ID_CURSO = " & Codigo)

		dt = cnn.AbrirDataTable(strSQL.ToString)

		If dt.Rows.Count > 0 Then
			dr = dt.Rows(0)
			
			DE15_ID_CURSO = DoBanco(dr("DE15_ID_CURSO"), eTipoValor.CHAVE)
            DE12_ID_AREA = DoBanco(dr("DE12_ID_AREA"), eTipoValor.CHAVE)
            DE15_NM_CURSO = DoBanco(dr("DE15_NM_CURSO"), eTipoValor.TEXTO)
			DE15_CD_CURSO_CENSO = DoBanco(dr("DE15_CD_CURSO_CENSO"), eTipoValor.TEXTO)
		End If

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Function Pesquisar(Optional ByVal Sort as String = "", Optional Codigo as Integer = 0, Optional Area as Integer = 0, Optional Nome as String = "", Optional CodigoCenso as String = "") as DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder

        strSQL.Append(" Select DE15.DE15_ID_CURSO, DE15.DE15_NM_CURSO, DE15.DE15_CD_CURSO_CENSO ")
        strSQL.Append(" , DE12.DE12_ID_AREA, DE12.DE12_NM_AREA")
        strSQL.Append(" from DBDIARIO..DE15_CURSO as DE15")
        strSQL.Append(" left join DE12_AREA as DE12 On DE12.DE12_ID_AREA = DE15.DE12_ID_AREA ")
        strSQL.Append(" where DE15.DE15_ID_CURSO Is Not null")

        If Codigo > 0 Then
            strSQL.Append(" And DE15.DE15_ID_CURSO = " & Codigo)
        End If
		
		If Area > 0 Then
            strSQL.Append(" And DE12.DE12_ID_AREA = " & Area)
        End If
		
		If Nome <> "" Then
            strSQL.Append(" And upper(DE15.DE15_NM_CURSO) like '%" & Nome.ToUpper & "%'")
        End If
		
		If CodigoCenso <> "" Then
            strSQL.Append(" and upper(DE15.DE15_CD_CURSO_CENSO) like '%" & CodigoCenso.ToUpper & "%'")
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "DE15.DE15_ID_CURSO", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

    Public Function ObterTabela(Optional Area As Integer = 0) As DataTable
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" select DE15_ID_CURSO as CODIGO, DE15_NM_CURSO as DESCRICAO")
        strSQL.Append(" from DBDIARIO..DE15_CURSO")
        strSQL.Append(" where DE15_ID_CURSO Is Not null")

        If Area > 0 Then
            strSQL.Append(" And DE12_ID_AREA = " & Area)
        End If

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

        strSQL.Append(" select max(DE15_ID_CURSO) from DBDIARIO..DE15_CURSO")

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
	Public Function Excluir(ByVal Codigo as String) As Integer
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		Dim LinhasAfetadas As Integer
		
		strSQL.Append(" delete ")
        strSQL.Append(" from DBDIARIO..DE15_CURSO")
        strSQL.Append(" where DE15_ID_CURSO = " & Codigo)

		LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

		cnn.FecharBanco()
		cnn = Nothing

		Return LinhasAfetadas
	End Function

End Class

'******************************************************************************
'*                                 16/10/2018                                 *
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

