Imports Microsoft.VisualBasic
Imports System.Data

Public Class Periodo
	Private DE04_ID_PERIODO as Integer
	Private DE04_ID_PERIODO_1 as Integer
	Private DE04_TP_PERIODO as String
	Private DE04_NM_PERIODO as String

	Public Property Codigo() as Integer
		Get
			Return DE04_ID_PERIODO
		End Get
		Set(ByVal Value As Integer)
			DE04_ID_PERIODO = Value
		End Set
	End Property
	Public Property Periodo() as Integer
		Get
			Return DE04_ID_PERIODO_1
		End Get
		Set(ByVal Value As Integer)
			DE04_ID_PERIODO_1 = Value
		End Set
	End Property
	Public Property TipoPeriodo() as String
		Get
			Return DE04_TP_PERIODO
		End Get
		Set(ByVal Value As String)
			DE04_TP_PERIODO = Value
		End Set
	End Property
	Public Property Numero() as String
		Get
			Return DE04_NM_PERIODO
		End Get
		Set(ByVal Value As String)
			DE04_NM_PERIODO = Value
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
		
		strSQL.Append(" select * ")
		strSQL.Append(" from DE04_PERIODO")
		strSQL.Append(" where DE04_ID_PERIODO = " & Codigo)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("DE04_ID_PERIODO_1") = ProBanco(DE04_ID_PERIODO_1, eTipoValor.CHAVE)
		dr("DE04_TP_PERIODO") = ProBanco(DE04_TP_PERIODO, eTipoValor.TEXTO)
		dr("DE04_NM_PERIODO") = ProBanco(DE04_NM_PERIODO, eTipoValor.TEXTO)

		cnn.SalvarDataTable(dr)

		dt.Dispose()
		dt = Nothing

		
		cnn = Nothing
	End Sub

    Public Sub Obter(ByVal Codigo As Integer)
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from DE04_PERIODO")
        strSQL.Append(" where DE04_ID_PERIODO = " & Codigo)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            DE04_ID_PERIODO = DoBanco(dr("DE04_ID_PERIODO"), eTipoValor.CHAVE)
            DE04_ID_PERIODO_1 = DoBanco(dr("DE04_ID_PERIODO_1"), eTipoValor.CHAVE)
            DE04_TP_PERIODO = DoBanco(dr("DE04_TP_PERIODO"), eTipoValor.TEXTO)
            DE04_NM_PERIODO = DoBanco(dr("DE04_NM_PERIODO"), eTipoValor.TEXTO)
        End If

        
        cnn = Nothing
    End Sub

    Public Function Pesquisar(Optional ByVal Sort as String = "", Optional Codigo as Integer = 0, Optional Periodo as Integer = 0, Optional TipoPeriodo as String = "", Optional Numero as String = "") as DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from DE04_PERIODO")
		'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
		strSQL.Append(" where DE04_ID_PERIODO is not null")
		
		If Codigo > 0 then 
			strSQL.Append(" and DE04_ID_PERIODO = " & Codigo)
		End If
		
		If Periodo > 0 then 
			strSQL.Append(" and DE04_ID_PERIODO_1 = " & Periodo)
		End If
		
		If TipoPeriodo <> "" then 
			strSQL.Append(" and upper(DE04_TP_PERIODO) like '%" & TipoPeriodo.toUpper & "%'")
		End If
		
		If Numero <> "" then 
			strSQL.Append(" and upper(DE04_NM_PERIODO) like '%" & Numero.toUpper & "%'")
		End If
		
		strSQL.Append(" Order By " & IIf(Sort = "", "DE04_ID_PERIODO", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

    Public Function ObterTabela(Optional Ano As Integer = 0) As DataTable
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

		strSQL.Append(" select rh88_ID_PERIODO as CODIGO, RH88_NM_PERIODO as DESCRICAO")
		strSQL.Append(" from  rh88_PERIODO")
		strSQL.Append(" where rh88_ID_PERIODO is not null ")

		'If Ano > 0 Then
		'	strSQL.Append(" and LEFT(RH88_NM_PERIODO,4) = '" & Ano & "' ")
		'Else
		'          strSQL.Append(" and DE04_TP_PERIODO = 1 ")
		'      End If

		strSQL.Append(" order by 2 desc ")

        dt = cnn.AbrirDataTable(strSQL.ToString)

        
        cnn = Nothing

        Return dt
    End Function

    Public Function ObterUltimo() as Integer
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		Dim CodigoUltimo As Integer
		
		strSQL.Append(" select max(DE04_ID_PERIODO) from DE04_PERIODO")

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
    Public Function Excluir(ByVal Codigo As Integer) As Integer
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim LinhasAfetadas As Integer

        strSQL.Append(" delete ")
        strSQL.Append(" from DE04_PERIODO")
        strSQL.Append(" where DE04_ID_PERIODO = " & Codigo)

        LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

        
        cnn = Nothing

        Return LinhasAfetadas
    End Function

End Class

'******************************************************************************
'*                                 25/06/2018                                 *
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

