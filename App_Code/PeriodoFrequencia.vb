Imports Microsoft.VisualBasic
Imports System.Data

Public Class PeriodoFrequencia

    Implements IDisposable

	Private RH17_ID_PERIODO_FREQ as Integer
	Private RH17_NR_ANO as Integer
	Private RH17_NR_MES as Integer
	Private RH17_DT_INICIO_LANCAMENTO as String
	Private RH17_DT_TERMINO_LANCAMENTO as String
	Private RH17_DT_LIMITE_LANCAMENTO as String
    Private RH17_ST_PERIODO_FREQ As String
    Private RH17_DH_ST_PERIODO_FREQ As String

    Public Property Codigo() As Integer
        Get
            Return RH17_ID_PERIODO_FREQ
        End Get
        Set(ByVal Value As Integer)
            RH17_ID_PERIODO_FREQ = Value
        End Set
    End Property
    Public Property Ano() As Integer
        Get
            Return RH17_NR_ANO
        End Get
        Set(ByVal Value As Integer)
            RH17_NR_ANO = Value
        End Set
    End Property
    Public Property Mes() As Integer
        Get
            Return RH17_NR_MES
        End Get
        Set(ByVal Value As Integer)
            RH17_NR_MES = Value
        End Set
    End Property
    Public Property DataInicioLancamento() As String
        Get
            Return RH17_DT_INICIO_LANCAMENTO
        End Get
        Set(ByVal Value As String)
            RH17_DT_INICIO_LANCAMENTO = Value
        End Set
    End Property
    Public Property DataTerminoLancamento() As String
        Get
            Return RH17_DT_TERMINO_LANCAMENTO
        End Get
        Set(ByVal Value As String)
            RH17_DT_TERMINO_LANCAMENTO = Value
        End Set
    End Property
    Public Property DataLimiteLancamento() As String
        Get
            Return RH17_DT_LIMITE_LANCAMENTO
        End Get
        Set(ByVal Value As String)
            RH17_DT_LIMITE_LANCAMENTO = Value
        End Set
    End Property
    Public Property SituacaoPeriodo() As String
        Get
            Return RH17_ST_PERIODO_FREQ
        End Get
        Set(ByVal Value As String)
            RH17_ST_PERIODO_FREQ = Value
        End Set
    End Property
    Public Property DataHoraSituacao() As String
        Get
            Return RH17_DH_ST_PERIODO_FREQ
        End Get
        Set(ByVal Value As String)
            RH17_DH_ST_PERIODO_FREQ = Value
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
        strSQL.Append(" from RH17_PERIODO_FREQ")
        strSQL.Append(" where RH17_ID_PERIODO_FREQ = " & Codigo)

        dt = cnn.EditarDataTable(strSQL.ToString)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
        Else
            dr = dt.Rows(0)
        End If

        dr("RH17_NR_ANO") = ProBanco(RH17_NR_ANO, eTipoValor.CHAVE)
        dr("RH17_NR_MES") = ProBanco(RH17_NR_MES, eTipoValor.CHAVE)
        dr("RH17_DT_INICIO_LANCAMENTO") = ProBanco(RH17_DT_INICIO_LANCAMENTO, eTipoValor.TEXTO)
        dr("RH17_DT_TERMINO_LANCAMENTO") = ProBanco(RH17_DT_TERMINO_LANCAMENTO, eTipoValor.TEXTO)
        dr("RH17_DT_LIMITE_LANCAMENTO") = ProBanco(RH17_DT_LIMITE_LANCAMENTO, eTipoValor.TEXTO)
        dr("RH17_ST_PERIODO_FREQ") = ProBanco(RH17_ST_PERIODO_FREQ, eTipoValor.TEXTO)
        dr("RH17_DH_ST_PERIODO_FREQ") = ProBanco(RH17_DH_ST_PERIODO_FREQ, eTipoValor.DATA)

        cnn.SalvarDataTable(dr)

        dt.Dispose()
        dt = Nothing

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Sub Obter(ByVal Codigo As String)
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH17_PERIODO_FREQ")
        strSQL.Append(" where RH17_ID_PERIODO_FREQ = " & Codigo)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            RH17_ID_PERIODO_FREQ = DoBanco(dr("RH17_ID_PERIODO_FREQ"), eTipoValor.CHAVE)
            RH17_NR_ANO = DoBanco(dr("RH17_NR_ANO"), eTipoValor.CHAVE)
            RH17_NR_MES = DoBanco(dr("RH17_NR_MES"), eTipoValor.CHAVE)
            RH17_DT_INICIO_LANCAMENTO = DoBanco(dr("RH17_DT_INICIO_LANCAMENTO"), eTipoValor.TEXTO)
            RH17_DT_TERMINO_LANCAMENTO = DoBanco(dr("RH17_DT_TERMINO_LANCAMENTO"), eTipoValor.TEXTO)
            RH17_DT_LIMITE_LANCAMENTO = DoBanco(dr("RH17_DT_LIMITE_LANCAMENTO"), eTipoValor.TEXTO)
            RH17_ST_PERIODO_FREQ = DoBanco(dr("RH17_ST_PERIODO_FREQ"), eTipoValor.TEXTO)
            RH17_DH_ST_PERIODO_FREQ = DoBanco(dr("RH17_DH_ST_PERIODO_FREQ"), eTipoValor.DATA)
        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub
    Public Sub ObterPeriodoAtivo()
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH17_PERIODO_FREQ")
        strSQL.Append(" where RH17_ST_PERIODO_FREQ = 'A'")

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            RH17_ID_PERIODO_FREQ = DoBanco(dr("RH17_ID_PERIODO_FREQ"), eTipoValor.CHAVE)
            RH17_NR_ANO = DoBanco(dr("RH17_NR_ANO"), eTipoValor.CHAVE)
            RH17_NR_MES = DoBanco(dr("RH17_NR_MES"), eTipoValor.CHAVE)
            RH17_DT_INICIO_LANCAMENTO = DoBanco(dr("RH17_DT_INICIO_LANCAMENTO"), eTipoValor.TEXTO)
            RH17_DT_TERMINO_LANCAMENTO = DoBanco(dr("RH17_DT_TERMINO_LANCAMENTO"), eTipoValor.TEXTO)
            RH17_DT_LIMITE_LANCAMENTO = DoBanco(dr("RH17_DT_LIMITE_LANCAMENTO"), eTipoValor.TEXTO)
            RH17_ST_PERIODO_FREQ = DoBanco(dr("RH17_ST_PERIODO_FREQ"), eTipoValor.TEXTO)
            RH17_DH_ST_PERIODO_FREQ = DoBanco(dr("RH17_DH_ST_PERIODO_FREQ"), eTipoValor.DATA)
        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Function Pesquisar(Optional ByVal Sort As String = "", Optional Codigo As Integer = 0, Optional Ano As Integer = 0, Optional Mes As Integer = 0, Optional DataInicioLancamento As String = "", Optional DataTerminoLancamento As String = "", Optional DataLimiteLancamento As String = "", Optional SituacaoPeriodo As String = "", Optional DataHoraSituacao As String = "") As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" select *,convert(varchar,rh17_nr_mes)+'/'+convert(varchar,RH17_NR_ANO) as PERIODO, convert(varchar,RH17_DT_INICIO_LANCAMENTO,103) + ' ' + convert(varchar,RH17_DT_TERMINO_LANCAMENTO,103) AS VINGENCIA,case RH17_ST_PERIODO_FREQ when 'A' then 'ABERTO' ELSE 'FECHADO' END SITUACAO ")
        strSQL.Append(" from RH17_PERIODO_FREQ")
        'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
        strSQL.Append(" where RH17_ID_PERIODO_FREQ is not null")

        If Codigo > 0 Then
            strSQL.Append(" and RH17_ID_PERIODO_FREQ = " & Codigo)
        End If

        If Ano > 0 Then
            strSQL.Append(" and RH17_NR_ANO = " & Ano)
        End If

        If Mes > 0 Then
            strSQL.Append(" and RH17_NR_MES = " & Mes)
        End If

        If DataInicioLancamento <> "" Then
            strSQL.Append(" and upper(RH17_DT_INICIO_LANCAMENTO) like '%" & DataInicioLancamento.ToUpper & "%'")
        End If

        If DataTerminoLancamento <> "" Then
            strSQL.Append(" and upper(RH17_DT_TERMINO_LANCAMENTO) like '%" & DataTerminoLancamento.ToUpper & "%'")
        End If

        If DataLimiteLancamento <> "" Then
            strSQL.Append(" and upper(RH17_DT_LIMITE_LANCAMENTO) like '%" & DataLimiteLancamento.ToUpper & "%'")
        End If

        If SituacaoPeriodo <> "" Then
            strSQL.Append(" and upper(RH17_ST_PERIODO_FREQ) like '%" & SituacaoPeriodo.ToUpper & "%'")
        End If
		
		If isDate(DataHoraSituacao) then 
			strSQL.Append(" and RH49_DH_ST_PERIODO = Convert(DateTime, '" & DataHoraSituacao & "', 103)")
		End If
		
		strSQL.Append(" Order By " & IIf(Sort = "", "RH17_ID_PERIODO_FREQ", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

	Public Function ObterTabela() as DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder

        strSQL.Append(" select RH17_ID_PERIODO_FREQ as CODIGO, convert(varchar,RH17_NR_MES) + '/'+ convert(varchar,RH17_NR_ANO) as DESCRICAO")
        strSQL.Append(" from RH17_PERIODO_FREQ")
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
		
		strSQL.Append(" select max(RH17_ID_PERIODO_FREQ) from RH17_PERIODO_FREQ")

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
		strSQL.Append(" from RH17_PERIODO_FREQ")
		strSQL.Append(" where RH17_ID_PERIODO_FREQ = " & Codigo)

		LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

		cnn.FecharBanco()
		cnn = Nothing

		Return LinhasAfetadas
	End Function

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        ' TODO: uncomment the following line if Finalize() is overridden above.
         GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class

'******************************************************************************
'*                                 24/05/2019                                 *
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

