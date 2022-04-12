Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Text
Public Class Relatorio
    Dim strAmbienteTeste As String = System.Configuration.ConfigurationManager.AppSettings("AmbienteTeste").ToString

    Private TG20_ID_RELATORIO As Integer
    Private CA01_ID_APLICACAO As Integer
    Private TG20_CD_RELATORIO As String
    Private TG20_NM_RELATORIO As String
    Private TG20_NM_ARQ_PROGRAMA_PDF As String
    Private TG20_NM_ARQ_PROGRAMA_EXCEL As String
    Private TG20_DS_PARAMETRO As String
    Private TG20_DS_FILTRO As String
    Private TG20_DS_CONDICAO As String
    Private TG20_DS_CONDICAO_OBRIGATORIA As String
    Private TG20_IN_ATIVO As String
    Private TG20_DT_IMPLEMENTACAO As String
    Private TG20_NR_AREA_RELATORIO As Integer

    Public Property Codigo() As Integer
        Get
            Return TG20_ID_RELATORIO
        End Get
        Set(ByVal Value As Integer)
            TG20_ID_RELATORIO = Value
        End Set
    End Property
    Public Property Aplicacao() As Integer
        Get
            Return CA01_ID_APLICACAO
        End Get
        Set(ByVal Value As Integer)
            CA01_ID_APLICACAO = Value
        End Set
    End Property

    Public Property Numeracao() As String
        Get
            Return TG20_CD_RELATORIO
        End Get
        Set(ByVal Value As String)
            TG20_CD_RELATORIO = Value
        End Set
    End Property
    Public Property Nome() As String
        Get
            Return TG20_NM_RELATORIO
        End Get
        Set(ByVal Value As String)
            TG20_NM_RELATORIO = Value
        End Set
    End Property

    Public Property ArquivoProgramaPdf() As String
        Get
            Return TG20_NM_ARQ_PROGRAMA_PDF
        End Get
        Set(ByVal Value As String)
            TG20_NM_ARQ_PROGRAMA_PDF = Value
        End Set
    End Property

    Public Property ArquivoProgramaExcel() As String
        Get
            Return TG20_NM_ARQ_PROGRAMA_EXCEL
        End Get
        Set(ByVal Value As String)
            TG20_NM_ARQ_PROGRAMA_EXCEL = Value
        End Set
    End Property

    Public Property Parametro() As String
        Get
            Return TG20_DS_PARAMETRO
        End Get
        Set(ByVal Value As String)
            TG20_DS_PARAMETRO = Value
        End Set
    End Property

    Public Property Filtro() As String
        Get
            Return TG20_DS_FILTRO
        End Get
        Set(ByVal Value As String)
            TG20_DS_FILTRO = Value
        End Set
    End Property

    Public Property DescricaoCondicao() As String
        Get
            Return TG20_DS_CONDICAO
        End Get
        Set(ByVal Value As String)
            TG20_DS_CONDICAO = Value
        End Set
    End Property

    Public Property CondicaoObrigatoria() As String
        Get
            Return TG20_DS_CONDICAO_OBRIGATORIA
        End Get
        Set(ByVal Value As String)
            TG20_DS_CONDICAO_OBRIGATORIA = Value
        End Set
    End Property
    Public Property Ativo() As String
        Get
            Return TG20_IN_ATIVO
        End Get
        Set(ByVal Value As String)
            TG20_IN_ATIVO = Value
        End Set
    End Property

    Public Property DataImplementacao() As String
        Get
            Return TG20_DT_IMPLEMENTACAO
        End Get
        Set(ByVal Value As String)
            TG20_DT_IMPLEMENTACAO = Value
        End Set
    End Property

    Public Property Area() As Integer
        Get
            Return TG20_NR_AREA_RELATORIO
        End Get
        Set(ByVal Value As Integer)
            TG20_NR_AREA_RELATORIO = Value
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
        strSQL.Append(" from DBGERAL.DBO.TG20_RELATORIO")
        strSQL.Append(" where TG20_ID_RELATORIO = " & Codigo)

        dt = cnn.EditarDataTable(strSQL.ToString)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
        Else
            dr = dt.Rows(0)
        End If

        dr("CA01_ID_APLICACAO") = ProBanco(CA01_ID_APLICACAO, eTipoValor.CHAVE)
        dr("TG20_CD_RELATORIO") = ProBanco(TG20_CD_RELATORIO, eTipoValor.TEXTO)
        dr("TG20_NM_RELATORIO") = ProBanco(TG20_NM_RELATORIO, eTipoValor.TEXTO)
        dr("TG20_NM_ARQ_PROGRAMA_PDF") = ProBanco(TG20_NM_ARQ_PROGRAMA_PDF, eTipoValor.TEXTO)
        dr("TG20_NM_ARQ_PROGRAMA_EXCEL") = ProBanco(TG20_NM_ARQ_PROGRAMA_EXCEL, eTipoValor.TEXTO)
        dr("TG20_DS_PARAMETRO") = ProBanco(TG20_DS_PARAMETRO, eTipoValor.TEXTO)
        dr("TG20_DS_FILTRO") = ProBanco(TG20_DS_FILTRO, eTipoValor.TEXTO)
        dr("TG20_DS_CONDICAO") = ProBanco(TG20_DS_CONDICAO, eTipoValor.TEXTO)
        dr("TG20_DS_CONDICAO_OBRIGATORIA") = ProBanco(TG20_DS_CONDICAO_OBRIGATORIA, eTipoValor.TEXTO)
        dr("TG20_IN_ATIVO") = ProBanco(TG20_IN_ATIVO, eTipoValor.BOOLEANO)
        dr("TG20_DT_IMPLEMENTACAO") = ProBanco(TG20_DT_IMPLEMENTACAO, eTipoValor.DATA)
        dr("TG20_NR_AREA_RELATORIO") = ProBanco(TG20_NR_AREA_RELATORIO, eTipoValor.NUMERO_INTEIRO)

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
        strSQL.Append(" from DBGERAL.DBO.TG20_RELATORIO")
        strSQL.Append(" where TG20_ID_RELATORIO = " & Codigo)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            TG20_ID_RELATORIO = DoBanco(dr("TG20_ID_RELATORIO"), eTipoValor.CHAVE)
            CA01_ID_APLICACAO = DoBanco(dr("CA01_ID_APLICACAO"), eTipoValor.CHAVE)
            TG20_CD_RELATORIO = DoBanco(dr("TG20_CD_RELATORIO"), eTipoValor.TEXTO)
            TG20_NM_RELATORIO = DoBanco(dr("TG20_NM_RELATORIO"), eTipoValor.TEXTO)
            TG20_NM_ARQ_PROGRAMA_PDF = DoBanco(dr("TG20_NM_ARQ_PROGRAMA_PDF"), eTipoValor.TEXTO)
            TG20_NM_ARQ_PROGRAMA_EXCEL = DoBanco(dr("TG20_NM_ARQ_PROGRAMA_EXCEL"), eTipoValor.TEXTO)
            TG20_DS_PARAMETRO = DoBanco(dr("TG20_DS_PARAMETRO"), eTipoValor.TEXTO_LIVRE)
            TG20_DS_FILTRO = DoBanco(dr("TG20_DS_FILTRO"), eTipoValor.TEXTO_LIVRE)
            TG20_DS_CONDICAO = DoBanco(dr("TG20_DS_CONDICAO"), eTipoValor.TEXTO_LIVRE)
            TG20_DS_CONDICAO_OBRIGATORIA = DoBanco(dr("TG20_DS_CONDICAO_OBRIGATORIA"), eTipoValor.TEXTO_LIVRE)
            TG20_IN_ATIVO = DoBanco(dr("TG20_IN_ATIVO"), eTipoValor.BOOLEANO)
            TG20_DT_IMPLEMENTACAO = DoBanco(dr("TG20_DT_IMPLEMENTACAO"), eTipoValor.DATA)
            TG20_NR_AREA_RELATORIO = DoBanco(dr("TG20_NR_AREA_RELATORIO"), eTipoValor.NUMERO_INTEIRO)
        End If


        cnn = Nothing
    End Sub

    Public Function Pesquisar(Optional ByVal Sort As String = "",
                              Optional Codigo As Integer = 0,
                              Optional Aplicacao As Integer = 0,
                              Optional Nome As String = "") As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" select TG20.TG20_ID_RELATORIO, TG20.TG20_NM_RELATORIO, TG20.TG20_NM_ARQ_PROGRAMA_PDF, TG20.TG20_NM_ARQ_PROGRAMA_EXCEL, TG20.TG20_DS_PARAMETRO, TG20.TG20_DS_FILTRO, TG20.TG20_DS_CONDICAO, TG20.TG20_DS_CONDICAO_OBRIGATORIA ")
        strSQL.Append(" , CA01.CA01_COD_APLICACAO, CA01.CA01_DES_APLICACAO, TG20.TG20_IN_ATIVO, IIF(TG20_IN_ATIVO = 1, 'SIM', 'NAO') AS ATIVO, TG20.TG20_CD_RELATORIO ")
        strSQL.Append(" , iif(len(TG20.TG20_NM_ARQ_PROGRAMA_EXCEL) > 30, SUBSTRING(TG20.TG20_NM_ARQ_PROGRAMA_EXCEL, 0, 30), TG20.TG20_NM_ARQ_PROGRAMA_EXCEL) AS NM_PROGRAMA_EXCEL")
        strSQL.Append(" , iif(len(TG20.TG20_NM_ARQ_PROGRAMA_PDF) > 30, SUBSTRING(TG20.TG20_NM_ARQ_PROGRAMA_PDF, 0, 30), TG20.TG20_NM_ARQ_PROGRAMA_PDF) AS NM_PROGRAMA_PDF, TG20_NR_AREA_RELATORIO ")
        strSQL.Append(" from DBGERAL.DBO.TG20_RELATORIO AS TG20")
        strSQL.Append(" left join [172.16.2.71].DBCONTROLEACESSO.DBO.CA01_APLICACAO AS CA01 on CA01.CA01_COD_APLICACAO = TG20.CA01_ID_APLICACAO ")
        strSQL.Append(" where TG20.TG20_ID_RELATORIO is not null")

        If Codigo > 0 Then
            strSQL.Append(" and TG20.TG20_ID_RELATORIO = " & Codigo)
        End If

        If Aplicacao > 0 Then
            strSQL.Append(" and CA01.CA01_COD_APLICACAO = " & Aplicacao)
        End If

        If Nome <> "" Then
            strSQL.Append(" and upper(TG20.TG20_NM_RELATORIO) like '%" & Nome.ToUpper & "%'")
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "TG20.TG20_CD_RELATORIO", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function ObterTabela(Optional ByVal aplicacao As Integer = 87, Optional ByVal area As Integer = 0) As DataTable
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" select TG20_ID_RELATORIO as CODIGO, TG20_NM_RELATORIO as DESCRICAO")
        strSQL.Append(" from DBGERAL.DBO.TG20_RELATORIO")
        strSQL.Append(" where TG20_ID_RELATORIO is not null")

        If strAmbienteTeste <> "AMBIENTE DE TESTE" Then
            strSQL.Append(" and TG20_IN_ATIVO = 1 ")
        End If
        strSQL.Append(" and CA01_ID_APLICACAO = " & aplicacao)

        strSQL.Append(" and isnull(TG20_NR_AREA_RELATORIO,0) = " & area)
        strSQL.Append(" order by 2 ")

        dt = cnn.AbrirDataTable(strSQL.ToString)


        cnn = Nothing

        Return dt
    End Function

    Public Function ObterUltimo() As Integer
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim CodigoUltimo As Integer

        strSQL.Append(" select max(TG20_ID_RELATORIO) from DBGERAL.DBO.TG20_RELATORIO")

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
        strSQL.Append(" from DBGERAL.DBO.TG20_RELATORIO")
        strSQL.Append(" where TG20_ID_RELATORIO = " & Codigo)

        LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)


        cnn = Nothing

        Return LinhasAfetadas
    End Function
End Class
