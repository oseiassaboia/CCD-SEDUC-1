Imports Microsoft.VisualBasic
Imports System.Data

Public Class HistSolicCorrecaoDoc
    Private RH101_ID_HIST_SOLIC_CORREC_DOC As Integer
    Private RH100_ID_SOLIC_CORREC_DOC As Integer
    Private CA04_ID_USUARIO_ALT As Integer
    Private RH100_ST_SOLIC_CORREC_DOC As Integer
    Private RH100_DH_ST_SOLIC_CORREC_DOC As String
    Private RH100_DS_OBSERVACAO As String

    Public Property IdHistSolicCorrecaoDoc() As Integer
        Get
            Return RH101_ID_HIST_SOLIC_CORREC_DOC
        End Get
        Set(ByVal value As Integer)
            RH101_ID_HIST_SOLIC_CORREC_DOC = value
        End Set
    End Property

    Public Property IdSolicCorrecaoDoc() As Integer
        Get
            Return RH100_ID_SOLIC_CORREC_DOC
        End Get
        Set(ByVal value As Integer)
            RH100_ID_SOLIC_CORREC_DOC = value
        End Set
    End Property

    Public Property UsuarioAlt() As Integer
        Get
            Return CA04_ID_USUARIO_ALT
        End Get
        Set(ByVal value As Integer)
            CA04_ID_USUARIO_ALT = value
        End Set
    End Property

    Public Property SituacaoSolicCorrecaoDoc() As Integer
        Get
            Return RH100_ST_SOLIC_CORREC_DOC
        End Get
        Set(ByVal value As Integer)
            RH100_ST_SOLIC_CORREC_DOC = value
        End Set
    End Property

    Public Property DataHoraSituacao() As String
        Get
            Return RH100_DH_ST_SOLIC_CORREC_DOC
        End Get
        Set(ByVal value As String)
            RH100_DH_ST_SOLIC_CORREC_DOC = value
        End Set
    End Property

    Public Property Obsevacao() As String
        Get
            Return RH100_DS_OBSERVACAO
        End Get
        Set(ByVal value As String)
            RH100_DS_OBSERVACAO = value
        End Set
    End Property

    Public Sub New(Optional ByVal codigo As Integer = 0)
        If codigo > 0 Then
            Obter(codigo)
        End If
    End Sub

    Public Sub Salvar(Optional ByRef transacao As Transacao = Nothing)
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" SELECT * ")
        strSQL.Append("   FROM RH101_HIST_SOLIC_CORREC_DOC")
        strSQL.Append("  WHERE RH101_ID_HIST_SOLIC_CORREC_DOC = " & IdHistSolicCorrecaoDoc)

        dt = cnn.EditarDataTable(strSQL.ToString, transacao)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
        Else
            dr = dt.Rows(0)
        End If

        dr("RH101_ID_HIST_SOLIC_CORREC_DOC") = ProBanco(RH101_ID_HIST_SOLIC_CORREC_DOC, eTipoValor.CHAVE)
        dr("RH100_ID_SOLIC_CORREC_DOC") = ProBanco(RH100_ID_SOLIC_CORREC_DOC, eTipoValor.CHAVE)
        dr("CA04_ID_USUARIO_ALT") = ProBanco(CA04_ID_USUARIO_ALT, eTipoValor.CHAVE)
        dr("RH100_ST_SOLIC_CORREC_DOC") = ProBanco(RH100_ST_SOLIC_CORREC_DOC, eTipoValor.NUMERO_INTEIRO)
        dr("RH100_DH_ST_SOLIC_CORREC_DOC") = ProBanco(RH100_DH_ST_SOLIC_CORREC_DOC, eTipoValor.DATA_COMPLETA)
        dr("RH100_DS_OBSERVACAO") = ProBanco(RH100_DS_OBSERVACAO, eTipoValor.TEXTO_LIVRE)

        cnn.SalvarDataTable(dr)

        dt.Dispose()
        dt = Nothing

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Sub Obter(ByVal HistSolicCadastroDoc As String)
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" SELECT * ")
        strSQL.Append("   FROM RH101_HIST_SOLIC_CORREC_DOC")
        strSQL.Append("  WHERE RH101_ID_HIST_SOLIC_CORREC_DOC = " & HistSolicCadastroDoc)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            RH101_ID_HIST_SOLIC_CORREC_DOC = DoBanco(dr("RH101_ID_HIST_SOLIC_CORREC_DOC"), eTipoValor.CHAVE)
            RH100_ID_SOLIC_CORREC_DOC = DoBanco(dr("RH100_ID_SOLIC_CORREC_DOC"), eTipoValor.CHAVE)
            CA04_ID_USUARIO_ALT = DoBanco(dr("CA04_ID_USUARIO_ALT"), eTipoValor.CHAVE)
            RH100_ST_SOLIC_CORREC_DOC = DoBanco(dr("RH100_ST_SOLIC_CORREC_DOC"), eTipoValor.NUMERO_INTEIRO)
            RH100_DH_ST_SOLIC_CORREC_DOC = DoBanco(dr("RH100_DH_ST_SOLIC_CORREC_DOC"), eTipoValor.DATA_COMPLETA)
            RH100_DS_OBSERVACAO = DoBanco(dr("RH100_DS_OBSERVACAO"), eTipoValor.TEXTO_LIVRE)
        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Function Pesquisar(Optional ByVal Sort As String = "",
                              Optional SolicCorrecaoDoc As Integer = 0,
                              Optional UsuarioAlt As Integer = 0,
                              Optional SituacaoSolicCorrecaoDoc As Integer = 0,
                              Optional DataHoraSitSolicCorrecaoDoc As String = "",
                              Optional Observacao As String = "") As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" SELECT * ")
        strSQL.Append("   FROM RH101_HIST_SOLIC_CORREC_DOC")
        strSQL.Append("  WHERE RH101_ID_HIST_SOLIC_CORREC_DOC is not null")

        If SolicCorrecaoDoc > 0 Then
            strSQL.Append(" AND RH101_ID_HIST_SOLIC_CORREC_DOC = " & SolicCorrecaoDoc)
        End If

        If UsuarioAlt > 0 Then
            strSQL.Append(" AND CA04_ID_USUARIO_ALT = " & UsuarioAlt)
        End If

        If SituacaoSolicCorrecaoDoc > 0 Then
            strSQL.Append(" AND RH100_ST_SOLIC_CORREC_DOC = " & SituacaoSolicCorrecaoDoc)
        End If

        If DataHoraSitSolicCorrecaoDoc <> "" Then
            strSQL.Append(" AND RH100_DH_ST_SOLIC_CORREC_DOC = " & DataHoraSitSolicCorrecaoDoc)
        End If

        If Observacao <> "" Then
            strSQL.Append(" and upper(RH100_DS_OBSERVACAO) like '%" & Observacao.ToUpper & "%'")
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "RH101_ID_HIST_SOLIC_CORREC_DOC", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function ObterTabela() As DataTable
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" SELECT RH101_ID_HIST_SOLIC_CORREC_DOC as CODIGO, RH101_ID_SOLIC_CORREC_DOC as DESCRICAO")
        strSQL.Append("   FROM RH101_HIST_SOLIC_CORREC_DOC")
        strSQL.Append("  ORDER BY ")

        dt = cnn.AbrirDataTable(strSQL.ToString)

        cnn.FecharBanco()
        cnn = Nothing

        Return dt
    End Function

    Public Function ObterUltimo() As Integer
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim CodigoUltimo As Integer

        strSQL.Append(" SELECT MAX(RH101_ID_HIST_SOLIC_CORREC_DOC) ")
        strSQL.Append("   FROM RH101_HIST_SOLIC_CORREC_DOC")

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

    Public Function ObterDescricao(ByVal solic As Integer) As String
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim Descricao As String

        strSQL.Append(" SELECT RH100_DS_OBSERVACAO " & vbCrLf)
        strSQL.Append("   FROM RH101_HIST_SOLIC_CORREC_DOC  " & vbCrLf)
        strSQL.Append("  WHERE RH101_ID_HIST_SOLIC_CORREC_DOC = (SELECT MAX(RH101_ID_HIST_SOLIC_CORREC_DOC) -1 FROM RH101_HIST_SOLIC_CORREC_DOC ) " & vbCrLf)
        strSQL.Append("    AND RH100_ID_SOLIC_CORREC_DOC = " & solic & vbCrLf)

        With cnn.AbrirDataTable(strSQL.ToString)
            If Not IsDBNull(.Rows(0)(0)) Then
                Descricao = .Rows(0)(0)
            Else
                Descricao = "-"
            End If
        End With

        cnn.FecharBanco()
        cnn = Nothing

        Return Descricao

    End Function

    Public Function Excluir(ByVal SolicCorrecaoDoc As String) As Integer
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim LinhasAfetadas As Integer

        strSQL.Append(" DELECT ")
        strSQL.Append("   FROM RH101_HIST_SOLIC_CORREC_DOC")
        strSQL.Append("  WHERE RH101_ID_HIST_SOLIC_CORREC_DOC = " & SolicCorrecaoDoc)

        LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

        cnn.FecharBanco()
        cnn = Nothing

        Return LinhasAfetadas
    End Function
End Class
