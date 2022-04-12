Imports Microsoft.VisualBasic
Imports System.Data

Public Class HistSolicitarCadastroDoc
    Private RH103_ID_HIST_SOLIC_CAD_DOC As Integer
    Private RH102_ID_SOLIC_CAD_DOC As Integer
    Private CA04_ID_USUARIO_ALT As Integer
    Private RH102_ST_SOLIC_CAD_DOC As Integer
    Private RH102_DH_ST_SOLIC_CAD_DOC As String
    Private RH102_DS_OBSERVACAO As String

    Public Property IdHistSolicCadastroDoc() As Integer
        Get
            Return RH103_ID_HIST_SOLIC_CAD_DOC
        End Get
        Set(ByVal value As Integer)
            RH103_ID_HIST_SOLIC_CAD_DOC = value
        End Set
    End Property

    Public Property SolicCadastroDoc() As Integer
        Get
            Return RH102_ID_SOLIC_CAD_DOC
        End Get
        Set(ByVal value As Integer)
            RH102_ID_SOLIC_CAD_DOC = value
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

    Public Property SituacaoSolicCadastroDoc() As Integer
        Get
            Return RH102_ST_SOLIC_CAD_DOC
        End Get
        Set(ByVal value As Integer)
            RH102_ST_SOLIC_CAD_DOC = value
        End Set
    End Property

    Public Property DataHoraSituacao() As String
        Get
            Return RH102_DH_ST_SOLIC_CAD_DOC
        End Get
        Set(ByVal value As String)
            RH102_DH_ST_SOLIC_CAD_DOC = value
        End Set
    End Property

    Public Property Obsevacao() As String
        Get
            Return RH102_DS_OBSERVACAO
        End Get
        Set(ByVal value As String)
            RH102_DS_OBSERVACAO = value
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
        strSQL.Append("   FROM RH103_HIST_SOLIC_CAD_DOC")
        strSQL.Append("  WHERE RH103_ID_HIST_SOLIC_CAD_DOC = " & IdHistSolicCadastroDoc)

        dt = cnn.EditarDataTable(strSQL.ToString, transacao)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
        Else
            dr = dt.Rows(0)
        End If

        dr("RH103_ID_HIST_SOLIC_CAD_DOC") = ProBanco(RH103_ID_HIST_SOLIC_CAD_DOC, eTipoValor.CHAVE)
        dr("RH102_ID_SOLIC_CAD_DOC") = ProBanco(RH102_ID_SOLIC_CAD_DOC, eTipoValor.CHAVE)
        dr("CA04_ID_USUARIO_ALT") = ProBanco(CA04_ID_USUARIO_ALT, eTipoValor.CHAVE)
        dr("RH102_ST_SOLIC_CAD_DOC") = ProBanco(RH102_ST_SOLIC_CAD_DOC, eTipoValor.NUMERO_INTEIRO)
        dr("RH102_DH_ST_SOLIC_CAD_DOC") = ProBanco(RH102_DH_ST_SOLIC_CAD_DOC, eTipoValor.DATA_COMPLETA)
        dr("RH102_DS_OBSERVACAO") = ProBanco(RH102_DS_OBSERVACAO, eTipoValor.TEXTO_LIVRE)

        cnn.SalvarDataTable(dr)

        dt.Dispose()
        dt = Nothing

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Sub Obter(ByVal SolicCadastro As String)
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" SELECT * ")
        strSQL.Append("   FROM RH103_ID_HIST_SOLIC_CAD_DOC")
        strSQL.Append("  WHERE RH103_HIST_SOLIC_CAD_DOC = " & SolicCadastro)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            RH103_ID_HIST_SOLIC_CAD_DOC = DoBanco(dr("RH103_ID_HIST_SOLIC_CAD_DOC"), eTipoValor.CHAVE)
            RH102_ID_SOLIC_CAD_DOC = DoBanco(dr("RH102_ID_SOLIC_CAD_DOC"), eTipoValor.CHAVE)
            CA04_ID_USUARIO_ALT = DoBanco(dr("CA04_ID_USUARIO_ALT"), eTipoValor.CHAVE)
            RH102_ST_SOLIC_CAD_DOC = DoBanco(dr("RH102_ST_SOLIC_CAD_DOC"), eTipoValor.NUMERO_INTEIRO)
            RH102_DH_ST_SOLIC_CAD_DOC = DoBanco(dr("RH102_DH_ST_SOLIC_CAD_DOC"), eTipoValor.DATA_COMPLETA)
            RH102_DS_OBSERVACAO = DoBanco(dr("RH102_DS_OBSERVACAO"), eTipoValor.TEXTO_LIVRE)
        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Function Pesquisar(Optional ByVal Sort As String = "",
                              Optional HistSolicCadastro As Integer = 0,
                              Optional SolicCadastro As Integer = 0,
                              Optional UsuarioAlt As Integer = 0,
                              Optional SituacaoSolicCadDoc As Integer = 0,
                              Optional DataHoraSituacaoSolicCadDoc As String = "",
                              Optional Observacao As String = "") As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" SELECT * ")
        strSQL.Append("   FROM RH103_HIST_SOLIC_CAD_DOC")
        strSQL.Append("  WHERE RH103_ID_HIST_SOLIC_CAD_DOC is not null")

        If HistSolicCadastro > 0 Then
            strSQL.Append(" AND RH103_ID_HIST_SOLIC_CAD_DOC = " & HistSolicCadastro)
        End If

        If SolicCadastro > 0 Then
            strSQL.Append(" AND RH102_ID_SOLIC_CAD_DOC = " & SolicCadastro)
        End If

        If UsuarioAlt > 0 Then
            strSQL.Append(" AND CA04_ID_USUARIO_ALT = " & UsuarioAlt)
        End If

        If SituacaoSolicCadDoc > 0 Then
            strSQL.Append(" AND RH102_ST_SOLIC_CAD_DOC = " & SituacaoSolicCadDoc)
        End If

        If DataHoraSituacaoSolicCadDoc <> "" Then
            strSQL.Append(" AND RH102_DH_ST_SOLIC_CAD_DOC = " & DataHoraSituacaoSolicCadDoc)
        End If

        If DataHoraSituacaoSolicCadDoc <> "" Then
            strSQL.Append(" AND CI02_DT_NASCIMENTO_ALUNO = " & DataHoraSituacaoSolicCadDoc)
        End If

        If Observacao <> "" Then
            strSQL.Append(" AND UPPER(RH102_DS_OBSERVACAO) LIKE '%" & Observacao.ToUpper & "%'")
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "RH103_ID_HIST_SOLIC_CAD_DOC", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function ObterTabela() As DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder

        strSQL.Append(" SELECT RH103_ID_HIST_SOLIC_CAD_DOC as CODIGO, RH102_DS_OBSERVACAO as DESCRICAO")
        strSQL.Append("   FROM RH103_HIST_SOLIC_CAD_DOC")
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

        strSQL.Append(" SELECT max(RH103_ID_HIST_SOLIC_CAD_DOC) ")
        strSQL.Append("   FROM RH103_HIST_SOLIC_CAD_DOC")

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
        Dim valido As Boolean = False
        Dim Descricao As String

        strSQL.Append(" SELECT RH102_DS_OBSERVACAO " & vbCrLf)
        strSQL.Append("   FROM RH103_HIST_SOLIC_CAD_DOC  " & vbCrLf)
        strSQL.Append("  WHERE RH102_ID_SOLIC_CAD_DOC = " & solic & vbCrLf)

        With cnn.AbrirDataTable(strSQL.ToString)
            If .Rows.Count = 1 Then
                Descricao = .Rows(0)(0)
                valido = True
            Else
                Descricao = "-"
            End If
        End With

        If valido = False Then
            strSQL.Append(" SELECT RH102_DS_OBSERVACAO " & vbCrLf)
            strSQL.Append("   FROM RH103_HIST_SOLIC_CAD_DOC  " & vbCrLf)
            strSQL.Append("  WHERE RH103_ID_HIST_SOLIC_CAD_DOC  = (SELECT MAX(RH103_ID_HIST_SOLIC_CAD_DOC) -1 FROM RH103_HIST_SOLIC_CAD_DOC ) " & vbCrLf)
            strSQL.Append("    AND RH102_ID_SOLIC_CAD_DOC = " & solic & vbCrLf)

            With cnn.AbrirDataTable(strSQL.ToString)
                If Not IsDBNull(.Rows(0)(0)) Then
                    Descricao = .Rows(0)(0)
                Else
                    Descricao = "-"
                End If
            End With
        End If

        cnn.FecharBanco()
        cnn = Nothing

        Return Descricao
    End Function

    Public Function Excluir(ByVal HistSolicCadastro As String) As Integer
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim LinhasAfetadas As Integer

        strSQL.Append(" DELETE ")
        strSQL.Append("   FROM RH103_HIST_SOLIC_CAD_DOC")
        strSQL.Append("  WHERE RH103_ID_HIST_SOLIC_CAD_DOC = " & HistSolicCadastro)

        LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

        cnn.FecharBanco()
        cnn = Nothing

        Return LinhasAfetadas
    End Function

End Class
