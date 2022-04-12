Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.VisualBasic

Public Class Processo
    Implements IDisposable

    Private PC01_ID_PROCESSO As Integer
    Private DI17_ID_TIPO_PROCESSO As Integer
    Private CA01_ID_APLICACAO As Integer
    Private PC01_ID_EPROCESSO As Integer
    Private PC01_NR_ANO_EPROCESSO As Integer
    Private PC01_NR_ANO_PROCESSO As Integer
    Private PC01_NR_PROCESSO As Integer
    Private RH02_ID_SERVIDOR_BENEF As Integer
    Private CA04_ID_USUARIO As Integer
    Private PC01_DH_CADASTRO As String

    Private disposedValue As Boolean


#Region "Getters e  Setters"
    Public Property Codigo() As Integer
        Get
            Return PC01_ID_PROCESSO
        End Get
        Set(value As Integer)
            PC01_ID_PROCESSO = value
        End Set
    End Property

    Public Property TipoProcesso() As Integer
        Get
            Return DI17_ID_TIPO_PROCESSO
        End Get
        Set(value As Integer)
            DI17_ID_TIPO_PROCESSO = value
        End Set
    End Property

    Public Property CodAplicacao() As Integer
        Get
            Return CA01_ID_APLICACAO
        End Get
        Set(value As Integer)
            CA01_ID_APLICACAO = value
        End Set
    End Property

    Public Property IdEprocesso() As Integer
        Get
            Return PC01_ID_EPROCESSO
        End Get
        Set(value As Integer)
            PC01_ID_EPROCESSO = value
        End Set
    End Property

    Public Property AnoEprocesso() As Integer
        Get
            Return PC01_NR_ANO_EPROCESSO
        End Get
        Set(value As Integer)
            PC01_NR_ANO_EPROCESSO = value
        End Set
    End Property

    Public Property AnoProcesso() As Integer
        Get
            Return PC01_NR_ANO_PROCESSO
        End Get
        Set(value As Integer)
            PC01_NR_ANO_PROCESSO = value
        End Set
    End Property

    Public Property NumeroProcesso() As Integer
        Get
            Return PC01_NR_PROCESSO
        End Get
        Set(value As Integer)
            PC01_NR_PROCESSO = value
        End Set
    End Property

    Public Property DataCadastro() As String
        Get
            Return PC01_DH_CADASTRO
        End Get
        Set(value As String)
            PC01_DH_CADASTRO = value
        End Set
    End Property

    Public Property CodServidor() As Integer
        Get
            Return RH02_ID_SERVIDOR_BENEF
        End Get
        Set(value As Integer)
            RH02_ID_SERVIDOR_BENEF = value
        End Set
    End Property

    Public Property Usuario As Integer
        Get
            Return CA04_ID_USUARIO
        End Get
        Set(value As Integer)
            CA04_ID_USUARIO = value
        End Set
    End Property
#End Region

    Public Sub New(Optional ByVal IdProcesso As Integer = 0)
        If IdProcesso > 0 Then
            Obter(IdProcesso)
        End If
    End Sub

    Public Sub Salvar(Optional ByRef transacao As Transacao = Nothing, Optional cnn As Conexao = Nothing)
        If cnn Is Nothing Then
            cnn = New Conexao("StringConexaoDiarias")
        End If
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from PC01_PROCESSO")
        strSQL.Append(" where PC01_ID_PROCESSO = " & Codigo)

        dt = cnn.EditarDataTable(strSQL.ToString, transacao)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
        Else
            dr = dt.Rows(0)
        End If

        dr("PC01_ID_PROCESSO") = ProBanco(PC01_ID_PROCESSO, eTipoValor.CHAVE)
        dr("DI17_ID_TIPO_PROCESSO") = ProBanco(DI17_ID_TIPO_PROCESSO, eTipoValor.CHAVE)
        dr("CA01_ID_APLICACAO") = ProBanco(CA01_ID_APLICACAO, eTipoValor.CHAVE)
        dr("PC01_ID_EPROCESSO") = ProBanco(PC01_ID_EPROCESSO, eTipoValor.CHAVE)
        dr("PC01_NR_ANO_EPROCESSO") = ProBanco(PC01_NR_ANO_EPROCESSO, eTipoValor.NUMERO_INTEIRO)
        dr("PC01_NR_ANO_PROCESSO") = ProBanco(PC01_NR_ANO_PROCESSO, eTipoValor.NUMERO_INTEIRO)
        dr("PC01_NR_PROCESSO") = ProBanco(PC01_NR_PROCESSO, eTipoValor.NUMERO_INTEIRO)
        dr("RH02_ID_SERVIDOR_BENEF") = ProBanco(RH02_ID_SERVIDOR_BENEF, eTipoValor.CHAVE)
        dr("CA04_ID_USUARIO") = ProBanco(RH02_ID_SERVIDOR_BENEF, eTipoValor.CHAVE)
        dr("PC01_DH_CADASTRO") = ProBanco(PC01_DH_CADASTRO, eTipoValor.DATA_COMPLETA)

        cnn.SalvarDataTable(dr)

        dt.Dispose()

        cnn.FecharBanco()
    End Sub

    Public Sub Obter(idProcesso As Integer)
        Dim cnn As New Conexao("StringConexaoDiarias")
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from PC01_PROCESSO")
        strSQL.Append(" where PC01_ID_PROCESSO = " & idProcesso)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            PC01_ID_PROCESSO = DoBanco(dr("PC01_ID_PROCESSO"), eTipoValor.CHAVE)
            DI17_ID_TIPO_PROCESSO = DoBanco(dr("DI17_ID_TIPO_PROCESSO"), eTipoValor.CHAVE)
            CA01_ID_APLICACAO = DoBanco(dr("CA01_ID_APLICACAO"), eTipoValor.CHAVE)
            PC01_ID_EPROCESSO = DoBanco(dr("PC01_ID_EPROCESSO"), eTipoValor.CHAVE)
            PC01_NR_ANO_EPROCESSO = DoBanco(dr("PC01_NR_ANO_EPROCESSO"), eTipoValor.NUMERO_INTEIRO)
            PC01_NR_ANO_PROCESSO = DoBanco(dr("PC01_NR_ANO_PROCESSO"), eTipoValor.NUMERO_INTEIRO)
            PC01_NR_PROCESSO = DoBanco(dr("PC01_NR_PROCESSO"), eTipoValor.NUMERO_INTEIRO)
            RH02_ID_SERVIDOR_BENEF = DoBanco(dr("RH02_ID_SERVIDOR_BENEF"), eTipoValor.CHAVE)
            CA04_ID_USUARIO = DoBanco(dr("CA04_ID_USUARIO"), eTipoValor.CHAVE)
            PC01_DH_CADASTRO = DoBanco(dr("PC01_DH_CADASTRO"), eTipoValor.DATA_COMPLETA)

        End If

        cnn.FecharBanco()
    End Sub

    Public Function ConsultarMovimentoProcessoAtual(IdServidor As String) As DataTable
        Dim cnn As New Conexao("StringConexaoDiarias")
        Dim strSQL As New StringBuilder
        strSQL.Append(" SELECT *" & vbCrLf)
        strSQL.Append("   FROM PC01_PROCESSO                            PC01 " & vbCrLf)
        strSQL.Append("   LEFT JOIN DI17_TIPO_PROCESSO                  DI17 ON DI17.DI17_COD_TIPO_PROCESSO = PC01.DI17_ID_TIPO_PROCESSO " & vbCrLf)
        strSQL.Append("   JOIN PC02_MOVIMENTO_PROCESSO                  PC02 ON PC02.PC01_ID_PROCESSO       = PC01.PC01_ID_PROCESSO " & vbCrLf)
        strSQL.Append("   JOIN DI09_STATUS                              DI09 ON DI09.DI09_COD_STATUS        = PC02.DI09_ID_STATUS " & vbCrLf)
        strSQL.Append("   JOIN [10.31.35.4].[DBRH].[DBO].[RH36_LOTACAO] RH36 ON RH36.RH36_ID_LOTACAO        = DI09.RH36_ID_LOTACAO " & vbCrLf)
        strSQL.Append("  WHERE PC01.PC01_ID_PROCESSO IS NOT NULL" & vbCrLf)
        strSQL.Append("    AND PC01.CA01_ID_APLICACAO = 91" & vbCrLf)
        strSQL.Append("    AND PC02.PC02_IN_ATIVO = 1" & vbCrLf)
        strSQL.Append("    AND RH02_ID_SERVIDOR_BENEF = " + IdServidor & vbCrLf)
        strSQL.Append("  ORDER BY PC01.PC01_ID_PROCESSO" & vbCrLf)

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function ConsultarMovimentosProcesso(IdServidor As String) As DataTable
        Dim cnn As New Conexao("StringConexaoDiarias")
        Dim strSQL As New StringBuilder
        strSQL.Append(" SELECT DI09.DI09_RESUMO," & vbCrLf)
        strSQL.Append("        RH36.RH36_NM_LOTACAO, " & vbCrLf)
        strSQL.Append("        PC02.PC02_DH_CADASTRO " & vbCrLf)
        strSQL.Append("   FROM PC01_PROCESSO                            PC01 " & vbCrLf)
        strSQL.Append("   LEFT JOIN DI17_TIPO_PROCESSO                  DI17 ON DI17.DI17_COD_TIPO_PROCESSO = PC01.DI17_ID_TIPO_PROCESSO " & vbCrLf)
        strSQL.Append("   JOIN PC02_MOVIMENTO_PROCESSO                  PC02 ON PC02.PC01_ID_PROCESSO       = PC01.PC01_ID_PROCESSO " & vbCrLf)
        strSQL.Append("   JOIN DI09_STATUS                              DI09 ON DI09.DI09_COD_STATUS        = PC02.DI09_ID_STATUS " & vbCrLf)
        strSQL.Append("   JOIN [10.31.35.4].[DBRH].[DBO].[RH36_LOTACAO] RH36 ON RH36.RH36_ID_LOTACAO        = DI09.RH36_ID_LOTACAO " & vbCrLf)
        strSQL.Append("  WHERE PC01.PC01_ID_PROCESSO IS NOT NULL" & vbCrLf)
        strSQL.Append("    AND PC01.CA01_ID_APLICACAO = 91" & vbCrLf)
        strSQL.Append("    AND RH02_ID_SERVIDOR_BENEF = " + IdServidor & vbCrLf)
        strSQL.Append("    AND PC02.PC02_IN_ATIVO IS NOT NULL " & vbCrLf)
        strSQL.Append("  ORDER BY PC02.DI09_ID_STATUS " & vbCrLf)

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function Pesquisar(Optional ByVal Sort As String = "",
                              Optional IdProcesso As Integer = 0,
                              Optional IdTipoProcesso As Integer = 0,
                              Optional IdAplicacao As Integer = 0,
                              Optional IdEprocesso As Integer = 0,
                              Optional AnoEprocesso As Integer = 0,
                              Optional AnoProcesso As Integer = 0,
                              Optional NumeroProcesso As Integer = 0,
                              Optional IdServidorBenef As Integer = 0,
                              Optional DataHoraCadastro As String = "") As DataTable

        Dim cnn As New Conexao("StringConexaoDiarias")
        Dim strSQL As New StringBuilder

        strSQL.Append(" SELECT *")
        strSQL.Append("   FROM PC01_PROCESSO                            PC01")
        strSQL.Append("   LEFT JOIN DI17_TIPO_PROCESSO                  DI17 ON DI17.DI17_COD_TIPO_PROCESSO = PC01.DI17_ID_TIPO_PROCESSO")
        'strSQL.Append("   JOIN [10.31.35.4].[DBRH].[DBO].[RH36_LOTACAO] RH36 ON RH36.RH36_ID_LOTACAO = DI09_STATUS.RH36_ID_LOTACAO")
        strSQL.Append("  WHERE PC01_ID_PROCESSO Is Not NULL")


        If IdProcesso > 0 Then
            strSQL.Append(" And PC01.PC01_ID_PROCESSO = " & IdProcesso)
        End If

        If IdTipoProcesso > 0 Then
            strSQL.Append(" And DI17.DI17_ID_TIPO_PROCESSO = " & IdTipoProcesso)
        End If

        If IdAplicacao > 0 Then
            strSQL.Append(" And PC01.CA01_ID_APLICACAO = " & IdAplicacao)
        Else
            strSQL.Append(" And PC01.CA01_ID_APLICACAO = 91") 'PROJETOS DO SIGEP - LICENÇAS
        End If

        If IdEprocesso > 0 Then
            strSQL.Append(" And PC01.PC01_ID_EPROCESSO = " & IdEprocesso)
        End If

        If AnoEprocesso > 0 Then
            strSQL.Append(" And PC01.PC01_NR_ANO_EPROCESSO = " & AnoEprocesso)
        End If

        If AnoEprocesso > 0 Then
            strSQL.Append(" And PC01.PC01_NR_ANO_EPROCESSO = " & AnoEprocesso)
        End If

        If IdServidorBenef > 0 Then
            strSQL.Append(" And PC01.RH02_ID_SERVIDOR_BENEF = " & IdServidorBenef)
        End If

        If AnoProcesso > 0 Then
            strSQL.Append(" And PC01.PC01_NR_ANO_PROCESSO = " & AnoProcesso)
        End If

        If NumeroProcesso > 0 Then
            strSQL.Append(" And PC01.PC01_NR_PROCESSO = " & NumeroProcesso)
        End If

        If IsDate(DataHoraCadastro) Then
            strSQL.Append(" And PC01.PC01_DH_CADASTRO = Convert(DateTime, '" & DataHoraCadastro & "', 103)")
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "PC01.PC01_ID_PROCESSO", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)

    End Function

    Public Function ObterUltimo(Optional ByRef transacao As Transacao = Nothing) As Integer
        Dim cnn As New Conexao("StringConexaoDiarias")
        Dim strSQL As New StringBuilder
        Dim CodigoUltimo As Integer

        strSQL.Append(" select max(PC01_ID_PROCESSO) from PC01_PROCESSO")

        With cnn.AbrirDataTable(strSQL.ToString, transacao)
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

    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects)
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override finalizer
            ' TODO: set large fields to null
            disposedValue = True
        End If
    End Sub

    ' ' TODO: override finalizer only if 'Dispose(disposing As Boolean)' has code to free unmanaged resources
    ' Protected Overrides Sub Finalize()
    '     ' Não altere este código. Coloque o código de limpeza no método 'Dispose(disposing As Boolean)'
    '     Dispose(disposing:=False)
    '     MyBase.Finalize()
    ' End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        ' Não altere este código. Coloque o código de limpeza no método 'Dispose(disposing As Boolean)'
        Dispose(disposing:=True)
        GC.SuppressFinalize(Me)
    End Sub
End Class
