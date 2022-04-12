Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.VisualBasic

Public Class MovimentoProcesso
    Implements IDisposable

    Private PC02_ID_MOVIMENTO_PROCESSO As Integer
    Private PC02_ID_MOVIMENTO_EPROCESSO As Integer
    Private PC01_ID_PROCESSO As Integer
    'Private SV04_ID_LOTACAO As Integer VERIFICAR POSSIBILIDADE DE CRIAÇÃO DE UMA RH36_ID_LOTACAO
    Private DI09_ID_STATUS As Integer
    Private CA04_ID_USUARIO As Integer
    Private DI24_ID_COLABORADOR As Integer
    Private PC02_IN_PROCESSADO As Integer
    Private PC02_IN_ATIVO As Integer
    Private PC02_DH_CADASTRO As String
    Private PC02_DS_OBSERVACAO As String
    Private disposedValue As Boolean

#Region "Getters e  Setters"
    Public Property Codigo() As Integer
        Get
            Return PC02_ID_MOVIMENTO_PROCESSO
        End Get
        Set(value As Integer)
            PC02_ID_MOVIMENTO_PROCESSO = value
        End Set
    End Property

    Public Property IdMovimentoEprocesso() As Integer
        Get
            Return PC02_ID_MOVIMENTO_EPROCESSO
        End Get
        Set(value As Integer)
            PC02_ID_MOVIMENTO_EPROCESSO = value
        End Set
    End Property

    Public Property IdProcesso() As Integer
        Get
            Return PC01_ID_PROCESSO
        End Get
        Set(value As Integer)
            PC01_ID_PROCESSO = value
        End Set
    End Property

    Public Property IdStatus() As Integer
        Get
            Return DI09_ID_STATUS
        End Get
        Set(value As Integer)
            DI09_ID_STATUS = value
        End Set
    End Property

    Public Property IdUsuario() As Integer
        Get
            Return CA04_ID_USUARIO
        End Get
        Set(value As Integer)
            CA04_ID_USUARIO = value
        End Set
    End Property

    Public Property IdColaborador() As Integer
        Get
            Return DI24_ID_COLABORADOR
        End Get
        Set(value As Integer)
            DI24_ID_COLABORADOR = value
        End Set
    End Property

    Public Property Processado() As Integer
        Get
            Return PC02_IN_PROCESSADO
        End Get
        Set(value As Integer)
            PC02_IN_PROCESSADO = value
        End Set
    End Property

    Public Property Ativo() As Integer
        Get
            Return PC02_IN_ATIVO
        End Get
        Set(value As Integer)
            PC02_IN_ATIVO = value
        End Set
    End Property

    Public Property DataHoraCadastro() As String
        Get
            Return PC02_DH_CADASTRO
        End Get
        Set(value As String)
            PC02_DH_CADASTRO = value
        End Set
    End Property

    Public Property Observacao() As String
        Get
            Return PC02_DS_OBSERVACAO
        End Get
        Set(value As String)
            PC02_DS_OBSERVACAO = value
        End Set
    End Property

#End Region
    Public Sub New(Optional ByVal IdMovimentoProcesso As Integer = 0)
        If IdMovimentoProcesso > 0 Then
            Obter(IdMovimentoProcesso)
        End If
    End Sub

    Public Sub Salvar(Optional ByRef transacao As Transacao = Nothing)
        Dim cnn As New Conexao("StringConexaoDiarias")
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from PC02_MOVIMENTO_PROCESSO")
        strSQL.Append(" where PC02_ID_MOVIMENTO_PROCESSO = " & Codigo)

        dt = cnn.EditarDataTable(strSQL.ToString, transacao)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
        Else
            dr = dt.Rows(0)
        End If

        dr("PC02_ID_MOVIMENTO_PROCESSO") = ProBanco(PC02_ID_MOVIMENTO_PROCESSO, eTipoValor.CHAVE)
        dr("PC02_ID_MOVIMENTO_EPROCESSO") = ProBanco(PC02_ID_MOVIMENTO_EPROCESSO, eTipoValor.CHAVE)
        dr("PC01_ID_PROCESSO") = ProBanco(PC01_ID_PROCESSO, eTipoValor.CHAVE)
        dr("DI09_ID_STATUS") = ProBanco(DI09_ID_STATUS, eTipoValor.CHAVE)
        dr("CA04_ID_USUARIO") = ProBanco(CA04_ID_USUARIO, eTipoValor.CHAVE)
        dr("DI24_ID_COLABORADOR") = ProBanco(DI24_ID_COLABORADOR, eTipoValor.NUMERO_INTEIRO)
        dr("PC02_IN_PROCESSADO") = ProBanco(PC02_IN_PROCESSADO, eTipoValor.NUMERO_INTEIRO)
        dr("PC02_IN_ATIVO") = ProBanco(PC02_IN_ATIVO, eTipoValor.NUMERO_INTEIRO)
        dr("PC02_DH_CADASTRO") = ProBanco(PC02_DH_CADASTRO, eTipoValor.DATA_COMPLETA)
        dr("PC02_DS_OBSERVACAO") = ProBanco(PC02_DS_OBSERVACAO, eTipoValor.TEXTO_LIVRE)

        cnn.SalvarDataTable(dr)

        dt.Dispose()
        dt = Nothing

        cnn.FecharBanco()
        cnn = Nothing
    End Sub


    Public Sub Obter(IdMovimentoProcesso As Integer)
        Dim cnn As New Conexao("StringConexaoDiarias")
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from PC02_MOVIMENTO_PROCESSO")
        strSQL.Append(" where PC02_ID_MOVIMENTO_PROCESSO = " & IdMovimentoProcesso)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            PC02_ID_MOVIMENTO_PROCESSO = DoBanco(dr("PC02_ID_MOVIMENTO_PROCESSO"), eTipoValor.CHAVE)
            PC02_ID_MOVIMENTO_EPROCESSO = DoBanco(dr("PC02_ID_MOVIMENTO_EPROCESSO"), eTipoValor.CHAVE)
            PC01_ID_PROCESSO = DoBanco(dr("PC01_ID_PROCESSO"), eTipoValor.CHAVE)
            DI09_ID_STATUS = DoBanco(dr("DI09_ID_STATUS"), eTipoValor.CHAVE)
            CA04_ID_USUARIO = DoBanco(dr("CA04_ID_USUARIO"), eTipoValor.CHAVE)
            PC02_IN_PROCESSADO = DoBanco(dr("PC02_IN_PROCESSADO"), eTipoValor.BOOLEANO)
            PC02_IN_ATIVO = DoBanco(dr("PC02_IN_ATIVO"), eTipoValor.BOOLEANO)
            PC02_DH_CADASTRO = DoBanco(dr("PC02_DH_CADASTRO"), eTipoValor.DATA_COMPLETA)
            PC02_DS_OBSERVACAO = DoBanco(dr("PC02_DS_OBSERVACAO"), eTipoValor.TEXTO_LIVRE)

        End If

        cnn.FecharBanco()
    End Sub

    Public Function Pesquisar(Optional ByVal Sort As String = "",
                              Optional IdMovimentoProcesso As Integer = 0,
                              Optional IdMovimentoEProcesso As Integer = 0,
                              Optional IdProcesso As Integer = 0,
                              Optional IdStatus As Integer = 0,
                              Optional IdUsuario As Integer = 0,
                              Optional Processado As Boolean = Nothing,
                              Optional Ativo As Boolean = Nothing,
                              Optional Observacao As String = "",
                              Optional DataHoraCadastro As String = "") As DataTable

        Dim cnn As New Conexao("StringConexaoDiarias")
        Dim strSQL As New StringBuilder

        strSQL.Append(" SELECT *")
        strSQL.Append(" FROM PC02_MOVIMENTO_PROCESSO")

        strSQL.Append(" WHERE PC02_ID_MOVIMENTO_PROCESSO IS NOT NULL")

        If IdMovimentoProcesso > 0 Then
            strSQL.Append(" and PC02_ID_MOVIMENTO_PROCESSO = " & IdMovimentoProcesso)
        End If

        If IdMovimentoEProcesso > 0 Then
            strSQL.Append(" and PC02_ID_MOVIMENTO_EPROCESSO = " & IdMovimentoEProcesso)
        End If

        If IdProcesso > 0 Then
            strSQL.Append(" and PC01_ID_PROCESSO = " & IdProcesso)
        End If

        If IdStatus > 0 Then
            strSQL.Append(" and DI09_ID_STATUS = " & IdStatus)
        End If

        If IdUsuario > 0 Then
            strSQL.Append(" and CA04_ID_USUARIO = " & IdUsuario)
        End If

        If IdUsuario > 0 Then
            strSQL.Append(" and CA04_ID_USUARIO = " & IdUsuario)
        End If

        If Processado <> Nothing Then
            strSQL.Append(" and PC02_IN_PROCESSADO = " & Processado)
        End If

        If Ativo <> Nothing Then
            strSQL.Append(" and PC02_IN_ATIVO = '" & Ativo.ToString + "'")
        End If

        If Observacao <> "" Then
            strSQL.Append(" and PC02_DS_OBSERVACAO = " & Observacao)
        End If

        If Observacao <> "" Then
            strSQL.Append(" and PC02_DS_OBSERVACAO = " & Observacao)
        End If

        If IsDate(DataHoraCadastro) Then
            strSQL.Append(" and PC02_DH_CADASTRO = Convert(DateTime, '" & DataHoraCadastro & "', 103)")
        End If

        strSQL.Append(" Order By PC02_ID_MOVIMENTO_PROCESSO " & IIf(Sort = "", "ASC", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)

    End Function

    Public Function ObterUltimo() As Integer
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim CodigoUltimo As Integer

        strSQL.Append(" select max(PC02_ID_MOVIMENTO_PROCESSO) from PC02_MOVIMENTO_PROCESSO")

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
