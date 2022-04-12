Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.VisualBasic

Public Class LicencaProcesso
    Implements IDisposable

    Private RH92_ID_LICENCA_PROCESSO As Integer
    Private RH30_ID_LICENCA As Integer
    Private PC01_ID_PROCESSO As Integer
    Private CA04_ID_USUARIO As Integer
    Private RH92_DH_CADASTRO As String

    Private disposedValue As Boolean

#Region "Getters e  Setters"
    Public Property Codigo() As Integer
        Get
            Return RH92_ID_LICENCA_PROCESSO
        End Get
        Set(value As Integer)
            RH92_ID_LICENCA_PROCESSO = value
        End Set
    End Property

    Public Property IdLicenca() As Integer
        Get
            Return RH30_ID_LICENCA
        End Get
        Set(value As Integer)
            RH30_ID_LICENCA = value
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

    Public Property Usuario() As Integer
        Get
            Return CA04_ID_USUARIO
        End Get
        Set(value As Integer)
            CA04_ID_USUARIO = value
        End Set
    End Property

    Public Property DataCadastro() As String
        Get
            Return RH92_DH_CADASTRO
        End Get
        Set(value As String)
            RH92_DH_CADASTRO = value
        End Set
    End Property
#End Region
    Public Sub New(Optional ByVal IdLicenca As Integer = 0)
        If IdLicenca > 0 Then
            Obter(IdLicenca)
        End If
    End Sub

    Public Sub Salvar(Optional ByRef transacao As Transacao = Nothing)
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH92_LICENCA_PROCESSO")
        strSQL.Append(" where RH92_ID_LICENCA_PROCESSO  = " & Codigo)

        dt = cnn.EditarDataTable(strSQL.ToString, transacao)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
        Else
            dr = dt.Rows(0)
        End If

        dr("RH92_ID_LICENCA_PROCESSO") = ProBanco(RH92_ID_LICENCA_PROCESSO, eTipoValor.CHAVE)
        dr("RH30_ID_LICENCA") = ProBanco(RH30_ID_LICENCA, eTipoValor.CHAVE)
        dr("PC01_ID_PROCESSO") = ProBanco(PC01_ID_PROCESSO, eTipoValor.CHAVE)
        dr("CA04_ID_USUARIO") = ProBanco(CA04_ID_USUARIO, eTipoValor.NUMERO_INTEIRO)
        dr("RH92_DH_CADASTRO") = ProBanco(RH92_DH_CADASTRO, eTipoValor.DATA_COMPLETA)

        cnn.SalvarDataTable(dr)

        dt.Dispose()
        dt = Nothing

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Sub Obter(ByVal Id As Integer)
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH92_LICENCA_PROCESSO")
        strSQL.Append(" where RH92_ID_LICENCA_PROCESSO = " & Id)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            RH92_ID_LICENCA_PROCESSO = DoBanco(dr("RH92_ID_LICENCA_PROCESSO"), eTipoValor.CHAVE)
            RH30_ID_LICENCA = DoBanco(dr("RH30_ID_LICENCA"), eTipoValor.CHAVE)
            PC01_ID_PROCESSO = DoBanco(dr("PC01_ID_PROCESSO"), eTipoValor.CHAVE)
            CA04_ID_USUARIO = DoBanco(dr("CA04_ID_USUARIO"), eTipoValor.NUMERO_INTEIRO)
            RH92_DH_CADASTRO = DoBanco(dr("RH92_DH_CADASTRO"), eTipoValor.DATA_COMPLETA)
        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Function Pesquisar(Optional ByVal Sort As String = "",
                              Optional IdLicenca As Integer = 0,
                              Optional IdServidor As Integer = 0,
                              Optional IdTipoLicenca As Integer = 0,
                              Optional NrProcesso As Integer = 0,
                              Optional IdPessoa As Integer = 0) As DataTable

        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" SELECT *")
        strSQL.Append(" FROM RH92_LICENCA_PROCESSO")
        strSQL.Append(" inner join RH30_LICENCA on RH30_LICENCA.RH30_ID_LICENCA = RH92_LICENCA_PROCESSO.RH30_ID_LICENCA")
        strSQL.Append(" INNER JOIN RH02_SERVIDOR ON RH30_LICENCA.RH02_ID_SERVIDOR = RH02_SERVIDOR.RH02_ID_SERVIDOR")
        strSQL.Append(" INNER JOIN RH29_TIPO_LICENCA ON RH30_LICENCA.RH29_ID_TIPO_LICENCA = RH29_TIPO_LICENCA.RH29_ID_TIPO_LICENCA")

        strSQL.Append(" WHERE RH30_LICENCA.RH30_ID_LICENCA IS NOT NULL")

        If IdLicenca > 0 Then
            strSQL.Append(" and RH30_LICENCA.RH30_ID_LICENCA = " & IdLicenca)
        End If

        If IdServidor > 0 Then
            strSQL.Append(" and RH30_LICENCA.RH02_ID_SERVIDOR = " & IdServidor)
        End If

        If IdTipoLicenca > 0 Then
            strSQL.Append(" and RH29_ID_TIPO_LICENCA = " & IdTipoLicenca)
        End If

        If NrProcesso > 0 Then
            ' strSQL.Append(" and RH29_NR_PROCESSO = " & NrProcesso)
            strSQL.Append(" and RH30_NR_PROCESSO = " & NrProcesso)
        End If

        If IdPessoa > 0 Then
            ' strSQL.Append(" and RH29_NR_PROCESSO = " & NrProcesso)
            strSQL.Append(" and RH02_SERVIDOR.RH01_ID_PESSOA = " & IdPessoa)
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "RH30_LICENCA.RH30_ID_LICENCA", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)

    End Function

    Public Function ObterTabela() As DataTable
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" select RH30_ID_LICENCA as CODIGO, RH30_ID_LICENCA as DESCRICAO")
        strSQL.Append(" from RH30_LICENCA")
        strSQL.Append(" order by 2 ")

        dt = cnn.AbrirDataTable(strSQL.ToString)

        cnn.FecharBanco()
        cnn = Nothing

        Return dt

    End Function

    Public Function ObterUltimo() As Integer

        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim CodigoUltimo As Integer

        strSQL.Append(" select max(RH30_ID_LICENCA) from RH30_LICENCA")

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

#Region "IDisposable Support"

    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' Tarefa pendente: descartar o estado gerenciado (objetos gerenciados)
            End If

            ' Tarefa pendente: liberar recursos não gerenciados (objetos não gerenciados) e substituir o finalizador
            ' Tarefa pendente: definir campos grandes como nulos
            disposedValue = True
        End If
    End Sub

    ' ' Tarefa pendente: substituir o finalizador somente se 'Dispose(disposing As Boolean)' tiver o código para liberar recursos não gerenciados
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
#End Region
End Class
