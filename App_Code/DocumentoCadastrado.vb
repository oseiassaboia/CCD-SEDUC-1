Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.VisualBasic

Public Class DocumentoCadastrado
    Private RH104_ID_DOC_CADASTRADO As Integer
    Private RH102_ID_SOLIC_CAD_DOC As Integer
    Private RH96_ID_LICENCA_DOC As Integer

    Private disposedValue As Boolean

#Region "Getters e  Setters"

    Public Property Codigo As Integer
        Get
            Return RH104_ID_DOC_CADASTRADO
        End Get
        Set(value As Integer)
            RH104_ID_DOC_CADASTRADO = value
        End Set
    End Property

    Public Property IdSolicCadDoc As Integer
        Get
            Return RH102_ID_SOLIC_CAD_DOC
        End Get
        Set(value As Integer)
            RH102_ID_SOLIC_CAD_DOC = value
        End Set
    End Property

    Public Property idLicencaDoc As Integer
        Get
            Return RH96_ID_LICENCA_DOC
        End Get
        Set(value As Integer)
            RH96_ID_LICENCA_DOC = value
        End Set
    End Property

#End Region

    Public Sub New(Optional ByVal Codigo As Integer = 0)
        If Codigo > 0 Then
            Obter(Codigo)
        End If
    End Sub



    Private Sub Obter(codigo As Integer)
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH104_DOC_CADASTRADO")
        strSQL.Append(" where RH104_ID_DOC_CADASTRADO = " & codigo)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            RH104_ID_DOC_CADASTRADO = DoBanco(dr("RH104_ID_DOC_CADASTRADO"), eTipoValor.CHAVE)
            RH102_ID_SOLIC_CAD_DOC = DoBanco(dr("RH102_ID_SOLIC_CAD_DOC"), eTipoValor.CHAVE)
            RH96_ID_LICENCA_DOC = DoBanco(dr("RH96_ID_LICENCA_DOC"), eTipoValor.CHAVE)

        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Function Pesquisar(Optional ByVal Sort As String = "",
                              Optional Codigo As Integer = 0,
                              Optional IdCadastroDocumento As Integer = 0,
                              Optional idLicencaDocumento As Integer = 0) As DataTable

        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH104_DOC_CADASTRADO")
        strSQL.Append(" where RH104_ID_DOC_CADASTRADO is not null")

        If Codigo > 0 Then
            strSQL.Append(" AND RH104_ID_DOC_CADASTRADO = " & Codigo)
        End If

        If IdCadastroDocumento > 0 Then
            strSQL.Append(" AND RH102_ID_SOLIC_CAD_DOC = " & IdCadastroDocumento)
        End If

        If idLicencaDocumento > 0 Then
            strSQL.Append(" AND RH96_ID_LICENCA_DOC = " & idLicencaDocumento)
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "RH104_ID_DOC_CADASTRADO", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)

    End Function

    Public Sub Salvar(Optional ByRef transacao As Transacao = Nothing)
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH104_DOC_CADASTRADO")
        strSQL.Append(" where RH104_ID_DOC_CADASTRADO = " & Codigo)

        dt = cnn.EditarDataTable(strSQL.ToString, transacao)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
        Else
            dr = dt.Rows(0)
        End If

        dr("RH104_ID_DOC_CADASTRADO") = ProBanco(RH104_ID_DOC_CADASTRADO, eTipoValor.CHAVE)
        dr("RH102_ID_SOLIC_CAD_DOC") = ProBanco(RH102_ID_SOLIC_CAD_DOC, eTipoValor.CHAVE)
        dr("RH96_ID_LICENCA_DOC") = ProBanco(RH96_ID_LICENCA_DOC, eTipoValor.CHAVE)

        cnn.SalvarDataTable(dr)

        dt.Dispose()
        dt = Nothing

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Function ObterUltimo(Optional ByRef transacao As Transacao = Nothing) As Integer

        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim CodigoUltimo As Integer

        strSQL.Append(" select max(RH104_ID_DOC_CADASTRADO) from RH104_DOC_CADASTRADO")

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

End Class
