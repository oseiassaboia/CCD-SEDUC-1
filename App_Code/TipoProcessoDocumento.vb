Imports System.Data
Imports Microsoft.VisualBasic

Public Class TipoProcessoDocumento
    Implements IDisposable

    Private RH95_ID_STATUS_CATEG_TIPO_DOC As Integer
    Private DI09_ID_STATUS As Integer
    Private DI02_ID_TIPO_DOCUMENTO As Integer
    Private RH94_ID_CATEGORIA_DOC As Integer
    Private RH95_IN_PERMITE_APROVA_PARCIAL As Boolean
    Private RH95_IN_SUBMETE_ANALISE_DOC As Boolean
    Private RH95_IN_CADASTRO_OBRIGATORIO As Boolean
    Private RH95_NR_ORDEM_SHOW As Integer

    Private disposedValue As Boolean

    Public Property CodigoStatus() As Integer
        Get
            Return RH95_ID_STATUS_CATEG_TIPO_DOC
        End Get
        Set(value As Integer)
            RH95_ID_STATUS_CATEG_TIPO_DOC = value
        End Set
    End Property

    Public Property StatusCategTipoDoc() As Integer
        Get
            Return DI09_ID_STATUS
        End Get
        Set(value As Integer)
            DI09_ID_STATUS = value
        End Set
    End Property

    Public Property CategoriaDocumento() As Integer
        Get
            Return RH94_ID_CATEGORIA_DOC
        End Get
        Set(value As Integer)
            RH94_ID_CATEGORIA_DOC = value
        End Set
    End Property

    Public Property TipoDocumento() As Integer
        Get
            Return DI02_ID_TIPO_DOCUMENTO
        End Get
        Set(value As Integer)
            DI02_ID_TIPO_DOCUMENTO = value
        End Set
    End Property

    Public Property PermiteAprovarParcial() As Boolean
        Get
            Return RH95_IN_PERMITE_APROVA_PARCIAL
        End Get
        Set(value As Boolean)
            RH95_IN_PERMITE_APROVA_PARCIAL = value
        End Set
    End Property

    Public Property SubmeterAnalise() As Boolean
        Get
            Return RH95_IN_SUBMETE_ANALISE_DOC
        End Get
        Set(value As Boolean)
            RH95_IN_SUBMETE_ANALISE_DOC = value
        End Set
    End Property

    Public Property CadastroObrigatoria() As Boolean
        Get
            Return RH95_IN_CADASTRO_OBRIGATORIO
        End Get
        Set(value As Boolean)
            RH95_IN_CADASTRO_OBRIGATORIO = value
        End Set
    End Property

    Public Property NroOrdem() As Integer
        Get
            Return RH95_NR_ORDEM_SHOW
        End Get
        Set(value As Integer)
            RH95_NR_ORDEM_SHOW = value
        End Set
    End Property

    Public Sub New(Optional IdCategoriaProcessoDocumento As Integer = 0)
        If IdCategoriaProcessoDocumento > 0 Then
            Obter(IdCategoriaProcessoDocumento)
        End If
    End Sub

    Public Sub Salvar()
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH95_STATUS_CATEG_TIPO_DOC")
        strSQL.Append(" where RH95_ID_STATUS_CATEG_TIPO_DOC = " & CodigoStatus)

        dt = cnn.EditarDataTable(strSQL.ToString)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
        Else
            dr = dt.Rows(0)
        End If

        dr("RH95_ID_STATUS_CATEG_TIPO_DOC") = ProBanco(RH95_ID_STATUS_CATEG_TIPO_DOC, eTipoValor.CHAVE)
        dr("DI09_ID_STATUS") = ProBanco(DI09_ID_STATUS, eTipoValor.CHAVE)
        dr("RH94_ID_CATEGORIA_DOC") = ProBanco(RH94_ID_CATEGORIA_DOC, eTipoValor.CHAVE)
        dr("DI02_ID_TIPO_DOCUMENTO") = ProBanco(DI02_ID_TIPO_DOCUMENTO, eTipoValor.CHAVE)
        dr("RH95_IN_PERMITE_APROVA_PARCIAL") = ProBanco(RH95_IN_PERMITE_APROVA_PARCIAL, eTipoValor.BOOLEANO)
        dr("RH95_IN_SUBMETE_ANALISE_DOC") = ProBanco(RH95_IN_SUBMETE_ANALISE_DOC, eTipoValor.BOOLEANO)
        dr("RH95_IN_CADASTRO_OBRIGATORIO") = ProBanco(RH95_IN_CADASTRO_OBRIGATORIO, eTipoValor.BOOLEANO)
        dr("RH95_NR_ORDEM_SHOW") = ProBanco(RH95_NR_ORDEM_SHOW, eTipoValor.NUMERO_INTEIRO)

        cnn.SalvarDataTable(dr)

        dt.Dispose()
        dt = Nothing

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Sub Obter(CodigoTipoProcessoDocumento As Integer)
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH95_STATUS_CATEG_TIPO_DOC")
        strSQL.Append(" where RH95_ID_STATUS_CATEG_TIPO_DOC = " & CodigoTipoProcessoDocumento)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            RH95_ID_STATUS_CATEG_TIPO_DOC = DoBanco(dr("RH95_ID_STATUS_CATEG_TIPO_DOC"), eTipoValor.CHAVE)
            DI09_ID_STATUS = DoBanco(dr("DI09_ID_STATUS"), eTipoValor.CHAVE)
            RH94_ID_CATEGORIA_DOC = DoBanco(dr("RH94_ID_CATEGORIA_DOC"), eTipoValor.CHAVE)
            DI02_ID_TIPO_DOCUMENTO = DoBanco(dr("DI02_ID_TIPO_DOCUMENTO"), eTipoValor.CHAVE)
            RH95_IN_PERMITE_APROVA_PARCIAL = DoBanco(dr("RH95_IN_PERMITE_APROVA_PARCIAL"), eTipoValor.BOOLEANO)
            RH95_IN_SUBMETE_ANALISE_DOC = DoBanco(dr("RH95_IN_SUBMETE_ANALISE_DOC"), eTipoValor.BOOLEANO)
            RH95_IN_CADASTRO_OBRIGATORIO = DoBanco(dr("RH95_IN_CADASTRO_OBRIGATORIO"), eTipoValor.BOOLEANO)
            RH95_NR_ORDEM_SHOW = DoBanco(dr("RH95_NR_ORDEM_SHOW"), eTipoValor.NUMERO_INTEIRO)
        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Function Pesquisar(Optional Sort As String = "",
                              Optional IdTipoProcessoDocumento As Integer = 0,
                              Optional IdTipoProcesso As Integer = 0,
                              Optional IdTipoDocumento As Integer = 0,
                              Optional IdCategoriaDocumento As Integer = 0,
                              Optional IdTipoLicencaProcesso As Integer = 0) As DataTable

        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" SELECT *")
        strSQL.Append("   FROM RH95_STATUS_CATEG_TIPO_DOC RH95  ")
        strSQL.Append("   JOIN RH94_CATEGORIA_DOC         RH94 ON RH94.RH94_ID_CATEGORIA_DOC = RH95.RH94_ID_CATEGORIA_DOC  ")
        strSQL.Append("   JOIN RH93_TIPO_LICENCA_PROCESSO RH93 ON RH93.RH93_ID_TIPO_LICENCA_PROCESSO = RH94.RH93_ID_TIPO_LICENCA_PROCESSO ")
        strSQL.Append("  WHERE RH95.RH95_ID_STATUS_CATEG_TIPO_DOC IS NOT NULL ")


        If IdTipoProcessoDocumento > 0 Then
            strSQL.Append(" and RH95.RH95_ID_STATUS_CATEG_TIPO_DOC = " & IdTipoProcessoDocumento)
        End If

        If IdTipoProcesso > 0 Then
            strSQL.Append(" and RH93.DI17_ID_TIPO_PROCESSO = " & IdTipoProcesso)
        End If

        If IdTipoDocumento > 0 Then
            strSQL.Append(" AND RH95.DI02_ID_TIPO_DOCUMENTO = " & IdTipoDocumento)
        End If

        If IdCategoriaDocumento > 0 Then
            strSQL.Append(" AND RH95.RH94_ID_CATEGORIA_DOC = " & IdCategoriaDocumento)
        End If

        If IdTipoLicencaProcesso > 0 Then
            strSQL.Append(" AND RH93.RH93_ID_TIPO_LICENCA_PROCESSO= " & IdTipoLicencaProcesso)
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "RH95.RH95_ID_STATUS_CATEG_TIPO_DOC", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)

    End Function

    Public Function ObterDocumentos(ListDocumento As String) As DataTable

        Dim cnn As New Conexao("StringConexao")
        Dim strSQL As New StringBuilder

        strSQL.Append(" SELECT DI09.DI09_DESCRICAO, DI09.DI09_RESUMO, DI02.DI02_DESCRICAO, RH95.DI02_ID_TIPO_DOCUMENTO, RH95.RH95_ID_STATUS_CATEG_TIPO_DOC ")
        strSQL.Append("   FROM RH95_STATUS_CATEG_TIPO_DOC RH95 ")
        strSQL.Append("   JOIN [172.16.2.71].DBDIARIAS_TREINAMENTO.DBO.DI02_TIPO_DOCUMENTO DI02 ON DI02.DI02_COD_TIPO_DOCUMENTO = RH95.DI02_ID_TIPO_DOCUMENTO ")
        strSQL.Append("   JOIN [172.16.2.71].DBDIARIAS_TREINAMENTO.DBO.DI09_STATUS         DI09 ON DI09.DI09_COD_STATUS = RH95.DI09_ID_STATUS ")
        strSQL.Append("  WHERE RH95.DI02_ID_TIPO_DOCUMENTO IS NOT NULL    ")
        If ListDocumento <> "" Then
            strSQL.Append(" AND RH95.DI02_ID_TIPO_DOCUMENTO IN ( " & ListDocumento & " )")
        Else
            strSQL.Append(" AND RH95.DI02_ID_TIPO_DOCUMENTO IN ( 0 )")
        End If

        strSQL.Append(" Order By RH95.DI02_ID_TIPO_DOCUMENTO  ")

        Return cnn.AbrirDataTable(strSQL.ToString)

    End Function

    Public Function ObterDocumentosPorParametros(IdTipoLicencaProcesso As Integer, Optional IdCategoriaDoc As Integer = 0) As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" SELECT *")
        strSQL.Append("   FROM RH95_STATUS_CATEG_TIPO_DOC RH95")
        strSQL.Append("   JOIN RH94_CATEGORIA_DOC         RH94 ON RH94.RH94_ID_CATEGORIA_DOC = RH95.RH94_ID_CATEGORIA_DOC")
        strSQL.Append("   JOIN RH93_TIPO_LICENCA_PROCESSO RH93 ON RH93.RH93_ID_TIPO_LICENCA_PROCESSO = RH94.RH93_ID_TIPO_LICENCA_PROCESSO")
        strSQL.Append("  WHERE RH95.RH95_ID_STATUS_CATEG_TIPO_DOC is not null")

        If IdTipoLicencaProcesso > 0 Then
            strSQL.Append(" AND RH93.RH93_ID_TIPO_LICENCA_PROCESSO= " & IdTipoLicencaProcesso)
        End If

        If IdCategoriaDoc > 0 Then
            strSQL.Append(" AND RH94.RH94_ID_CATEGORIA_DOC= " & IdCategoriaDoc)
        End If

        strSQL.Append(" Order By RH95.RH95_ID_STATUS_CATEG_TIPO_DOC")

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function CarregarComboDocumentoTramite(Optional IdTipoProcesso As Integer = 0,
                                                  Optional IdStatus As Integer = 0,
                                                  Optional IdLotacao As Integer = 0) As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" SELECT DI02_ID_TIPO_DOCUMENTO as CODIGO, DI02_DESCRICAO as DESCRICAO, " & vbCrLf)
        strSQL.Append("        DI02_ID_TIPO_DOCUMENTO " & vbCrLf)
        strSQL.Append("   FROM RH95_STATUS_CATEG_TIPO_DOC                                      RH95 " & vbCrLf)
        strSQL.Append("   JOIN RH94_CATEGORIA_DOC                                              RH94 ON RH94.RH94_ID_CATEGORIA_DOC  = RH95.RH94_ID_CATEGORIA_DOC " & vbCrLf)
        strSQL.Append("   JOIN [172.16.2.71].[DBDIARIAS_TREINAMENTO].[dbo].DI02_TIPO_DOCUMENTO DI02 ON RH95.DI02_ID_TIPO_DOCUMENTO = DI02.DI02_COD_TIPO_DOCUMENTO " & vbCrLf)
        strSQL.Append("   JOIN [172.16.2.71].[DBDIARIAS_TREINAMENTO].[dbo].DI09_STATUS         DI09 ON DI09.DI09_COD_STATUS        = RH95.DI09_ID_STATUS " & vbCrLf)
        strSQL.Append("   JOIN RH36_LOTACAO                                                    RH36 ON RH36.RH36_ID_LOTACAO        = DI09.RH36_ID_LOTACAO " & vbCrLf)
        strSQL.Append("  WHERE RH94.RH94_NM_CATEGORIA_DOC like '%trâmite%'    " & vbCrLf)

        If IdStatus > 0 Then
            strSQL.Append("  AND RH95.DI09_ID_STATUS = " & IdStatus & vbCrLf)
        End If

        If IdLotacao > 0 Then
            strSQL.Append("  AND RH36.RH36_ID_LOTACAO = " & IdLotacao & vbCrLf)
        End If

        strSQL.Append(" Order By RH95.RH95_ID_STATUS_CATEG_TIPO_DOC")

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function CarregarComboDocumentoSolicitado(IdTipoProcesso As Integer,
                                                  Optional IdStatus As Integer = 0,
                                                  Optional IdLotacao As Integer = 0,
                                                  Optional IdServidor As Integer = 0,
                                                  Optional IdTipoDocumento As Integer = 0) As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" SELECT RH95.DI02_ID_TIPO_DOCUMENTO as CODIGO, DI02_DESCRICAO as DESCRICAO ")
        strSQL.Append("   FROM RH95_STATUS_CATEG_TIPO_DOC                                      RH95 ")
        strSQL.Append("   JOIN RH94_CATEGORIA_DOC                                              RH94  ON RH94.RH94_ID_CATEGORIA_DOC   = RH95.RH94_ID_CATEGORIA_DOC ")
        strSQL.Append("   JOIN [172.16.2.71].[DBDIARIAS_TREINAMENTO].[dbo].DI02_TIPO_DOCUMENTO DI02  ON RH95.DI02_ID_TIPO_DOCUMENTO  = DI02.DI02_COD_TIPO_DOCUMENTO ")
        strSQL.Append("   JOIN [172.16.2.71].[DBDIARIAS_TREINAMENTO].[dbo].DI09_STATUS         DI09  ON DI09.DI09_COD_STATUS         = RH95.DI09_ID_STATUS ")
        strSQL.Append("   JOIN RH36_LOTACAO                                                    RH36  ON RH36.RH36_ID_LOTACAO         = DI09.RH36_ID_LOTACAO ")
        strSQL.Append("   JOIN RH102_SOLIC_CAD_DOC                                             RH102 ON RH102.DI02_ID_TIPO_DOCUMENTO = DI02.DI02_COD_TIPO_DOCUMENTO ")
        strSQL.Append("   JOIN RH30_LICENCA                                                    RH30  ON RH30.RH30_ID_LICENCA         = RH102.RH30_ID_LICENCA ")
        strSQL.Append("  WHERE RH94.RH94_NM_CATEGORIA_DOC like '%trâmite%' ")

        If IdStatus > 0 Then
            strSQL.Append("  AND RH95.DI09_ID_STATUS = " & IdStatus)
        End If

        If IdLotacao > 0 Then
            strSQL.Append("  AND RH36.RH36_ID_LOTACAO = " & IdLotacao)
        End If

        If IdServidor > 0 Then
            strSQL.Append("  AND RH30.RH02_ID_SERVIDOR = " & IdServidor)
        End If

        If IdTipoDocumento > 0 Then
            strSQL.Append("  AND RH95.DI02_ID_TIPO_DOCUMENTO = " & IdTipoDocumento)
        End If

        strSQL.Append(" Order By RH95.RH95_ID_STATUS_CATEG_TIPO_DOC")

        Return cnn.AbrirDataTable(strSQL.ToString)
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
