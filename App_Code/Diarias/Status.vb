Imports System.Data
Imports Microsoft.VisualBasic

Public Class Status
    Implements IDisposable

    Private DI09_COD_STATUS As Integer
    'Private FKDI09SV04_COD_LOTACAO As Integer VERIFICAR POSSIBILIDADE DE CRIACAO RH36_LOTACAO
    Private FKDI09DI20_COD_SITUACAO_EPROCESSO As Integer
    Private FKDI09DI21_COD_FASE_EPROCESSO As Integer
    Private FKDI09DI18_COD_LOTACAO_EPROCESSO As Integer
    Private FKDI09DI17_COD_TIPO_PROCESSO As Integer
    Private DI09_DESCRICAO As String
    Private DI09_RESUMO As String
    Private DI09_INICIO As Boolean
    Private DI09_FIM As Boolean
    Private DI09_RECEBER As Boolean
    Private DI09_RH_VISUALIZAR As Boolean
    Private RH36_ID_LOTACAO As Integer

    Private disposedValue As Boolean

    Public Property Codigo() As Integer
        Get
            Return DI09_COD_STATUS
        End Get
        Set(value As Integer)
            DI09_COD_STATUS = value
        End Set
    End Property

    Public Property CodSituacaoEProcesso() As Integer
        Get
            Return FKDI09DI20_COD_SITUACAO_EPROCESSO
        End Get
        Set(value As Integer)
            FKDI09DI20_COD_SITUACAO_EPROCESSO = value
        End Set
    End Property

    Public Property CodFaseEprocesso() As Integer
        Get
            Return FKDI09DI21_COD_FASE_EPROCESSO
        End Get
        Set(value As Integer)
            FKDI09DI21_COD_FASE_EPROCESSO = value
        End Set
    End Property

    Public Property CodLotacaoEprocesso() As Integer
        Get
            Return FKDI09DI18_COD_LOTACAO_EPROCESSO
        End Get
        Set(value As Integer)
            FKDI09DI18_COD_LOTACAO_EPROCESSO = value
        End Set
    End Property

    Public Property CodTipoProcesso() As Integer
        Get
            Return FKDI09DI17_COD_TIPO_PROCESSO
        End Get
        Set(value As Integer)
            FKDI09DI17_COD_TIPO_PROCESSO = value
        End Set
    End Property

    Public Property Descricao() As String
        Get
            Return DI09_DESCRICAO
        End Get
        Set(value As String)
            DI09_DESCRICAO = value
        End Set
    End Property

    Public Property Resumo() As String
        Get
            Return DI09_RESUMO
        End Get
        Set(value As String)
            DI09_RESUMO = value
        End Set
    End Property

    Public Property Inicio() As Boolean
        Get
            Return DI09_INICIO
        End Get
        Set(value As Boolean)
            DI09_INICIO = value
        End Set
    End Property

    Public Property Fim() As Boolean
        Get
            Return DI09_FIM
        End Get
        Set(value As Boolean)
            DI09_FIM = value
        End Set
    End Property

    Public Property RhVisualizar() As Boolean
        Get
            Return DI09_RH_VISUALIZAR
        End Get
        Set(value As Boolean)
            DI09_RH_VISUALIZAR = value
        End Set
    End Property

    Public Property Lotacao() As Integer
        Get
            Return RH36_ID_LOTACAO
        End Get
        Set(value As Integer)
            RH36_ID_LOTACAO = value
        End Set
    End Property

    Public Property Receber As Boolean
        Get
            Return DI09_RECEBER
        End Get
        Set(value As Boolean)
            DI09_RECEBER = value
        End Set
    End Property

    Public Sub New(Optional ByVal IdStatus As Integer = 0)
        If IdStatus > 0 Then
            Obter(IdStatus)
        End If
    End Sub

    Public Sub Salvar()
        Dim cnn As New Conexao("StringConexaoDiarias")
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from DI09_STATUS")
        strSQL.Append(" where DI09_COD_STATUS = " & Codigo)

        dt = cnn.EditarDataTable(strSQL.ToString)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
        Else
            dr = dt.Rows(0)
        End If

        dr("DI09_COD_STATUS") = ProBanco(DI09_COD_STATUS, eTipoValor.CHAVE)
        dr("FKDI09DI20_COD_SITUACAO_EPROCESSO") = ProBanco(FKDI09DI20_COD_SITUACAO_EPROCESSO, eTipoValor.CHAVE)
        dr("FKDI09DI21_COD_FASE_EPROCESSO") = ProBanco(FKDI09DI21_COD_FASE_EPROCESSO, eTipoValor.CHAVE)
        dr("FKDI09DI18_COD_LOTACAO_EPROCESSO") = ProBanco(FKDI09DI18_COD_LOTACAO_EPROCESSO, eTipoValor.CHAVE)
        dr("FKDI09DI17_COD_TIPO_PROCESSO") = ProBanco(FKDI09DI17_COD_TIPO_PROCESSO, eTipoValor.CHAVE)
        dr("DI09_DESCRICAO") = ProBanco(DI09_DESCRICAO, eTipoValor.TEXTO_LIVRE)
        dr("DI09_RESUMO") = ProBanco(DI09_RESUMO, eTipoValor.TEXTO_LIVRE)
        dr("DI09_INICIO") = ProBanco(DI09_INICIO, eTipoValor.BOOLEANO)
        dr("DI09_FIM") = ProBanco(DI09_FIM, eTipoValor.BOOLEANO)
        dr("DI09_RECEBER") = ProBanco(DI09_RECEBER, eTipoValor.BOOLEANO)
        dr("DI09_RH_VISUALIZAR") = ProBanco(DI09_RH_VISUALIZAR, eTipoValor.BOOLEANO)
        dr("RH36_ID_LOTACAO") = ProBanco(RH36_ID_LOTACAO, eTipoValor.CHAVE)

        cnn.SalvarDataTable(dr)

        dt.Dispose()
        dt = Nothing

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Sub Obter(IdStatus As Integer)
        Dim cnn As New Conexao("StringConexaoDiarias")
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from DI09_STATUS")
        strSQL.Append(" where DI09_COD_STATUS = " & IdStatus)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            DI09_COD_STATUS = DoBanco(dr("DI09_COD_STATUS"), eTipoValor.CHAVE)
            FKDI09DI20_COD_SITUACAO_EPROCESSO = DoBanco(dr("FKDI09DI20_COD_SITUACAO_EPROCESSO"), eTipoValor.CHAVE)
            FKDI09DI21_COD_FASE_EPROCESSO = DoBanco(dr("FKDI09DI21_COD_FASE_EPROCESSO"), eTipoValor.CHAVE)
            FKDI09DI18_COD_LOTACAO_EPROCESSO = DoBanco(dr("FKDI09DI18_COD_LOTACAO_EPROCESSO"), eTipoValor.CHAVE)
            FKDI09DI17_COD_TIPO_PROCESSO = DoBanco(dr("FKDI09DI17_COD_TIPO_PROCESSO"), eTipoValor.CHAVE)
            DI09_DESCRICAO = DoBanco(dr("DI09_DESCRICAO"), eTipoValor.TEXTO_LIVRE)
            DI09_RESUMO = DoBanco(dr("DI09_RESUMO"), eTipoValor.TEXTO_LIVRE)
            DI09_INICIO = DoBanco(dr("DI09_INICIO"), eTipoValor.BOOLEANO)
            DI09_FIM = DoBanco(dr("DI09_FIM"), eTipoValor.BOOLEANO)
            DI09_RECEBER = DoBanco(dr("DI09_RECEBER"), eTipoValor.BOOLEANO)
            DI09_RH_VISUALIZAR = DoBanco(dr("DI09_RH_VISUALIZAR"), eTipoValor.BOOLEANO)
            RH36_ID_LOTACAO = DoBanco(dr("RH36_ID_LOTACAO"), eTipoValor.CHAVE)

        End If

        cnn.FecharBanco()
    End Sub

    Public Function Pesquisar(Optional ByVal Sort As String = "",
                              Optional IdStatus As Integer = 0,
                              Optional IdSituacaoEProcesso As Integer = 0,
                              Optional IdFaseEProcesso As Integer = 0,
                              Optional IdLotacaoEProcesso As Integer = 0,
                              Optional IdTipoProcesso As Integer = 0,
                              Optional Descricao As String = "",
                              Optional Resumo As String = "",
                              Optional Inicio As Boolean = Nothing,
                              Optional Fim As Boolean = Nothing,
                              Optional RHVisualizar As Boolean = Nothing,
                              Optional Lotacao As Integer = 0) As DataTable

        Dim cnn As New Conexao("StringConexaoDiarias")
        Dim strSQL As New StringBuilder

        strSQL.Append(" SELECT *")
        strSQL.Append(" FROM DI09_STATUS")

        strSQL.Append(" WHERE DI09_COD_STATUS IS NOT NULL")

        If IdStatus > 0 Then
            strSQL.Append(" and DI09_COD_STATUS = " & IdStatus)
        End If

        If IdSituacaoEProcesso > 0 Then
            strSQL.Append(" and FKDI09DI20_COD_SITUACAO_EPROCESSO = " & IdSituacaoEProcesso)
        End If

        If IdFaseEProcesso > 0 Then
            strSQL.Append(" and FKDI09DI21_COD_FASE_EPROCESSO = " & IdFaseEProcesso)
        End If

        If IdLotacaoEProcesso > 0 Then
            strSQL.Append(" and FKDI09DI18_COD_LOTACAO_EPROCESSO = " & IdLotacaoEProcesso)
        End If

        If IdTipoProcesso > 0 Then
            strSQL.Append(" and FKDI09DI17_COD_TIPO_PROCESSO = " & IdTipoProcesso)
        End If

        If Descricao <> "" Then
            strSQL.Append(" and upper(DI09_DESCRICAO) like '%" & Descricao.ToUpper & "%'")
        End If

        If Resumo <> "" Then
            strSQL.Append(" and upper(DI09_RESUMO) like '%" & Resumo.ToUpper & "%'")
        End If

        If Inicio Then
            strSQL.Append(" and DI09_INICIO = 1")
        End If

        If Fim Then
            strSQL.Append(" and DI09_FIM = 1")
        End If

        If RHVisualizar Then
            strSQL.Append(" and DI09_RH_VISUALIZAR = 1")
        End If

        If Lotacao > 0 Then
            strSQL.Append(" and RH36_ID_LOTACAO = " & Lotacao)
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "DI09_COD_STATUS", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)

    End Function

    Public Function ObterStatus(ByVal ServidorBenef As Integer) As Integer
        Dim cnn As New Conexao("StringConexaoDiarias")
        Dim strSQL As New StringBuilder
        Dim CodigoStatus As Integer

        strSQL.Append("  SELECT PC02.DI09_ID_STATUS  " & vbCrLf)
        strSQL.Append("    FROM PC01_PROCESSO                            PC01   " & vbCrLf)
        strSQL.Append("    LEFT JOIN DI17_TIPO_PROCESSO                  DI17 ON DI17.DI17_COD_TIPO_PROCESSO = PC01.DI17_ID_TIPO_PROCESSO  " & vbCrLf)
        strSQL.Append("    JOIN PC02_MOVIMENTO_PROCESSO                  PC02 ON PC02.PC01_ID_PROCESSO       = PC01.PC01_ID_PROCESSO   " & vbCrLf)
        strSQL.Append("    JOIN DI09_STATUS                              DI09 ON DI09.DI09_COD_STATUS        = PC02.DI09_ID_STATUS   " & vbCrLf)
        strSQL.Append("    JOIN [10.31.35.4].[DBRH].[DBO].[RH36_LOTACAO] RH36 ON RH36.RH36_ID_LOTACAO        = DI09.RH36_ID_LOTACAO   " & vbCrLf)
        strSQL.Append("   WHERE PC01.PC01_ID_PROCESSO IS NOT NULL  " & vbCrLf)
        strSQL.Append("     AND PC01.CA01_ID_APLICACAO = 91  " & vbCrLf)
        strSQL.Append("     AND PC02.PC02_IN_ATIVO = 1  " & vbCrLf)
        strSQL.Append("     AND RH02_ID_SERVIDOR_BENEF = " & ServidorBenef & vbCrLf)
        strSQL.Append("   ORDER BY PC01.PC01_ID_PROCESSO  " & vbCrLf)

        With cnn.AbrirDataTable(strSQL.ToString)
            If Not IsDBNull(.Rows(0)(0)) Then
                CodigoStatus = .Rows(0)(0)
            Else
                CodigoStatus = 0
            End If
        End With

        cnn.FecharBanco()
        cnn = Nothing

        Return CodigoStatus

    End Function

    Public Function ObterUltimo() As Integer
        Dim cnn As New Conexao("StringConexaoDiarias")
        Dim strSQL As New StringBuilder
        Dim CodigoUltimo As Integer

        strSQL.Append(" select max(DI09_COD_STATUS) from DI09_STATUS")

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

    Public Function ObterTabela(Optional CodTipoProcesso As Integer = 0) As DataTable
        Dim cnn As New Conexao("StringConexaoDiarias")
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" select DI09_COD_STATUS as CODIGO, DI09_DESCRICAO + ' - ' + DI09_RESUMO as DESCRICAO")
        strSQL.Append(" from DI09_STATUS")
        If CodTipoProcesso > 0 Then
            strSQL.Append(" WHERE FKDI09DI17_COD_TIPO_PROCESSO = " & CodTipoProcesso)
        End If

        dt = cnn.AbrirDataTable(strSQL.ToString)

        cnn.FecharBanco()
        cnn = Nothing

        Return dt

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

    Public Function EmissaoPortariaConcessivaValida(ByVal ServidorBenef As Integer) As Boolean
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim retorno As Boolean

        strSQL.Append("  SELECT PC02.DI09_ID_STATUS  " & vbCrLf)
        strSQL.Append("    FROM PC01_PROCESSO                            PC01   " & vbCrLf)
        strSQL.Append("    LEFT JOIN DI17_TIPO_PROCESSO                  DI17 ON DI17.DI17_COD_TIPO_PROCESSO = PC01.DI17_ID_TIPO_PROCESSO  " & vbCrLf)
        strSQL.Append("    JOIN PC02_MOVIMENTO_PROCESSO                  PC02 ON PC02.PC01_ID_PROCESSO       = PC01.PC01_ID_PROCESSO   " & vbCrLf)
        strSQL.Append("    JOIN DI09_STATUS                              DI09 ON DI09.DI09_COD_STATUS        = PC02.DI09_ID_STATUS   " & vbCrLf)
        strSQL.Append("    JOIN [10.31.35.4].[DBRH].[DBO].[RH36_LOTACAO] RH36 ON RH36.RH36_ID_LOTACAO        = DI09.RH36_ID_LOTACAO   " & vbCrLf)
        strSQL.Append("   WHERE PC01.PC01_ID_PROCESSO IS NOT NULL  " & vbCrLf)
        strSQL.Append("     AND PC01.CA01_ID_APLICACAO = 91  " & vbCrLf)
        strSQL.Append("     AND PC02.PC02_IN_ATIVO = 1  " & vbCrLf)
        strSQL.Append("     AND RH02_ID_SERVIDOR_BENEF = " & ServidorBenef & vbCrLf)
        strSQL.Append("     AND DI09.DI09_RESUMO LIKE '%Portaria Concessiva Emitida' " & vbCrLf)
        strSQL.Append("   ORDER BY PC01.PC01_ID_PROCESSO  " & vbCrLf)

        Try

            With cnn.AbrirDataTable(strSQL.ToString)
                If .Rows.Count > 0 Then
                    retorno = True
                Else
                    retorno = False
                End If
            End With

        Catch ex As Exception
            Dim erro As String = ex.ToString
            retorno = False
        End Try

        cnn = Nothing
        Return retorno
    End Function

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
