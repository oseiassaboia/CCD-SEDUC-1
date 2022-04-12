Imports Microsoft.VisualBasic
Imports System.Data

Public Class Regional
    Private TG05_ID_REGIONAL As Integer
    Private TG05_NM_REGIONAL As String
    Private TG05_NR_SIAFEM As String

    Public Property Codigo() As Integer
        Get
            Return TG05_ID_REGIONAL
        End Get
        Set(ByVal Value As Integer)
            TG05_ID_REGIONAL = Value
        End Set
    End Property
    Public Property Nome() As String
        Get
            Return TG05_NM_REGIONAL
        End Get
        Set(ByVal Value As String)
            TG05_NM_REGIONAL = Value
        End Set
    End Property
    Public Property CodigoSiafem() As String
        Get
            Return TG05_NR_SIAFEM
        End Get
        Set(ByVal Value As String)
            TG05_NR_SIAFEM = Value
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
        strSQL.Append(" from DBGERAL.DBO.TG05_REGIONAL")
        strSQL.Append(" where TG05_ID_REGIONAL = " & Codigo)

        dt = cnn.EditarDataTable(strSQL.ToString)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
        Else
            dr = dt.Rows(0)
        End If

        dr("TG05_NM_REGIONAL") = ProBanco(TG05_NM_REGIONAL, eTipoValor.TEXTO)
        dr("TG05_NR_SIAFEM") = ProBanco(TG05_NR_SIAFEM, eTipoValor.NUMERO_INTEIRO)

        cnn.SalvarDataTable(dr)

        dt.Dispose()
        dt = Nothing

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Sub Obter(ByVal Codigo As Integer)
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from DBGERAL.DBO.TG05_REGIONAL")
        strSQL.Append(" where TG05_ID_REGIONAL = " & Codigo)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            TG05_ID_REGIONAL = DoBanco(dr("TG05_ID_REGIONAL"), eTipoValor.CHAVE)
            TG05_NM_REGIONAL = DoBanco(dr("TG05_NM_REGIONAL"), eTipoValor.TEXTO)
            TG05_NR_SIAFEM = DoBanco(dr("TG05_NR_SIAFEM"), eTipoValor.NUMERO_INTEIRO)
        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Function Pesquisar(Optional ByVal Sort As String = "", Optional Codigo As Integer = 0, Optional Nome As String = "", Optional CodigoSiafem As String = "") As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from DBGERAL.DBO.TG05_REGIONAL")
        'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
        strSQL.Append(" where TG05_ID_REGIONAL is not null")

        If Codigo > 0 Then
            strSQL.Append(" and TG05_ID_REGIONAL = " & Codigo)
        End If

        If Nome <> "" Then
            strSQL.Append(" and upper(TG05_NM_REGIONAL) like '%" & Nome.toUpper & "%'")
        End If

        If CodigoSiafem <> "" Then
            strSQL.Append(" and upper(TG05_NR_SIAFEM) like '%" & CodigoSiafem.toUpper & "%'")
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "TG05_ID_REGIONAL", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function ObterTabela() As DataTable
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" select TG05_ID_REGIONAL as CODIGO, TG05_NM_REGIONAL as DESCRICAO")
        strSQL.Append(" from DBGERAL.DBO.TG05_REGIONAL")
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

        strSQL.Append(" select max(TG05_ID_REGIONAL) from DBGERAL.DBO.TG05_REGIONAL")

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
    Public Function Excluir(ByVal Codigo As Integer) As Integer
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim LinhasAfetadas As Integer

        strSQL.Append(" delete ")
        strSQL.Append(" from DBGERAL.DBO.TG05_REGIONAL")
        strSQL.Append(" where TG05_ID_REGIONAL = " & Codigo)

        LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

        cnn.FecharBanco()
        cnn = Nothing

        Return LinhasAfetadas
    End Function

End Class

