Imports Microsoft.VisualBasic
Imports System.Data

Public Class RacaCor
    Private TG11_ID_RACA_COR As Integer
    Private TG11_NM_RACA_COR As String
    Private TG11_NR_RACA_COR_CENSO As String

    Public Property Codigo() As Integer
        Get
            Return TG11_ID_RACA_COR
        End Get
        Set(ByVal Value As Integer)
            TG11_ID_RACA_COR = Value
        End Set
    End Property
    Public Property Nome() As String
        Get
            Return TG11_NM_RACA_COR
        End Get
        Set(ByVal Value As String)
            TG11_NM_RACA_COR = Value
        End Set
    End Property
    Public Property CodigoCenso() As String
        Get
            Return TG11_NR_RACA_COR_CENSO
        End Get
        Set(ByVal Value As String)
            TG11_NR_RACA_COR_CENSO = Value
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
        strSQL.Append(" from DBGERAL.DBO.TG11_RACA_COR")
        strSQL.Append(" where TG11_ID_RACA_COR = " & Codigo)

        dt = cnn.EditarDataTable(strSQL.ToString)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
        Else
            dr = dt.Rows(0)
        End If

        dr("TG11_NM_RACA_COR") = ProBanco(TG11_NM_RACA_COR, eTipoValor.TEXTO)
        dr("TG11_NR_RACA_COR_CENSO") = ProBanco(TG11_NR_RACA_COR_CENSO, eTipoValor.NUMERO_INTEIRO)

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
        strSQL.Append(" from DBGERAL.DBO.TG11_RACA_COR")
        strSQL.Append(" where TG11_ID_RACA_COR = " & Codigo)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            TG11_ID_RACA_COR = DoBanco(dr("TG11_ID_RACA_COR"), eTipoValor.CHAVE)
            TG11_NM_RACA_COR = DoBanco(dr("TG11_NM_RACA_COR"), eTipoValor.TEXTO)
            TG11_NR_RACA_COR_CENSO = DoBanco(dr("TG11_NR_RACA_COR_CENSO"), eTipoValor.NUMERO_INTEIRO)
        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Function Pesquisar(Optional ByVal Sort As String = "", Optional Codigo As Integer = 0, Optional Nome As String = "", Optional CodigoCenso As String = "") As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from DBGERAL.DBO.TG11_RACA_COR")
        'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
        strSQL.Append(" where TG11_ID_RACA_COR is not null")

        If Codigo > 0 Then
            strSQL.Append(" and TG11_ID_RACA_COR = " & Codigo)
        End If

        If Nome <> "" Then
            strSQL.Append(" and upper(TG11_NM_RACA_COR) like '%" & Nome.toUpper & "%'")
        End If

        If CodigoCenso <> "" Then
            strSQL.Append(" and upper(TG11_NR_RACA_COR_CENSO) like '%" & CodigoCenso.toUpper & "%'")
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "TG11_ID_RACA_COR", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function ObterTabela() As DataTable
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" select TG11_ID_RACA_COR as CODIGO, TG11_NM_RACA_COR as DESCRICAO")
        strSQL.Append(" from DBGERAL.DBO.TG11_RACA_COR")
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

        strSQL.Append(" select max(TG11_ID_RACA_COR) from DBGERAL.DBO.TG11_RACA_COR")

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
        strSQL.Append(" from DBGERAL.DBO.TG11_RACA_COR")
        strSQL.Append(" where TG11_ID_RACA_COR = " & Codigo)

        LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

        cnn.FecharBanco()
        cnn = Nothing

        Return LinhasAfetadas
    End Function

End Class

