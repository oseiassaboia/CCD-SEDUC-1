Imports Microsoft.VisualBasic
Imports System.Data

Public Class Pais
    Private TG01_ID_PAIS As Integer
    Private TG01_NM_PAIS As String
    Private TG01_NR_CENSO_PAIS As String

    Public Property Codigo() As Integer
        Get
            Return TG01_ID_PAIS
        End Get
        Set(ByVal Value As Integer)
            TG01_ID_PAIS = Value
        End Set
    End Property
    Public Property Nome() As String
        Get
            Return TG01_NM_PAIS
        End Get
        Set(ByVal Value As String)
            TG01_NM_PAIS = Value
        End Set
    End Property
    Public Property CodigoCenso() As String
        Get
            Return TG01_NR_CENSO_PAIS
        End Get
        Set(ByVal Value As String)
            TG01_NR_CENSO_PAIS = Value
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
        strSQL.Append(" from DBGERAL.DBO.TG01_PAIS")
        strSQL.Append(" where TG01_ID_PAIS = " & Codigo)

        dt = cnn.EditarDataTable(strSQL.ToString)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
        Else
            dr = dt.Rows(0)
        End If

        dr("TG01_NM_PAIS") = ProBanco(TG01_NM_PAIS, eTipoValor.TEXTO)
        dr("TG01_NR_CENSO_PAIS") = ProBanco(TG01_NR_CENSO_PAIS, eTipoValor.NUMERO_DECIMAL)

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
        strSQL.Append(" from DBGERAL.DBO.TG01_PAIS")
        strSQL.Append(" where TG01_ID_PAIS = " & Codigo)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            TG01_ID_PAIS = DoBanco(dr("TG01_ID_PAIS"), eTipoValor.CHAVE)
            TG01_NM_PAIS = DoBanco(dr("TG01_NM_PAIS"), eTipoValor.TEXTO)
            TG01_NR_CENSO_PAIS = DoBanco(dr("TG01_NR_CENSO_PAIS"), eTipoValor.NUMERO_DECIMAL)
        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Function Pesquisar(Optional ByVal Sort As String = "", Optional Codigo As Integer = 0, Optional Nome As String = "", Optional CodigoCenso As String = "") As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from DBGERAL.DBO.TG01_PAIS")
        'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
        strSQL.Append(" where TG01_ID_PAIS is not null")

        If Codigo > 0 Then
            strSQL.Append(" and TG01_ID_PAIS = " & Codigo)
        End If

        If Nome <> "" Then
            strSQL.Append(" and upper(TG01_NM_PAIS) like '%" & Nome.toUpper & "%'")
        End If

        If IsNumeric(CodigoCenso.Replace(".", "")) Then
            strSQL.Append(" and TG01_NR_CENSO_PAIS = " & CodigoCenso.Replace(".", "").Replace(",", "."))
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "TG01_ID_PAIS", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function ObterTabela() As DataTable
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" select TG01_ID_PAIS as CODIGO, TG01_NM_PAIS as DESCRICAO")
        strSQL.Append(" from DBGERAL.DBO.TG01_PAIS")
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

        strSQL.Append(" select max(TG01_ID_PAIS) from DBGERAL.DBO.TG01_PAIS")

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
        strSQL.Append(" from DBGERAL.DBO.TG01_PAIS")
        strSQL.Append(" where TG01_ID_PAIS = " & Codigo)

        LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

        cnn.FecharBanco()
        cnn = Nothing

        Return LinhasAfetadas
    End Function

End Class

