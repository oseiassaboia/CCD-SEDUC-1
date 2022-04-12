Imports Microsoft.VisualBasic
Imports System.Data

Public Class Area
    Private DE12_ID_AREA As Integer
    Private DE12_NM_AREA As String
    Private DE12_CD_AREA_CENSO As String

    Public Property Codigo() As Integer
        Get
            Return DE12_ID_AREA
        End Get
        Set(ByVal Value As Integer)
            DE12_ID_AREA = Value
        End Set
    End Property
    Public Property Nome() As String
        Get
            Return DE12_NM_AREA
        End Get
        Set(ByVal Value As String)
            DE12_NM_AREA = Value
        End Set
    End Property

    Public Property CodigoCenso() As String
        Get
            Return DE12_CD_AREA_CENSO
        End Get
        Set(ByVal Value As String)
            DE12_CD_AREA_CENSO = Value
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
        strSQL.Append(" from DBDIARIO..DE12_AREA")
        strSQL.Append(" where DE12_ID_AREA = " & Codigo)

        dt = cnn.EditarDataTable(strSQL.ToString)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
        Else
            dr = dt.Rows(0)
        End If

        dr("DE12_NM_AREA") = ProBanco(DE12_NM_AREA, eTipoValor.TEXTO)
        dr("DE12_CD_AREA_CENSO") = ProBanco(DE12_CD_AREA_CENSO, eTipoValor.TEXTO)

        cnn.SalvarDataTable(dr)

        dt.Dispose()
        dt = Nothing

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Sub Obter(ByVal Codigo As String)
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from DBDIARIO..DE12_AREA")
        strSQL.Append(" where DE12_ID_AREA = " & Codigo)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            DE12_ID_AREA = DoBanco(dr("DE12_ID_AREA"), eTipoValor.CHAVE)
            DE12_NM_AREA = DoBanco(dr("DE12_NM_AREA"), eTipoValor.TEXTO)
            DE12_CD_AREA_CENSO = DoBanco(dr("DE12_CD_AREA_CENSO"), eTipoValor.TEXTO)
        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Function Pesquisar(Optional ByVal Sort As String = "", Optional Codigo As Integer = 0, Optional Nome As String = "", Optional CodigoCenso As String = "") As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from DBDIARIO..DE12_AREA")
        'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
        strSQL.Append(" where DE12_ID_AREA is not null")

        If Codigo > 0 Then
            strSQL.Append(" and DE12_ID_AREA = " & Codigo)
        End If

        If Nome <> "" Then
            strSQL.Append(" and upper(DE12_NM_AREA) like '%" & Nome.ToUpper & "%'")
        End If

        If CodigoCenso <> "" Then
            strSQL.Append(" and upper(DE12_CD_AREA_CENSO) like '%" & CodigoCenso.ToUpper & "%'")
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "DE12_ID_AREA", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function ObterTabela() As DataTable
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" select DE12_ID_AREA as CODIGO, DE12_NM_AREA as DESCRICAO")
        strSQL.Append(" from DBDIARIO..DE12_AREA")
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

        strSQL.Append(" select max(DE12_ID_AREA) from DBDIARIO..DE12_AREA")

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
    Public Function Excluir(ByVal Codigo As String) As Integer
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim LinhasAfetadas As Integer

        strSQL.Append(" delete ")
        strSQL.Append(" from DBDIARIO..DE12_AREA")
        strSQL.Append(" where DE12_ID_AREA = " & Codigo)

        LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

        cnn.FecharBanco()
        cnn = Nothing

        Return LinhasAfetadas
    End Function

End Class

