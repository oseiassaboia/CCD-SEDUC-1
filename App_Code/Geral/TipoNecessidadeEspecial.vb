Imports Microsoft.VisualBasic
Imports System.Data

Public Class TipoNecessidadeEspecial
    Private TG18_ID_TIPO_NECESSIDADE_ESP As Integer
    Private TG18_NM_TIPO_NECESSIDADE_ESP As String

    Public Property Codigo() As Integer
        Get
            Return TG18_ID_TIPO_NECESSIDADE_ESP
        End Get
        Set(ByVal Value As Integer)
            TG18_ID_TIPO_NECESSIDADE_ESP = Value
        End Set
    End Property
    Public Property Nome() As String
        Get
            Return TG18_NM_TIPO_NECESSIDADE_ESP
        End Get
        Set(ByVal Value As String)
            TG18_NM_TIPO_NECESSIDADE_ESP = Value
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
        strSQL.Append(" from DBGERAL.DBO.TG18_TIPO_NECESSIDADE_ESP")
        strSQL.Append(" where TG18_ID_TIPO_NECESSIDADE_ESP = " & Codigo)

        dt = cnn.EditarDataTable(strSQL.ToString)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
        Else
            dr = dt.Rows(0)
        End If

        dr("TG18_NM_TIPO_NECESSIDADE_ESP") = ProBanco(TG18_NM_TIPO_NECESSIDADE_ESP, eTipoValor.TEXTO)

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
        strSQL.Append(" from DBGERAL.DBO.TG18_TIPO_NECESSIDADE_ESP")
        strSQL.Append(" where TG18_ID_TIPO_NECESSIDADE_ESP = " & Codigo)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            TG18_ID_TIPO_NECESSIDADE_ESP = DoBanco(dr("TG18_ID_TIPO_NECESSIDADE_ESP"), eTipoValor.CHAVE)
            TG18_NM_TIPO_NECESSIDADE_ESP = DoBanco(dr("TG18_NM_TIPO_NECESSIDADE_ESP"), eTipoValor.TEXTO)
        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Function Pesquisar(Optional ByVal Sort As String = "", Optional Codigo As Integer = 0, Optional Nome As String = "") As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from DBGERAL.DBO.TG18_TIPO_NECESSIDADE_ESP")
        'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
        strSQL.Append(" where TG18_ID_TIPO_NECESSIDADE_ESP is not null")

        If Codigo > 0 Then
            strSQL.Append(" and TG18_ID_TIPO_NECESSIDADE_ESP = " & Codigo)
        End If

        If Nome <> "" Then
            strSQL.Append(" and upper(TG18_NM_TIPO_NECESSIDADE_ESP) like '%" & Nome.toUpper & "%'")
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "TG18_ID_TIPO_NECESSIDADE_ESP", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function ObterTabela() As DataTable
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" select TG18_ID_TIPO_NECESSIDADE_ESP as CODIGO, TG18_NM_TIPO_NECESSIDADE_ESP as DESCRICAO")
        strSQL.Append(" from DBGERAL.DBO.TG18_TIPO_NECESSIDADE_ESP")
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

        strSQL.Append(" select max(TG18_ID_TIPO_NECESSIDADE_ESP) from DBGERAL.DBO.TG18_TIPO_NECESSIDADE_ESP")

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
        strSQL.Append(" from DBGERAL.DBO.TG18_TIPO_NECESSIDADE_ESP")
        strSQL.Append(" where TG18_ID_TIPO_NECESSIDADE_ESP = " & Codigo)

        LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

        cnn.FecharBanco()
        cnn = Nothing

        Return LinhasAfetadas
    End Function

End Class

