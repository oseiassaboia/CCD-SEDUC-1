Imports Microsoft.VisualBasic
Imports System.Data

Public Class Sexo
    Private TG08_ID_SEXO As Integer
    Private TG08_NM_SEXO As String

    Public Property Codigo() As Integer
        Get
            Return TG08_ID_SEXO
        End Get
        Set(ByVal Value As Integer)
            TG08_ID_SEXO = Value
        End Set
    End Property
    Public Property Nome() As String
        Get
            Return TG08_NM_SEXO
        End Get
        Set(ByVal Value As String)
            TG08_NM_SEXO = Value
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
        strSQL.Append(" from DBGERAL.DBO.TG08_SEXO")
        strSQL.Append(" where TG08_ID_SEXO = " & Codigo)

        dt = cnn.EditarDataTable(strSQL.ToString)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
        Else
            dr = dt.Rows(0)
        End If

        dr("TG08_NM_SEXO") = ProBanco(TG08_NM_SEXO, eTipoValor.TEXTO)

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
        strSQL.Append(" from DBGERAL.DBO.TG08_SEXO")
        strSQL.Append(" where TG08_ID_SEXO = " & Codigo)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            TG08_ID_SEXO = DoBanco(dr("TG08_ID_SEXO"), eTipoValor.CHAVE)
            TG08_NM_SEXO = DoBanco(dr("TG08_NM_SEXO"), eTipoValor.TEXTO)
        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Function Pesquisar(Optional ByVal Sort As String = "", Optional Codigo As Integer = 0, Optional Nome As String = "") As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from DBGERAL.DBO.TG08_SEXO")
        'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
        strSQL.Append(" where TG08_ID_SEXO is not null")

        If Codigo > 0 Then
            strSQL.Append(" and TG08_ID_SEXO = " & Codigo)
        End If

        If Nome <> "" Then
            strSQL.Append(" and upper(TG08_NM_SEXO) like '%" & Nome.toUpper & "%'")
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "TG08_ID_SEXO", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function ObterTabela() As DataTable
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" select TG08_ID_SEXO as CODIGO, TG08_NM_SEXO as DESCRICAO")
        strSQL.Append(" from DBGERAL.DBO.TG08_SEXO")
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

        strSQL.Append(" select max(TG08_ID_SEXO) from DBGERAL.DBO.TG08_SEXO")

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
        strSQL.Append(" from DBGERAL.DBO.TG08_SEXO")
        strSQL.Append(" where TG08_ID_SEXO = " & Codigo)

        LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

        cnn.FecharBanco()
        cnn = Nothing

        Return LinhasAfetadas
    End Function

End Class

