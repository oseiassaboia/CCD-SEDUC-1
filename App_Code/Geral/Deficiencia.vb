Imports Microsoft.VisualBasic
Imports System.Data

Public Class Deficiencia
    Private TG15_ID_DEFICIENCIA As Integer
    Private TG15_NM_DEFICIENCIA As String

    Public Property Codigo() As Integer
        Get
            Return TG15_ID_DEFICIENCIA
        End Get
        Set(ByVal Value As Integer)
            TG15_ID_DEFICIENCIA = Value
        End Set
    End Property
    Public Property Nome() As String
        Get
            Return TG15_NM_DEFICIENCIA
        End Get
        Set(ByVal Value As String)
            TG15_NM_DEFICIENCIA = Value
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
        strSQL.Append(" from DBGERAL.DBO.TG15_DEFICIENCIA")
        strSQL.Append(" where TG15_ID_DEFICIENCIA = " & Codigo)

        dt = cnn.EditarDataTable(strSQL.ToString)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
        Else
            dr = dt.Rows(0)
        End If

        dr("TG15_NM_DEFICIENCIA") = ProBanco(TG15_NM_DEFICIENCIA, eTipoValor.TEXTO)

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
        strSQL.Append(" from DBGERAL.DBO.TG15_DEFICIENCIA")
        strSQL.Append(" where TG15_ID_DEFICIENCIA = " & Codigo)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            TG15_ID_DEFICIENCIA = DoBanco(dr("TG15_ID_DEFICIENCIA"), eTipoValor.CHAVE)
            TG15_NM_DEFICIENCIA = DoBanco(dr("TG15_NM_DEFICIENCIA"), eTipoValor.TEXTO)
        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Function Pesquisar(Optional ByVal Sort As String = "", Optional Codigo As Integer = 0, Optional Nome As String = "") As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from DBGERAL.DBO.TG15_DEFICIENCIA")
        'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
        strSQL.Append(" where TG15_ID_DEFICIENCIA is not null")

        If Codigo > 0 Then
            strSQL.Append(" and TG15_ID_DEFICIENCIA = " & Codigo)
        End If

        If Nome <> "" Then
            strSQL.Append(" and upper(TG15_NM_DEFICIENCIA) like '%" & Nome.toUpper & "%'")
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "TG15_ID_DEFICIENCIA", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function ObterTabela() As DataTable
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" select TG15_ID_DEFICIENCIA as CODIGO, TG15_NM_DEFICIENCIA as DESCRICAO")
        strSQL.Append(" from DBGERAL.DBO.TG15_DEFICIENCIA")
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

        strSQL.Append(" select max(TG15_ID_DEFICIENCIA) from DBGERAL.DBO.TG15_DEFICIENCIA")

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
        strSQL.Append(" from DBGERAL.DBO.TG15_DEFICIENCIA")
        strSQL.Append(" where TG15_ID_DEFICIENCIA = " & Codigo)

        LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

        cnn.FecharBanco()
        cnn = Nothing

        Return LinhasAfetadas
    End Function

End Class

