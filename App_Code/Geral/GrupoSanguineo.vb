Imports Microsoft.VisualBasic
Imports System.Data

Public Class GrupoSanguineo
    Private TG09_ID_GRUPO_SANGUINEO As Integer
    Private TG09_SG_GRUPO_SANGUINEO As String

    Public Property Codigo() As Integer
        Get
            Return TG09_ID_GRUPO_SANGUINEO
        End Get
        Set(ByVal Value As Integer)
            TG09_ID_GRUPO_SANGUINEO = Value
        End Set
    End Property
    Public Property Nome() As String
        Get
            Return TG09_SG_GRUPO_SANGUINEO
        End Get
        Set(ByVal Value As String)
            TG09_SG_GRUPO_SANGUINEO = Value
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

        strSQL.Append(" Select * ")
        strSQL.Append(" from DBGERAL.DBO.TG09_GRUPO_SANGUINEO")
        strSQL.Append(" where TG09_ID_GRUPO_SANGUINEO = " & Codigo)

        dt = cnn.EditarDataTable(strSQL.ToString)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
        Else
            dr = dt.Rows(0)
        End If

        dr("TG09_SG_GRUPO_SANGUINEO") = ProBanco(TG09_SG_GRUPO_SANGUINEO, eTipoValor.TEXTO)

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

        strSQL.Append(" Select * ")
        strSQL.Append(" from DBGERAL.DBO.TG09_GRUPO_SANGUINEO")
        strSQL.Append(" where TG09_ID_GRUPO_SANGUINEO = " & Codigo)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            TG09_ID_GRUPO_SANGUINEO = DoBanco(dr("TG09_ID_GRUPO_SANGUINEO"), eTipoValor.CHAVE)
            TG09_SG_GRUPO_SANGUINEO = DoBanco(dr("TG09_SG_GRUPO_SANGUINEO"), eTipoValor.TEXTO)
        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Function Pesquisar(Optional ByVal Sort As String = "", Optional Codigo As Integer = 0, Optional Nome As String = "") As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" Select * ")
        strSQL.Append(" from DBGERAL.DBO.TG09_GRUPO_SANGUINEO")
        'strSQL.Append(" left join tabela On coluna1 = coluna2 ")
        strSQL.Append(" where TG09_ID_GRUPO_SANGUINEO Is Not null")

        If Codigo > 0 Then
            strSQL.Append(" And TG09_ID_GRUPO_SANGUINEO = " & Codigo)
        End If

        If Nome <> "" Then
            strSQL.Append(" And upper(TG09_SG_GRUPO_SANGUINEO) like '%" & Nome.toUpper & "%'")
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "TG09_ID_GRUPO_SANGUINEO", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function ObterTabela() As DataTable
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" select TG09_ID_GRUPO_SANGUINEO as CODIGO, TG09_SG_GRUPO_SANGUINEO as DESCRICAO")
        strSQL.Append(" from DBGERAL.DBO.TG09_GRUPO_SANGUINEO")
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

        strSQL.Append(" select max(TG09_ID_GRUPO_SANGUINEO) from DBGERAL.DBO.TG09_GRUPO_SANGUINEO")

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
        strSQL.Append(" from DBGERAL.DBO.TG09_GRUPO_SANGUINEO")
        strSQL.Append(" where TG09_ID_GRUPO_SANGUINEO = " & Codigo)

        LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

        cnn.FecharBanco()
        cnn = Nothing

        Return LinhasAfetadas
    End Function

End Class

