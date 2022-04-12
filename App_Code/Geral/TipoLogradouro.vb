Imports Microsoft.VisualBasic
Imports System.Data

Public Class TipoLogradouro
    Private TG16_ID_TIPO_LOGRADOURO As Integer
    Private TG16_NM_TIPO_LOGRADOURO As String

    Public Property Codigo() As Integer
        Get
            Return TG16_ID_TIPO_LOGRADOURO
        End Get
        Set(ByVal Value As Integer)
            TG16_ID_TIPO_LOGRADOURO = Value
        End Set
    End Property
    Public Property Nome() As String
        Get
            Return TG16_NM_TIPO_LOGRADOURO
        End Get
        Set(ByVal Value As String)
            TG16_NM_TIPO_LOGRADOURO = Value
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
        strSQL.Append(" from DBGERAL.DBO.TG16_TIPO_LOGRADOURO")
        strSQL.Append(" where TG16_ID_TIPO_LOGRADOURO = " & Codigo)

        dt = cnn.EditarDataTable(strSQL.ToString)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
        Else
            dr = dt.Rows(0)
        End If

        dr("TG16_NM_TIPO_LOGRADOURO") = ProBanco(TG16_NM_TIPO_LOGRADOURO, eTipoValor.TEXTO)

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
        strSQL.Append(" from DBGERAL.DBO.TG16_TIPO_LOGRADOURO")
        strSQL.Append(" where TG16_ID_TIPO_LOGRADOURO = " & Codigo)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            TG16_ID_TIPO_LOGRADOURO = DoBanco(dr("TG16_ID_TIPO_LOGRADOURO"), eTipoValor.CHAVE)
            TG16_NM_TIPO_LOGRADOURO = DoBanco(dr("TG16_NM_TIPO_LOGRADOURO"), eTipoValor.TEXTO)
        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Function Pesquisar(Optional ByVal Sort As String = "", Optional Codigo As Integer = 0, Optional Nome As String = "") As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from DBGERAL.DBO.TG16_TIPO_LOGRADOURO")
        'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
        strSQL.Append(" where TG16_ID_TIPO_LOGRADOURO is not null")

        If Codigo > 0 Then
            strSQL.Append(" and TG16_ID_TIPO_LOGRADOURO = " & Codigo)
        End If

        If Nome <> "" Then
            strSQL.Append(" and upper(TG16_NM_TIPO_LOGRADOURO) like '%" & Nome.ToUpper & "%'")
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "TG16_ID_TIPO_LOGRADOURO", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function PesquisarParaCep(ByVal Nome As String) As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" select TG16_ID_TIPO_LOGRADOURO, TG16_NM_TIPO_LOGRADOURO ")
        strSQL.Append(" from DBGERAL.DBO.TG16_TIPO_LOGRADOURO")
        strSQL.Append(" where TG16_ID_TIPO_LOGRADOURO is not null")
        strSQL.Append(" and TG16_NM_TIPO_LOGRADOURO = '" & Nome.ToUpper & "'")

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function ObterTabela() As DataTable
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" select TG16_ID_TIPO_LOGRADOURO as CODIGO, TG16_NM_TIPO_LOGRADOURO as DESCRICAO")
        strSQL.Append(" from DBGERAL.DBO.TG16_TIPO_LOGRADOURO")
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

        strSQL.Append(" select max(TG16_ID_TIPO_LOGRADOURO) from DBGERAL.DBO.TG16_TIPO_LOGRADOURO")

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
        strSQL.Append(" from DBGERAL.DBO.TG16_TIPO_LOGRADOURO")
        strSQL.Append(" where TG16_ID_TIPO_LOGRADOURO = " & Codigo)

        LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

        cnn.FecharBanco()
        cnn = Nothing

        Return LinhasAfetadas
    End Function

End Class

