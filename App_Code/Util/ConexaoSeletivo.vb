Imports Microsoft.VisualBasic
Imports System.Data
Imports Microsoft.Practices.EnterpriseLibrary.Data
Imports System.Configuration



Public Class ConexaoSeletivo
    Dim db As Database = DatabaseFactory.CreateDatabase("StringConexaoDBSELETIVO")
    Dim da As SqlClient.SqlDataAdapter
    Dim cmd As SqlClient.SqlCommand

    Public ReadOnly Property ConnectionString() As String
        Get
            Return db.CreateConnection.ConnectionString
        End Get
    End Property

    Public Function AbrirDataTable(ByVal SQL As String, Optional ByRef Tra As Transacao = Nothing) As DataTable
        Dim dt As New DataTable

        If Tra Is Nothing Then
            'Se n?o foi passado transa??o, seta a conex?o
            Try

                Return db.ExecuteDataSet(Data.CommandType.Text, SQL).Tables(0)

            Catch ex As Exception
                Return Nothing

            End Try
        Else
            'Caso tenha sido, o command recebe o command da transacao
            cmd = Tra.cmd

            cmd.CommandType = CommandType.Text
            cmd.CommandText = SQL

            da = New SqlClient.SqlDataAdapter(cmd)
            da.Fill(dt)

            Return dt
        End If

    End Function

    Public Function EditarDataTable(ByVal SQL As String, Optional ByRef Tra As Transacao = Nothing) As DataTable
        Dim dt As New DataTable

        If Tra Is Nothing Then
            'Se n?o foi passado transa??o, seta a conex?o
            cmd = New SqlClient.SqlCommand
            cmd.Connection = New SqlClient.SqlConnection(db.CreateConnection.ConnectionString)
        Else
            'Caso tenha sido, o command recebe o command da transacao
            cmd = Tra.cmd
        End If

        cmd.CommandType = CommandType.Text
        cmd.CommandText = SQL

        da = New SqlClient.SqlDataAdapter(cmd)
        da.Fill(dt)

        If Tra Is Nothing Then
            'Se n?o foi passado a transa??o, libera o command
            cmd.Dispose()
            cmd = Nothing
        End If

        Return dt
    End Function

    Public Sub SalvarDataTable(ByVal dRow As DataRow)
        Dim objBuilder As New SqlClient.SqlCommandBuilder(da)
        Dim dTable As DataTable

        dTable = dRow.Table

        If dTable.Rows.Count = 0 Then
            dTable.Rows.Add(dRow)
            da.InsertCommand = objBuilder.GetInsertCommand
        Else
            da.UpdateCommand = objBuilder.GetUpdateCommand
        End If

        da.Update(dTable)

        da.Dispose()
        objBuilder.Dispose()
        dTable.Dispose()

        da = Nothing
        objBuilder = Nothing
        dTable = Nothing

    End Sub

    Public Sub CancelarDataTable()
        da.Dispose()
        da = Nothing
    End Sub

    Public Function ExecutarSQL(ByVal SQL As String) As Integer
        Dim RowsAffected As Integer

        Try
            RowsAffected = db.ExecuteNonQuery(Data.CommandType.Text, SQL)

            Return RowsAffected
        Catch
            Return -1
        End Try
    End Function

    Public Sub New()

    End Sub

    Public Sub FecharBanco()
        If Not da Is Nothing Then
            da.Dispose()
        End If
        da = Nothing

        If Not cmd Is Nothing Then
            cmd.Dispose()
        End If
        cmd = Nothing

        db = Nothing
    End Sub

End Class
