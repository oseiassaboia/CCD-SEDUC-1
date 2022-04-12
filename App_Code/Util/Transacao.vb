Imports Microsoft.VisualBasic
Imports System.Data
Imports Microsoft.Practices.EnterpriseLibrary.Data

Public Class Transacao

    Public cmd As New SqlClient.SqlCommand

    Private db As Database
    Private cnn As SqlClient.SqlConnection
    Private tra As SqlClient.SqlTransaction

    Public Sub New(Optional ConnectionName As String = "")
        If (ConnectionName <> "") Then
            db = DatabaseFactory.CreateDatabase(ConnectionName)
        Else
            db = DatabaseFactory.CreateDatabase("StringConexao")
        End If

    End Sub

    Public Sub IniciarTransacao()
        cnn = New SqlClient.SqlConnection(db.CreateConnection.ConnectionString)
        cnn.Open()
        cmd = cnn.CreateCommand
        tra = cnn.BeginTransaction()
        cmd.Connection = cnn
        cmd.Transaction = tra
    End Sub

    Public Sub ConfirmarTransacao()
        tra.Commit()
        cmd.Dispose()
        cmd = Nothing
        cnn.Close()
        cnn = Nothing
    End Sub

    Public Sub CancelarTransacao()
        tra.Rollback()
        cmd.Dispose()
        cmd = Nothing
        cnn.Close()
        cnn = Nothing
    End Sub

End Class

