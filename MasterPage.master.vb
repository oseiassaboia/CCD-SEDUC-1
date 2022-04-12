
Partial Class MasterPage
    Inherits System.Web.UI.MasterPage

    Private Sub MasterPage_Load(sender As Object, e As EventArgs) Handles Me.Load
        If Not Page.IsPostBack Then

            'If Session("CodigoUsuario") = 0 Or Session("CodigoUsuario") Is Nothing Then
            '    Session.Abandon()
            '    Response.Redirect("http://sistemas.educacao.ma.gov.br/acesso")
            'End If

        End If
    End Sub

End Class

