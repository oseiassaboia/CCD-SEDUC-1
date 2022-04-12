
Partial Class ctrRelatorio
    Inherits System.Web.UI.UserControl

    Public WriteOnly Property Titulo As String
        Set(value As String)
            ViewState("Titulo") = IIf(value = "", "Relatório", value)
            lblTitulo.Text = ViewState("Titulo")
        End Set
    End Property

    Public WriteOnly Property Arquivo As String
        Set(value As String)
            ViewState("Arquivo") = value
            Session("Arquivo") = value
        End Set
    End Property

    Public Sub Show()
        'pnlRelatorio.Visible = True
        mpeModal.Show()
        'lnkVisualizar.NavigateUrl = ViewState("Arquivo")
    End Sub

    Protected Sub lbtFechar_Click(sender As Object, e As EventArgs) Handles lbtFechar.Click
        'pnlRelatorio.Visible = False
        mpeModal.Hide()
    End Sub

    Protected Sub lbtDownload_Click(sender As Object, e As EventArgs) Handles lbtDownload.Click
        If Left(ViewState("Arquivo"), 7).ToLower = "http://" Then
            Dim myWebClient As New System.Net.WebClient
            Dim CaminhoDestino As String = HttpContext.Current.Server.MapPath("Temp/" & HttpContext.Current.Session.SessionID & Now.ToString("ddMMyyyyhhmmss") & ".pdf")

            myWebClient.DownloadFile(ViewState("Arquivo"), CaminhoDestino)

            Utilidades.DownloadFile(CaminhoDestino, ViewState("Titulo") & ".pdf")
        Else
            Utilidades.DownloadFile(Server.MapPath(ViewState("Arquivo")), ViewState("Titulo") & ".pdf")
        End If
    End Sub

    Private Sub ctrRelatorio_Unload(sender As Object, e As EventArgs) Handles Me.Unload
        'pnlRelatorio.Visible = True
        'lnkVisualizar.NavigateUrl = ViewState("Arquivo")
        lblTitulo.Text = ViewState("Titulo")
    End Sub

    Protected Sub lbtVisualizar_Click(sender As Object, e As EventArgs) Handles lbtVisualizar.Click
        If Left(ViewState("Arquivo"), 7).ToLower = "http://" Then
            Dim executingPage As Object = HttpContext.Current.Handler
            ScriptManager.RegisterStartupScript(executingPage, executingPage.GetType, "Redirect", "window.location('" & HttpContext.Current.Server.MapPath("Temp/" & HttpContext.Current.Session.SessionID & Now.ToString("ddMMyyyyhhmmss") & ".pdf") & "')", True)
            'Response.Redirect(HttpContext.Current.Server.MapPath("Temp/" & HttpContext.Current.Session.SessionID & Now.ToString("ddMMyyyyhhmmss") & ".pdf"))
        Else
            Dim executingPage As Object = HttpContext.Current.Handler
            ScriptManager.RegisterStartupScript(executingPage, executingPage.GetType, "Redirect", "window.open('" & ViewState("Arquivo") & "', '_blank');", True)
            'Utilidades.DownloadFile(Server.MapPath(ViewState("Arquivo")), ViewState("Titulo") & ".pdf")

        End If
    End Sub
End Class
