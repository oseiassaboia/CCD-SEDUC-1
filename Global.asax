<%@ Application Language="VB" %>

<script runat="server">

    Sub Application_Start(ByVal sender As Object, ByVal e As EventArgs)
        ' Code that runs on application startup
        Application("Count") = 0
    End Sub

    Sub Application_End(ByVal sender As Object, ByVal e As EventArgs)
        ' Code that runs on application shutdown
    End Sub

    Sub Application_Error(ByVal sender As Object, ByVal e As EventArgs)
        ' Code that runs when an unhandled error occurs
        'Application("Count") = Application("Count")- 1
        'Response.Redirect("sistemas.educacao.ma.gov.br/acesso")

    End Sub

    Sub Session_Start(ByVal sender As Object, ByVal e As EventArgs)
        ' Code that runs when a new session is started
        Application.Lock()
        Application("Count") = Convert.ToInt32(Application("Count")) + 1
        Application.UnLock()
    End Sub

    Sub Session_End(ByVal sender As Object, ByVal e As EventArgs)
        ' Code that runs when a session ends. 
        ' Note: The Session_End event is raised only when the sessionstate mode
        ' is set to InProc in the Web.config file. If session mode is set to StateServer 
        ' or SQLServer, the event is not raised.
        Application.Lock()
        Application("Count") = Convert.ToInt32(Application("Count")) - 1
        Application.UnLock()
    End Sub



    Protected Sub Page_Init(sender As Object, e As EventArgs)
        If Context.Session IsNot Nothing Then
            If Session.IsNewSession Then
                Dim newSessionIdCookie As HttpCookie = Request.Cookies("ASP.NET_SessionId")
                If newSessionIdCookie IsNot Nothing Then
                    Dim newSessionIdCookieValue As String = newSessionIdCookie.Value
                    If newSessionIdCookieValue <> String.Empty Then
                        ' This means Session was timed Out and New Session was started
                        Response.Redirect("sistemas.educacao.ma.gov.br/acesso")
                    End If
                End If
            End If
        End If
    End Sub

</script>