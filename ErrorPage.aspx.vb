Partial Class ErrorPage
    Inherits System.Web.UI.Page

    Private cEnviro As Enviro
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            PageCaption.InnerHtml = "WebAces " & ConfigurationManager.AppSettings("VersionNumber") & " " & Ajaxer.ClientName
        Catch
        End Try

        'Try
        '    LoadStyleSheets()
        'Catch
        'End Try

        litMessage.Text = Ajaxer.ErrorMessage
    End Sub

    'Private Sub LoadStyleSheets()
    '    Dim sb As New System.Text.StringBuilder
    '    Dim Skin As String

    '    Skin = Ajaxer.Skin
    '    sb.Append("<link href=""css/" & Skin & "/enroll.css"" rel=""stylesheet"" type=""text/css"" />" & Environment.NewLine)
    '    sb.Append("<link href=""css/" & Skin & "/submenu.css"" rel=""stylesheet"" type=""text/css"" />" & Environment.NewLine)
    '    sb.Append("<link href=""css/" & Skin & "/enrollment_pane_newmenu.css"" rel=""stylesheet"" type=""text/css"" />" & Environment.NewLine)
    '    sb.Append("<link href=""css/" & Skin & "/enrollment_pane_welcome.css"" rel=""stylesheet"" type=""text/css"" />" & Environment.NewLine)
    '    sb.Append("<link href=""css/" & Skin & "/footer.css"" rel=""stylesheet"" type=""text/css"" />" & Environment.NewLine)
    '    sb.Append("<link href=""css/" & Skin & "/global.css"" rel=""stylesheet"" type=""text/css"" />" & Environment.NewLine)
    '    sb.Append("<link href=""css/" & Skin & "/header.css"" rel=""stylesheet"" type=""text/css"" />" & Environment.NewLine)
    '    sb.Append("<link href=""css/" & Skin & "/main.css"" rel=""stylesheet"" type=""text/css"" />" & Environment.NewLine)
    '    sb.Append("<link href=""css/" & Skin & "/portal_content.css"" rel=""stylesheet"" type=""text/css"" />" & Environment.NewLine)
    '    sb.Append("<link href=""css/" & Skin & "/forms.css"" rel=""stylesheet"" type=""text/css"" />" & Environment.NewLine)
    '    sb.Append("<link href=""css/" & Skin & "/calendar.css"" rel=""stylesheet"" type=""text/css"" />" & Environment.NewLine)
    '    sb.Append("<link href=""css/" & Skin & "/slider.css"" rel=""stylesheet"" type=""text/css"" />" & Environment.NewLine)
    '    sb.Append("<link href=""css/" & Skin & "/error.css"" rel=""stylesheet"" type=""text/css"" />" & Environment.NewLine)
    '    litStyleSheets.Text = sb.ToString
    'End Sub
End Class
