Imports System.Data

Partial Class _Default
    Inherits System.Web.UI.Page
    Private cEnviro As Enviro

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim RequestAction As Enums.RequestActionEnum
        Dim ResponseAction As Enums.ResponseActionEnum
        Dim VersionNumber As String
        Dim Message As String = Nothing

        Try

            cEnviro = Session("Enviro")
            cEnviro.AutoDetectEnrollerID = True
            VersionNumber = cEnviro.VersionNumber()

            LoadEnviro()

            RequestAction = GetRequestAction()
            ResponseAction = ExecuteRequestAction(RequestAction, Message)
     
            litBrowserName.Text = "<input type=""hidden"" id=""hdBrowserName""  value=""" & GetBrowserName() & """ />"

            Select Case ResponseAction
                Case Enums.ResponseActionEnum.CallWebAcesAsEmp, Enums.ResponseActionEnum.CallWebAcesAsEnr
                    CallWebAces(ResponseAction)
                Case Enums.ResponseActionEnum.DisplayPage_InputUserID, Enums.ResponseActionEnum.RedisplayPage_InputUserID_FailAuth
                    DisplayPage(ResponseAction, Message)
                Case Enums.ResponseActionEnum.DisplayErrorPage_AutoDetect_FailAuth
                    Ajaxer.ErrorMessage = Message
                    Response.Redirect("ErrorPage.aspx")
            End Select

        Catch ex As Exception
            Dim ErrorHandler As New ErrorHandler(ex, "Error #CA8602: Default.Page_Load. " & ex.Message)
        End Try
    End Sub

    Private Sub LoadEnviro()
        Dim IPAddress As String
        Dim Sql As String
        Dim dt As DataTable
        Dim TrackAdmin As String
        Dim Done As Boolean

        Try

            If cEnviro.Init Then
                Exit Sub
            End If

            TrackAdmin = ConfigurationManager.AppSettings("TrackAdmin")
            If TrackAdmin.Length > 0 Then
                Select Case TrackAdmin
                    Case "32-1c28"
                        cEnviro.Platform = "DEV"
                        Done = True
                    Case "*492_cn4r"
                        cEnviro.Platform = "TEST"
                        Done = True
                End Select
            End If

            If Not Done Then
                Select Case System.Net.Dns.GetHostName().ToUpper
                    Case "WADEV"
                        cEnviro.Platform = "DEV"
                    Case "WATEST"
                        cEnviro.Platform = "TEST"
                    Case "WAPROD"
                        cEnviro.Platform = "PROD"
                End Select
            End If

            ' ___ web.config
            'cEnviro.Platform = ConfigurationManager.AppSettings("Platform").ToUpper
            cEnviro.UserCatgy = ConfigurationManager.AppSettings("UserCatgy").ToUpper
            If Common.IsNotBlank(ConfigurationManager.AppSettings("OverrideClientID")) Then
                cEnviro.ClientID = ConfigurationManager.AppSettings("OverrideClientID").ToUpper
            Else
                cEnviro.ClientID = ConfigurationManager.AppSettings("ClientID").ToUpper
            End If


            ' ___ DBHost
            'If ConfigurationManager.AppSettings("OverrideDBHost").ToUpper.Length > 0 Then
            '    cEnviro.DBHost = ConfigurationManager.AppSettings("OverrideDBHost").ToUpper
            'Else
            Select Case cEnviro.Platform
                Case Enums.Platform.DEV
                    cEnviro.DBHost = "WADEV"
                Case Enums.Platform.TEST
                    cEnviro.DBHost = "hbg-tst"
                Case Enums.Platform.PROD
                    cEnviro.DBHost = "localhost"
            End Select
            'End If

            Sql = "SELECT ClientDB, ClientName FROM ClientMaster WHERE WAClientID = '" & cEnviro.ClientID & "' AND '" & Common.GetFullServerDateTime & "' BETWEEN StartDate AND ISNULL(EndDate, '12/31/2050')"
            dt = Common.GetDTSqlXpress(Sql, cEnviro.DBHost, "WebAces", "Error #CA8603.1: Default.LoadEnviro. ")
            cEnviro.ClientDB = dt.Rows(0)(0)
            cEnviro.ClientName = dt.Rows(0)(1)
            Ajaxer.ClientName = dt.Rows(0)(1)

            ' ___ IPAddress
            Dim context As System.Web.HttpContext = System.Web.HttpContext.Current()
            Dim sIPAddress As String = context.Request.ServerVariables("HTTP_X_FORWARDED_FOR")
            If String.IsNullOrEmpty(sIPAddress) Then
                IPAddress = context.Request.ServerVariables("REMOTE_ADDR")
            Else
                Dim ipArray As String() = sIPAddress.Split(New [Char]() {","c})
                IPAddress = ipArray(0)
            End If
            cEnviro.IPAddress = IPAddress
            cEnviro.Init = True

        Catch ex As Exception
            Throw New Exception("Error #CA8603: Default.LoadEnviro. " & ex.Message)
        End Try
    End Sub

    Private Function GetBrowserName() As String
        Dim BrowserName As String

        Select Case Request.Browser.Browser.ToLower
            Case "ie", "internet explorer", "internetexplorer"
                BrowserName = "IE"
            Case "chrome"
                BrowserName = "Chrome"
            Case "firefox"
                BrowserName = "Firefox"
            Case Else
                BrowserName = Request.Browser.Browser
        End Select
        Return BrowserName
    End Function

    Private Function GetRequestAction() As Enums.RequestActionEnum
        Dim RequestAction As Enums.RequestActionEnum

        Try
            If cEnviro.UserCatgy = "EMP" Then
                RequestAction = Enums.RequestActionEnum.Emp
            Else
                If Page.IsPostBack Then
                    RequestAction = Enums.RequestActionEnum.EnrPostbackInputUserID
                Else
                    If cEnviro.AutoDetectEnrollerID Then
                        RequestAction = Enums.RequestActionEnum.EnrInitialAutoDetect
                    Else
                        RequestAction = Enums.RequestActionEnum.EnrInitialInputUserID
                    End If
                End If
            End If

            Return RequestAction

        Catch ex As Exception
            Throw New Exception("Error #CA8604: Default.GetRequestAction. " & ex.Message)
        End Try
    End Function

    Private Function ExecuteRequestAction(ByVal RequestAction As Enums.RequestActionEnum, ByRef Message As String) As Enums.ResponseActionEnum
        Dim ResponseAction As Enums.ResponseActionEnum
        Dim LoggedInUserID As String

        Try

            Select Case RequestAction
                Case Enums.RequestActionEnum.Emp
                    ResponseAction = Enums.ResponseActionEnum.CallWebAcesAsEmp
                Case Enums.RequestActionEnum.EnrInitialAutoDetect
                    LoggedInUserID = Request.ServerVariables("REMOTE_USER")
                    LoggedInUserID = LoggedInUserID.Substring(InStr(LoggedInUserID, "\", CompareMethod.Binary))
                    LoggedInUserID = LoggedInUserID.ToLower
                    If AuthenticateUser(LoggedInUserID, Message) Then
                        ResponseAction = Enums.ResponseActionEnum.CallWebAcesAsEnr
                    Else
                        ResponseAction = Enums.ResponseActionEnum.DisplayErrorPage_AutoDetect_FailAuth
                    End If

                Case Enums.RequestActionEnum.EnrInitialInputUserID
                    ResponseAction = Enums.ResponseActionEnum.DisplayPage_InputUserID


                Case Enums.RequestActionEnum.EnrPostbackInputUserID
                    If AuthenticateUser(txtLoginUserName.Text, Message) Then
                        ResponseAction = Enums.ResponseActionEnum.CallWebAcesAsEnr
                    Else
                        ResponseAction = Enums.ResponseActionEnum.RedisplayPage_InputUserID_FailAuth
                    End If
            End Select
            Return ResponseAction

        Catch ex As Exception
            Throw New Exception("Error #CA8605: Default.ExecuteRequestAction. " & ex.Message)
        End Try
    End Function

    Private Function AuthenticateUser(ByVal UserID As String, ByRef Message As String) As Boolean
        Dim dt As DataTable
        Dim Result As Boolean
        Dim sb As New System.Text.StringBuilder

        Try

            Select Case cEnviro.AutoDetectEnrollerID
                Case True
                    sb.Append("SELECT UserID, Role FROM UserManagement..Users WHERE UserID = '" & UserID & "'")
                    dt = Common.GetDTSqlXpress(sb.ToString, cEnviro.DBHost, "WebAces", "Error #CA8606.1: Default.IsUserAuthenticated. ")
                    If dt.Rows.Count = 0 Then
                        Message = "User not found."
                    Else
                        If Common.InList(dt.Rows(0)("Role"), "Enroller|Supervisor|IT", "|", True) Then
                            Result = True
                            cEnviro.EnrollerID = UserID
                        Else
                            Message = "Insufficient permissions."
                        End If
                    End If

                Case False
                    sb.Append("SELECT WAUserID FROM Permissions WHERE WAUserID = '" & cEnviro.EnrollerID & "' AND ")
                    sb.Append("Password = '" & EnrollService.EnrollDVRConnect.PassTo(txtLoginPassword.Text.Trim, cEnviro.SRACode) & "'")
                    dt = Common.GetDTSqlXpress(sb.ToString, cEnviro.DBHost, "WebAces", "Error #CA8606.1: Default.IsUserAuthenticated. ")
                    If dt.Rows.Count > 0 Then
                        Result = True
                    End If
            End Select

            Return Result
        Catch ex As Exception
            Throw New Exception("Error #CA8606: Default.IsUserAuthenticated. " & ex.Message)
        End Try
    End Function

    Private Sub DisplayPage(ByVal ResponseAction As Enums.ResponseActionEnum, ByVal Message As String)
        Dim sb As New System.Text.StringBuilder

        Try
            pnlLogin.Visible = True
            Select Case ResponseAction
                Case Enums.ResponseActionEnum.DisplayPage_InputUserID
                    litClientName.Text = cEnviro.ClientName
                Case Enums.ResponseActionEnum.RedisplayPage_InputUserID_FailAuth
                    lblLoginMessage.Text = Message
            End Select

            sb.Append("<script type=""text/javascript"">")
            sb.Append(" function DoRedirect() {document.getElementById(""txtLoginUserName"").focus()}  ")
            sb.Append("</script> ")
            litRedirect.Text = sb.ToString

        Catch ex As Exception
            Throw New Exception("Error #CA8607: Default.DisplayPage. " & ex.Message)
        End Try
    End Sub

    Private Sub CallWebAces(ByVal ResponseAction As Enums.ResponseActionEnum)
        Dim Coll As New CollX
        Dim sb As New System.Text.StringBuilder

        Try

            sb.Append("<script type=""text/javascript""> ")
            sb.Append("function DoRedirect() { ")

            Select Case cEnviro.Platform
                Case Enums.Platform.DEV
                    sb.Append(" window.location.href = ""http://seattle.dev.benenroll.com/Portal.aspx")
                Case Enums.Platform.TEST
                    sb.Append(" window.location.href = ""http://test.benenroll.com/Portal.aspx")
                Case Enums.Platform.PROD
                    sb.Append(" window.location.href = ""https://benenroll.com/Portal.aspx")
            End Select

            Coll.Assign("WAClientID", cEnviro.ClientID)
            Coll.Assign("IPAddress", cEnviro.IPAddress)
            Coll.Assign("Platform", cEnviro.Platform)
            Coll.Assign("UserCatgy", cEnviro.UserCatgy)
            If cEnviro.UserCatgy = "ENR" Then
                Coll.Assign("EnrollerID", cEnviro.EnrollerID)
            End If
            Coll.Assign("Version", cEnviro.VersionNumber)
            Coll.Assign("CallingAppBuildNum", cEnviro.CallingAppBuildNum)

            If cEnviro.Platform <> Enums.Platform.PROD Then
                If Common.IsNotBlank(Common.GetActualUserName) Then
                    Coll.Assign("ActualUserName", Common.GetActualUserName)
                Else
                    Coll.Assign("ActualUserName", Common.GetLoggedInUserID)
                End If
            End If
            sb.Append(Coll.CollxToParameters())
            sb.Append("""} ")

            sb.Append("</script>")

            litRedirect.Text = sb.ToString
        Catch ex As Exception
            Throw New Exception("Error #CA8608: Default.CallWebAces. " & ex.Message)
        End Try
    End Sub
End Class
