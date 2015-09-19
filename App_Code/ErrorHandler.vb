Imports Microsoft.VisualBasic

Public Class ErrorHandler
    Private cEnviro As Enviro
    Private cExceptionInEffect As Boolean

    Public Sub New(ByRef Exception As Exception, ByVal ErrorTopLevel As String)
        Dim sb As New System.Text.StringBuilder
        Dim Response As System.Web.HttpResponse = Nothing
        Dim SessionIDInd As Boolean
        Dim WriteToLogTableSuccess As Boolean
        Dim SessionObj As System.Web.SessionState.HttpSessionState = System.Web.HttpContext.Current.Session
        Dim RecID As CollX
        Dim ErrorMessageColl As CollX = Nothing

        Try
            cEnviro = SessionObj("Enviro")

            ' ___ Do not allow entry for new error if current error is being processed.
            If cEnviro.ErrorCondtionInd Then
                Exit Sub
            End If
            cEnviro.ErrorCondtionInd = True

            ' ___ Get the error message.
            Try
                ErrorMessageColl = GetErrorMessage(Exception, ErrorTopLevel)
            Catch
                ErrorMessageColl.Assign("Success", False)
            End Try

            ' ___ Ensure that a SessionID is in place.
            If cEnviro.SessionID = 0 Then
                Try
                    SessionIDInd = False
                    RecID = Common.GetNewRecordID2("ErrorLog", "SessionID", 100000, 99999999, "WebAces", cEnviro.DBHost)
                    If RecID("Success") Then
                        cEnviro.SessionID = RecID(2)
                        SessionIDInd = True
                    End If
                Catch
                    SessionIDInd = False
                End Try
            Else
                SessionIDInd = True
            End If

            ' ___ Write to WebAces..ErrorLog
            Try
                If SessionIDInd Then
                    WriteToLogTableSuccess = WriteToErrorLogTable(ErrorMessageColl("ErrorMessage"))
                End If
            Catch
            End Try

            ' ___ Send email
            If SessionIDInd Then
                If WriteToLogTableSuccess Then
                    'Report.SendEmail("rbluestein@benefitvision.com; jkleiman@benefitvision.com", "automail@benefitvision.com", Nothing, "WebAces error notice", "Information about this error may be found under SessionID " & cEnviro.SessionID.ToString & " in the WebAces error log.", False)
                    Report.SendEmail("rbluestein@benefitvision.com; jkleiman@benefitvision.com", "automail@benefitvision.com", Nothing, "WebAces error notice", "Information about this error may be found under platform " & cEnviro.Platform & ", SessionID " & cEnviro.SessionID.ToString & " in the WebAces error log.", False)
                Else
                    'Report.SendEmail("rbluestein@benefitvision.com; jkleiman@benefitvision.com", "automail@benefitvision.com", Nothing, "WebAces error notice", "Unable to write to log. Platform: " & Enviro.Platform & ", error message: " & ErrorMessage, False)
                    Report.SendEmail("rbluestein@benefitvision.com; jkleiman@benefitvision.com", "automail@benefitvision.com", Nothing, "WebAces error notice", "Unable to write to log. Platform: " & cEnviro.Platform & ", error message: " & ErrorMessageColl("ErrorMessage"), False)
                End If
            Else
                'Report.SendEmail("rbluestein@benefitvision.com; jkleiman@benefitvision.com", "automail@benefitvision.com", Nothing, "WebAces error notice", "Unable to extract error message. Platform: " & Enviro.Platform, False)
                Report.SendEmail("rbluestein@benefitvision.com; jkleiman@benefitvision.com", "automail@benefitvision.com", Nothing, "WebAces error notice", "Unable to extract SessionID. Platform: " & cEnviro.Platform, False)
            End If

            If SessionIDInd Then
                Ajaxer.ErrorMessage = "Information about this error may be found under SessionID " & cEnviro.SessionID & " in the WebAces error log."
            Else
                Ajaxer.ErrorMessage = ErrorMessageColl("ErrorMessage")
            End If

            Response = HttpContext.Current.Response
            Response.Redirect("ErrorPage.aspx")

        Catch ex As Exception
        End Try
    End Sub

    Private Function GetErrorMessage(ByRef Exception As Exception, ByVal ErrorTopLevel As String) As CollX
        Dim ErrorMessage As String = Nothing
        Dim CurException As Exception
        Dim Coll As New Collection
        Dim ErrorMessageColl As New CollX

        Try

            CurException = Exception

            While Not (CurException Is Nothing)
                Coll.Add(CurException.Message)
                CurException = CurException.InnerException
            End While

            If Coll.Count = 0 Then
                ErrorMessage = ErrorTopLevel
            Else
                If InStr(Coll(Coll.Count), "Error #", CompareMethod.Binary) > 0 Then
                    ErrorMessage = Coll(Coll.Count)
                Else
                    ErrorMessage = ErrorTopLevel
                End If
            End If

            ErrorMessageColl.Assign("Success", False)
            For i = 0 To ErrorMessage.Length - 7
                If ErrorMessage.Substring(ErrorMessage.Length - 7 - i, 7) = "Error #" Then
                    ErrorMessage = ErrorMessage.Substring(ErrorMessage.Length - 1 - i)
                    ErrorMessage = ErrorMessage.Replace("""", "*")
                    ErrorMessageColl.Assign("Success", True)
                    ErrorMessageColl.Assign("ErrorMessage", ErrorMessage)
                    Exit For
                End If
            Next
            If Not ErrorMessageColl("Success") Then
                ErrorMessageColl.Assign("ErrorMessage", "Unable to extract error message.")
            End If

            Return ErrorMessageColl

        Catch ex As Exception
            Throw New Exception("Error #2303: ErrorHandler.GetErrorMessage " & ex.Message)
        End Try
    End Function

    Private Function WriteToErrorLogTable(ByVal ErrorMessage As String) As Boolean
        Dim sb As New System.Text.StringBuilder
        Dim FullServerDateTime As String

        Try

            FullServerDateTime = Common.GetFullServerDateTime

            sb.Append("INSERT INTO ErrorLog (")
            sb.Append("SessionID, EnrollerID, ClientID, Platform, ")
            sb.Append("Source, VersionNumber, ErrorMessage, AddDate) ")

            sb.Append("VALUES (")

            sb.Append(cEnviro.SessionID.ToString & ", ")
            If cEnviro.UserCatgy = "ENR" Then
                sb.Append("'" & cEnviro.EnrollerID & "', ")
            Else
                sb.Append("null, ")
            End If
            sb.Append(Common.StrOutHandler(cEnviro.ClientDB, True, Enums.StringTreatEnum.SideQts) & ", ")
            sb.Append("'" & cEnviro.Platform & "', ")
            sb.Append("'CA_" & cEnviro.UserCatgy & "', ")
            sb.Append("'" & cEnviro.VersionNum & "', ")
            sb.Append(Common.StrOutHandler(ErrorMessage, True, Enums.StringTreatEnum.SideQts_SecApost) & ", ")
            sb.Append(Common.StrOutHandler(FullServerDateTime, False, Enums.StringTreatEnum.SideQts) & ")")

            Common.GetDTSqlXpress(sb.ToString, cEnviro.DBHost, "WebAces", "Error #2304.1: ErrorHandler.WriteToErrorLogTable ")

            Return True

        Catch ex As Exception
            Throw New Exception("Error #2304: ErrorHandler.WriteToErrorLogTable " & ex.Message)
        End Try
    End Function
End Class

