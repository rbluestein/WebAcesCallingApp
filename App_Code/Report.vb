Imports System.Net.Mail

Public Class Report
    Public Shared Sub SendEmail(ByVal SendTo As String, ByVal From As String, ByVal cc As String, ByVal Subject As String, ByVal TextBody As String, ByVal SuppressException As Boolean, Optional ByRef AttachmentColl As Collection = Nothing)
        Dim i As Integer
        Dim CDOConfig As CDO.Configuration
        Dim iMsg As CDO.Message = Nothing

        Try

            CDOConfig = New CDO.Configuration
            CDOConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing").Value = 2
            ' CDOConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport").Value = 25
            CDOConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver").Value = "mail.benefitvision.com"
            CDOConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate").Value = 1
            CDOConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername").Value = "automail"
            CDOConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword").Value = "$bambam2004#"
            CDOConfig.Fields.Update()

            iMsg = New CDO.Message
            iMsg.To = SendTo
            iMsg.From = From
            iMsg.CC = cc
            iMsg.Subject = Subject

            If Not AttachmentColl Is Nothing Then
                For i = 1 To AttachmentColl.Count
                    iMsg.AddAttachment(AttachmentColl(i))
                Next
            End If

            iMsg.Configuration = CDOConfig
            iMsg.TextBody = TextBody
            'imsg.HTMLBody = htmlbody

            iMsg.Send()


        Catch ex As Exception
            Throw New Exception("Error #1510: Report SendEmail. " & ex.Message)
        Finally
            iMsg.Attachments.DeleteAll()
            CDOConfig = Nothing
            iMsg = Nothing
        End Try
    End Sub
End Class
