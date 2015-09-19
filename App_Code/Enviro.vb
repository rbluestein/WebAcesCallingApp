Imports Microsoft.VisualBasic

Public Class Enviro
    Private cAutoDetectEnrollerID As Boolean
    Private cPlatform As String
    Private cUserCatgy As String
    Private cClientID As String
    Private cOverrideWebAcesDB As String
    Private cInit As Boolean
    Private cSessionID As Integer
    Private cEnrollerID As String
    Private cIPAddress As String
    Private cDBHost As String
    Private cClientDB As String
    Private cClientName As String
    Private cErrorConditionInd As Boolean
    Private cSRACode As String = "/ngOQI6g0mNaT35vriINpkM5hNm/ckMzUgIx1HFeyQg="
    Private cWADebugColl As New CollX

    Public Sub New()
        cWADebugColl.Assign("Sun", False)
        cWADebugColl.Assign("MeKickedOff", False)
        cWADebugColl.Assign("TrackActivityID", False)
        cWADebugColl.Assign("SlidScalPlanBenAmt", False)
        cWADebugColl.Assign("LoginEnr", False)
        cWADebugColl.Assign("AnotherEnr", True)
    End Sub

    Public Function GetConnectionString(ByVal Database As String) As String
        Dim Text As String = "User ID=@;Password=~;database=|;server="
        Dim SRACode As String = "/ngOQI6g0mNaT35vriINpkM5hNm/ckMzUgIx1HFeyQg="
        Dim F3Sequencer As String = Nothing
        Dim F4Sequencer As String = Nothing

        Select Case cPlatform
            Case Enums.Platform.DEV, Enums.Platform.TEST
                F3Sequencer = "o+0RC/MOgd7Yy6+VeU/xRg=="
                F4Sequencer = "YepdU5s+CTSCHBYd9Kb/Sw=="
            Case Enums.Platform.PROD
                F3Sequencer = "Bmg0KeqzLr7dr+yOO9e9MA=="
                F4Sequencer = "6hgd1rC7ixxg2qFW0gLqXw=="
        End Select

        Text = Replace(Text, "@", EnrollService.EnrollDVRConnect.PassBack(F3Sequencer, SRACode))
        Text = Replace(Text, "~", EnrollService.EnrollDVRConnect.PassBack(F4Sequencer, SRACode))
        Text = Replace(Text, "|", Database) & cDBHost
        Return Text
    End Function

    Public Property AutoDetectEnrollerID As Boolean
        Get
            Return cAutoDetectEnrollerID
        End Get
        Set(value As Boolean)
            cAutoDetectEnrollerID = value
        End Set
    End Property

    Public Property Platform() As String
        Get
            Return cPlatform
        End Get
        Set(ByVal value As String)
            cPlatform = value
        End Set
    End Property

    Public Property UserCatgy() As String
        Get
            Return cUserCatgy
        End Get
        Set(ByVal value As String)
            cUserCatgy = value
        End Set
    End Property

    Public Property ClientID() As String
        Get
            Return cClientID
        End Get
        Set(ByVal value As String)
            cClientID = value
        End Set
    End Property
    Public Property WADebugColl() As CollX
        Get
            Return cWADebugColl
        End Get
        Set(ByVal value As CollX)
            cWADebugColl = value
        End Set
    End Property

    'Public Property OverrideDBHost() As String
    '    Get
    '        Return cOverrideDBHost
    '    End Get
    '    Set(ByVal value As String)
    '        cOverrideDBHost = value
    '    End Set
    'End Property


    Public ReadOnly Property VersionNum() As String
        Get
            Dim Version As String
            'Version = "0.02.31"  ' 11/06/2013: Provided for different userids and passwords across servers; set up version number; added OverrideDBHost to web.config
            'Version = "0.02.32"  ' 11/25/2013: Modified urls to match changed url names.
            'Version = "0.02.33"  ' 11/26/2013: Expose version number to default page.
            'Version = "0.02.34"  ' 11/26/2013: Adjust urls to reflect IIS changes.
            'Version = "0.02.35"  ' 12/16/2013: Rewrote code building the url, using CollxToParameters.
            'Version = "0.02.36"  ' 12/18/2013: Refined Platform. Switched test url to watest.
            'Version = "0.02.37"  ' 01/05/2014: Switched to Zalan skin. Converted DoRedirect dynamic for emp login.
            'Version = "0.02.38"  ' 01/09/2014: Added code to account for hbg-tst database names. 
            'Version = "0.02.39"  ' 01/09/2014:  Added Enviro.ClientDB. Switched from web.config to ClientMaster table for clientdb. 
            'Version = "0.02.40"  ' 02/25/2014: Set up WAClientID to enable aliases on same database. 
            'Version = "0.02.41"  ' 03/31/2014: No change. Precaution for publish to watest.
            'Version = "0.02.42"  ' 09/17/2015: Get actual user name for dev and test."
            Version = "0.02.43"   ' 09/18/2015: Auto detect environment with override."
            Return Version
        End Get
    End Property

    Public ReadOnly Property CallingAppBuildNum As String
        Get
            Dim BuildNum As String
            BuildNum = "108473" ' 08/13/2015
            Return BuildNum
        End Get
    End Property


    Public Property Init() As Boolean
        Get
            Return cInit
        End Get
        Set(ByVal value As Boolean)
            cInit = value
        End Set
    End Property

    Public Property SessionID() As Integer
        Get
            Return cSessionID
        End Get
        Set(ByVal value As Integer)
            cSessionID = value
        End Set
    End Property

    Public Property EnrollerID() As String
        Get
            Return cEnrollerID
        End Get
        Set(ByVal value As String)
            cEnrollerID = value
        End Set
    End Property

    Public Property IPAddress() As String
        Get
            Return cIPAddress
        End Get
        Set(ByVal value As String)
            cIPAddress = value
        End Set
    End Property

    Public Property DBHost() As String
        Get
            Return cDBHost
        End Get
        Set(ByVal value As String)
            cDBHost = value
        End Set
    End Property



    Public Property ClientDB() As String
        Get
            Return cClientDB
        End Get
        Set(ByVal value As String)
            cClientDB = value
        End Set
    End Property

    Public Property ClientName() As String
        Get
            Return cClientName
        End Get
        Set(ByVal value As String)
            cClientName = value
        End Set
    End Property



    Public Property ErrorCondtionInd() As Boolean
        Get
            Return cErrorConditionInd
        End Get
        Set(ByVal value As Boolean)
            cErrorConditionInd = value
        End Set
    End Property

    'Public Property DisplayErrorMessage() As Boolean
    '    Get
    '        Return cDisplayErrorMessage
    '    End Get
    '    Set(ByVal value As Boolean)
    '        cDisplayErrorMessage = value
    '    End Set
    'End Property

    Public ReadOnly Property SRACode() As String
        Get
            Return cSRACode
        End Get
    End Property
End Class
