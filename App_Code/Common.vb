Imports System.Data.SqlClient
Imports System.Data
Imports EnrollService

Public Class Enums
    Public Enum DBCallTypeEnum
        StoredProcedure = 1
        Sql = 2
        CommandObject = 3
    End Enum
    Public Class Platform
        Public Const DEV As String = "DEV"
        Public Const TEST As String = "TEST"
        Public Const PROD As String = "PROD"
    End Class
    Public Enum StringTreatEnum
        AsIs = 1
        SideQts = 2
        SecApost = 3
        SideQts_SecApost = 4
    End Enum
    Public Enum RequestActionEnum
        None = 0
        Emp = 1
        EnrInitialAutoDetect = 2
        EnrInitialInputUserID = 3
        EnrPostbackInputUserID = 4
    End Enum
    Public Enum ResponseActionEnum
        None = 0
        CallWebAcesAsEmp = 1
        CallWebAcesAsEnr = 2
        DisplayPage_InputUserID = 3
        DisplayErrorPage_AutoDetect_FailAuth = 4
        RedisplayPage_InputUserID_FailAuth = 5
    End Enum
End Class

Public Class Common
    Public Shared Sub GenerateError()
        Dim a, b, c As Integer
        a = b / c
    End Sub

    Public Shared Sub WriteToWADebug(ByVal TrackCode As String, ByVal Note As String, ByVal Val As String, Optional ByVal DB As String = "C3_WA2")
        Dim SessionObj As System.Web.SessionState.HttpSessionState = System.Web.HttpContext.Current.Session
        Dim mEnviro As Enviro = Nothing
        Dim dt As DataTable

        If SessionObj("Enviro") IsNot Nothing Then
            mEnviro = Common.GetSessionObj("Enviro")
        End If

        If mEnviro.WADebugColl.DoesKeyExist("TrackCode") Then
            If mEnviro.WADebugColl.DoesKeyExist("TrackCode") Then
                If InStr(Val, "''") > 0 Then
                    Val.Replace("''", "''''")
                End If
                dt = GetDTSqlXpress("INSERT INTO Debugger (TrackCode, SessionID, Note, Val, ActualUserName) VALUES ('" & TrackCode & "', '" & Note & "', '" & Val & "')", mEnviro.DBHost, "WebAces", "")
            End If
        End If
    End Sub

    Public Shared Function GetSessionObj(ByVal Name As String) As Object
        Dim SessionObj As System.Web.SessionState.HttpSessionState = System.Web.HttpContext.Current.Session
        Return SessionObj(Name)
    End Function

    Public Shared Function GetFullServerDateTime() As String
        Return Date.Now.ToUniversalTime.AddHours(-5).ToString("yyyy-MM-dd HH:mm:ss:fff")
    End Function

    Public Shared Function GetLoggedInUserID() As String
        Dim LoggedInUserID As String

        Try
            LoggedInUserID = HttpContext.Current.User.Identity.Name.ToString
            LoggedInUserID = LoggedInUserID.Substring(InStr(LoggedInUserID, "\", CompareMethod.Binary))
            LoggedInUserID = LoggedInUserID.ToLower
            Return LoggedInUserID
        Catch
            Return String.Empty
        End Try
    End Function

    Public Shared Function GetActualUserName() As String
        Dim ActualUserName As String = Nothing

        Try
            ActualUserName = HttpContext.Current.User.Identity.Name.ToString
            ActualUserName = ActualUserName.Substring(InStr(ActualUserName, "\", CompareMethod.Binary))
            Return ActualUserName.ToLower()

        Catch ex As Exception
            Throw New Exception("Error #2610: Welcome.GetActualUserName " & ex.Message)
        End Try
    End Function

    Public Shared Function StrOutHandler(ByRef Input As Object, ByVal AllowNull As Boolean, ByVal StringTreat As Enums.StringTreatEnum) As String
        Dim Output As String

        Try

            ' ___ Output, adjusting for AllowNull
            If IsDBNull(Input) Then
                If AllowNull Then
                    Output = "null"
                Else
                    Output = String.Empty
                End If
            ElseIf Input Is Nothing Then
                If AllowNull Then
                    Output = "null"
                Else
                    Output = String.Empty
                End If
            Else
                Try
                    Output = Input
                Catch
                    If AllowNull Then
                        Output = "null"
                    Else
                        Output = String.Empty
                    End If
                End Try
            End If

            ' ___ Apply string treatment
            If Output <> "null" Then
                Select Case StringTreat
                    Case Enums.StringTreatEnum.AsIs
                        ' no action
                    Case Enums.StringTreatEnum.SecApost
                        Output = Replace(Output, "'", "''")
                    Case Enums.StringTreatEnum.SideQts
                        Output = "'" & Output & "'"
                    Case Enums.StringTreatEnum.SideQts_SecApost
                        Output = Replace(Output, "'", "''")
                        Output = "'" & Output & "'"
                End Select
            End If

            Return Output

        Catch ex As Exception
            Throw New Exception("Error #2209: Common StrOutHandler. " & ex.Message, ex)
        End Try
    End Function

    Public Shared Function GetNewRecordID2(ByVal TableName As String, ByVal KeyFldName As String, ByVal MinValue As Integer, ByVal MaxValue As Int64, ByVal DBName As String, ByVal DBHost As String) As CollX
        Dim MaxNumPlaces As Integer
        Dim Factor As Int64
        Dim RandValue As Int64
        Dim dt As DataTable
        Dim NumValues As Int64
        Dim TriedColl As New CollX
        Dim RetColl As New CollX

        Try

            ' ___ What is the number of values in the specified range?
            If MaxValue - MinValue = 0 Then
                NumValues = 1
            Else
                NumValues = MaxValue - MinValue + 1
            End If

            MaxNumPlaces = MaxValue.ToString.Length
            Factor = "1".PadRight(MaxNumPlaces + 1, "0")

            Do
                Do
                    Randomize()
                    RandValue = CType(Rnd() * CInt(Factor), System.Int64)
                Loop Until (RandValue >= MinValue) AndAlso (RandValue <= MaxValue)

                ' // We found a value within the specified range

                ' ___ If the value is within the specified range and is not taken, we're out of here
                dt = GetDTSqlXpress("SELECT Count (*) FROM " & TableName & " WHERE " & KeyFldName & " = " & RandValue, DBHost, DBName, "Error #2237.1. Common.GetNewRecordID. ")

                If dt.Rows(0)(0) = 0 Then
                    RetColl.Assign("Success", True)
                    RetColl.Assign("Value", RandValue)
                    Return RetColl
                End If

                ' // That value is taken. Keep trying.

                ' ___ The count property of TriedColl keeps track of the number of values tested.
                TriedColl.Assign(RandValue)

                ' // If we have tried all of the values in the specified range, return an error

                If TriedColl.Count = NumValues Then
                    RetColl.Assign("Success", False)
                End If
            Loop Until TriedColl.Count = NumValues

            ' ___ Execution reaches here only all of the possible examples are exhausted
            Return RetColl

        Catch ex As Exception
            Throw New Exception("Error #2237. Common.GetNewRecordID2. " & ex.Message)
        End Try
    End Function

    'Public Shared Function GetConnectionString(ByVal DBHost As String, ByVal Database As String) As String
    '    Dim Text As String = "User ID=@;Password=~;database=|;server="
    '    Dim SRACode As String = "/ngOQI6g0mNaT35vriINpkM5hNm/ckMzUgIx1HFeyQg="
    '    Dim F3Sequencer As String = Nothing
    '    Dim F4Sequencer As String = Nothing

    '    Dim Platform As String
    '    Platform = ConfigurationManager.AppSettings("Platform")
    '    Select Case Platform.ToUpper
    '        Case Enums.Platform.DEV, Enums.Platform.TEST
    '            F3Sequencer = "o+0RC/MOgd7Yy6+VeU/xRg=="
    '            F4Sequencer = "YepdU5s+CTSCHBYd9Kb/Sw=="
    '        Case Enums.Platform.PROD
    '            F3Sequencer = "Bmg0KeqzLr7dr+yOO9e9MA=="
    '            F4Sequencer = "6hgd1rC7ixxg2qFW0gLqXw=="
    '    End Select

    '    Text = Replace(Text, "@", EnrollService.EnrollDVRConnect.PassBack(F3Sequencer, SRACode))
    '    Text = Replace(Text, "~", EnrollService.EnrollDVRConnect.PassBack(F4Sequencer, SRACode))
    '    Text = Replace(Text, "|", Database) & DBHost
    '    Return Text
    'End Function

    Public Shared Function GetDTSqlXpress(ByVal Sql As String, ByVal DBHost As String, ByVal Database As String, ByVal SourceInfo As String) As DataTable
        Dim QueryPack As QueryPack
        QueryPack = GetDTMaster(Sql, DBHost, Database, Nothing, Enums.DBCallTypeEnum.Sql, True, True)
        If QueryPack.Success Then
            Return QueryPack.dt
        Else
            If SourceInfo = Nothing Then
                Return Nothing
            Else
                Throw New Exception(SourceInfo & " " & QueryPack.TechErrMsg)
            End If
        End If
    End Function

    Private Shared Function GetDTMaster(ByVal Sql As String, ByVal DBHost As String, ByVal Database As String, ByVal SqlCmd As SqlCommand, ByVal DBCallType As Enums.DBCallTypeEnum, ByVal WithQuerypack As Boolean, ByVal IsXpress As Boolean) As QueryPack
        Dim DataAdapter As SqlDataAdapter
        Dim ds As New DataSet
        Dim QueryPack As New QueryPack
        Dim ErrorMessage As String = String.Empty
        Dim ConnectionString As String
        Dim Enviro As Enviro
        Dim SessionObj As System.Web.SessionState.HttpSessionState = System.Web.HttpContext.Current.Session
        Enviro = SessionObj("Enviro")

        If DBCallType <> Enums.DBCallTypeEnum.CommandObject Then
            SqlCmd = New SqlCommand(Sql)
            SqlCmd.CommandType = CommandType.Text
            ConnectionString = Enviro.GetConnectionString(Database)
            SqlCmd.Connection = New SqlConnection(ConnectionString)

        End If

        SqlCmd.CommandTimeout = 90
        DataAdapter = New SqlDataAdapter(SqlCmd)

        Try

            ErrorMessage = GetDTMaster2(DataAdapter, ds, Sql, Database, DBHost)

            If WithQuerypack Then
                If ErrorMessage.Length > 0 Then
                    QueryPack.Success = False
                    QueryPack.TechErrMsg = ErrorMessage
                Else
                    QueryPack.Success = True
                    QueryPack.NumResultSets = ds.Tables.Count
                    Select Case QueryPack.NumResultSets
                        Case 0
                        Case 1
                            QueryPack.dt = ds.Tables(0)
                            ds.Tables.RemoveAt(0)
                        Case Else
                            QueryPack.ds = ds
                    End Select
                End If
            Else
                If ErrorMessage.Length > 0 Then
                    Throw New Exception("Error #2216.1: Common ExecuteNonQueryMaster.Sql: " & Sql & " DBHost: " & DBHost & " Database: " & Database & " Error message: " & ErrorMessage)
                End If
            End If

            Return QueryPack

        Catch ex As Exception
            ErrorMessage = ex.Message
            If WithQuerypack Then
                If ErrorMessage.Length > 0 Then
                    QueryPack.Success = False
                    QueryPack.TechErrMsg = ErrorMessage
                Else
                    QueryPack.Success = True
                End If
            Else
                If ErrorMessage.Length > 0 Then
                    Throw New Exception("Error #2217: Common GetDTMaster. Sql: " & Sql & " DBHost: " & DBHost & " Database: " & Database & " Error message: " & ex.Message)
                End If
            End If
            Return QueryPack

        End Try

        DataAdapter.Dispose()
        SqlCmd.Dispose()
        SqlCmd.Connection.Close()
    End Function

    Private Shared Function GetDTMaster2(ByRef DataAdapter As SqlDataAdapter, ByRef ds As DataSet, ByVal Sql As String, ByVal Database As String, ByVal DBHost As String) As String
        Dim i As Integer
        Dim ErrorMessage As String = String.Empty

        Try
            For i = 1 To 3
                Try
                    DataAdapter.Fill(ds)
                    Exit For
                Catch ex As Exception
                    If ex.Message.Substring(0, 25) = "SQL Server does not exist" Then
                        ErrorMessage = ex.Message
                        System.Threading.Thread.Sleep(500)
                    Else
                        ErrorMessage = "Sql: " & Sql & "DBHost: " & DBHost & "~Database: " & Database & "Error message: " & ex.Message
                        Exit For
                    End If
                End Try
            Next

            Return ErrorMessage

        Catch ex As Exception
            ErrorMessage = ex.Message
            Return ErrorMessage
        End Try
    End Function

    Public Shared Function IsBlank(ByVal Value As Object) As Boolean
        If IsDBNull(Value) Then
            Return True
        ElseIf Value = Nothing Then
            Return True
        Else

            If IsNumeric(Value) Then
                Return False
            Else
                If Value.length = 0 Then
                    Return True
                Else
                    Return False
                End If
            End If
        End If
    End Function

    Public Shared Function IsNotBlank(ByVal Value As Object) As Boolean
        Return Not IsBlank(Value)
    End Function

    Public Shared Function InList(ByVal ItemSearchedFor As String, ByVal ListOfItems As String, ByVal Delimiter As String, ByVal IgnoreCase As Boolean) As Boolean
        Dim i As Integer
        Dim Box() As String
        Dim Results As Boolean

        If ListOfItems Is Nothing OrElse ListOfItems.Length = 0 Then
            ' no action
        Else

            Box = Split(ListOfItems, Delimiter)
            For i = 0 To Box.GetUpperBound(0)
                If IsStrEqual(ItemSearchedFor, Trim(Box(i)), IgnoreCase) Then
                    Results = True
                    Exit For
                End If
            Next

        End If

        Return Results
    End Function

    Public Shared Function IsStrEqual(ByVal FirstValue As String, ByVal SecondValue As String, ByVal IgnoreCase As Boolean) As Boolean
        Dim Output As Integer
        Dim Results As Boolean
        Output = String.Compare(Trim(FirstValue), Trim(SecondValue), IgnoreCase)
        If Output = 0 Then
            Results = True
        End If
        Return Results
    End Function
End Class


Public Class QueryPack
    Private cReturnDataTable As Boolean
    Private cReturnDataSet As Boolean
    Private cSuccess As Boolean
    Private cGenErrMsg As String
    Private cTechErrMsg As String
    Private cNumResultSets As Integer
    Private cdt As DataTable
    Private cds As DataSet

    Public Property Success() As Boolean
        Get
            Return cSuccess
        End Get
        Set(ByVal Value As Boolean)
            cSuccess = Value
        End Set
    End Property

    Public ReadOnly Property GenErrMsg() As String
        Get
            Return cGenErrMsg
        End Get
    End Property
    Public Property TechErrMsg() As String
        Get
            Return cTechErrMsg
        End Get
        Set(ByVal Value As String)
            cTechErrMsg = Value
        End Set
    End Property
    Public Property NumResultSets() As Integer
        Get
            Return cNumResultSets
        End Get
        Set(ByVal value As Integer)
            cNumResultSets = value
        End Set
    End Property
    Public Property dt() As DataTable
        Get
            Return cdt
        End Get
        Set(ByVal Value As DataTable)
            cdt = Value
        End Set
    End Property
    Public Property ds() As DataSet
        Get
            Return cds
        End Get
        Set(ByVal Value As DataSet)
            cds = Value
        End Set
    End Property
End Class

Public Class CollX
    Inherits System.Collections.CollectionBase
    ' Private Bittem As ListItem

    Public Sub New()
        List.Add(DBNull.Value)
    End Sub

    Public Overloads ReadOnly Property Count() As Integer
        Get
            Return List.Count - 1
        End Get
    End Property

    Default Public ReadOnly Property Coll(ByVal Idx As Integer) As Object
        Get
            Return List(Idx).Value
        End Get
    End Property

    Default Public ReadOnly Property Coll(ByVal Key As String) As Object
        Get
            Dim i As Integer
            Dim KeyUpper As String
            KeyUpper = Key.ToUpper
            For i = 1 To List.Count - 1
                If List(i).Key.ToUpper = KeyUpper Then
                    Return List(i).Value
                End If
            Next
            Throw New CollXError("Error #3604: CallX.Coll item not found error. Key: " & Key)  'CollXError
        End Get
    End Property

    Public Function TreatKeyAsString(ByVal Key As String) As Object
        Dim i As Integer
        Dim KeyUpper As String

        Try
            KeyUpper = Key.ToUpper
            For i = 1 To List.Count - 1
                If List(i).Key.ToUpper = KeyUpper Then
                    Return List(i).Value
                End If
            Next
            Throw New CollXError("Error #3611: CallX.GetValueKeyAsString item not found error. Key: " & Key)  'CollXError

        Catch ex As Exception
            Throw New CollXError("Error #3611: CallX.GetValueKeyAsString item not found error. Key: " & Key)  'CollXError
        End Try
    End Function
    Public ReadOnly Property Key(ByVal Idx As Integer) As String
        Get
            Dim i As Integer
            For i = 1 To List.Count - 1
                If i = Idx Then
                    'Return List(i).Value
                    Return List(i).Key
                End If
            Next
            Return Nothing
        End Get
    End Property
    Public ReadOnly Property DoesKeyExist(ByVal Key As String) As Boolean
        Get
            Dim i As Integer
            Key = Key.ToUpper
            For i = 1 To List.Count - 1
                If List(i).Key.ToUpper = Key Then
                    Return True
                End If
            Next
            Return False
        End Get
    End Property

    Public Sub Assign(ByVal Key As String, ByVal Value As Object)
        Dim i As Integer
        Dim Found As Boolean

        Try

            For i = 1 To List.Count - 1
                If List(i).Key = Key Then
                    List(i).Value = Nothing
                    List(i).Value = Value
                    Found = True
                End If
            Next

            'For Each Item In List
            '    If Item.Key = Key Then
            '        Item.Value = Value
            '        Found = True
            '    End If
            'Next
            If Not Found Then
                List.Add(New KeyValuePair(Key, Value))
            End If

        Catch ex As Exception
            Throw New CollXError("Error #3604: CallX.Assign. Item not found error. Key: " & Key)
        End Try
    End Sub

    Public Sub Assign(ByVal Key_Value As String)
        Dim i As Integer
        Dim Found As Boolean

        Try

            For i = 1 To List.Count - 1
                If List(i).Key = Key_Value Then
                    List(i).Value = Nothing
                    List(i).Value = Key_Value
                    Found = True
                End If
            Next

            'For Each Item In List
            '    If Item.Key = Key Then
            '        Item.Value = Value
            '        Found = True
            '    End If
            'Next
            If Not Found Then
                List.Add(New KeyValuePair(Key_Value, Key_Value))
            End If

        Catch ex As Exception
            Throw New CollXError("Error #3605: CallX.Assign. " & ex.Message)
        End Try
    End Sub

    Public Sub ConvertRow(ByRef dr As DataRow)
        Try

            Dim i As Integer
            For i = 0 To dr.ItemArray.GetUpperBound(0)
                Assign(dr.Table.Columns(i).ColumnName, dr(i))
            Next

        Catch ex As Exception
            Throw New CollXError("Error #3610: CallX.ConvertRow. " & ex.Message)
        End Try
    End Sub

    Public Function ConvertToStr(ByVal Delimiter As String) As String
        Dim i As Integer
        Dim sb As New System.Text.StringBuilder

        Try

            For i = 1 To List.Count - 1
                If i < List.Count - 1 Then
                    sb.Append(List(i).Value & Delimiter)
                Else
                    sb.Append(List(i).Value)
                End If
            Next
            Return sb.ToString
        Catch ex As Exception
            Throw New CollXError("Error #3612: CallX.ConvertToStr. " & ex.Message)
        End Try
    End Function

    Public Function CollxToSql() As String
        Dim i As Integer
        Dim sb As New System.Text.StringBuilder

        For i = 1 To List.Count - 1
            If i < List.Count - 1 Then
                sb.Append(List(i).Key & "=" & List(i).Value & ", ")
            Else
                sb.Append(List(i).Key & "=" & List(i).Value)
            End If
        Next
        Return sb.ToString
    End Function

    Public Function CollxToParameters() As String
        Dim i As Integer
        Dim sb As New System.Text.StringBuilder

        For i = 1 To List.Count - 1
            If i = 1 Then
                sb.Append("?" & List(i).Key & "=" & List(i).Value)
            Else
                sb.Append("&" & List(i).Key & "=" & List(i).Value)
            End If
        Next

        Return sb.ToString
    End Function

#Region " New from... "
    Public Shared Function NewFromList(ByVal Input As String, ByVal Delimiter As String) As CollX
        Dim i As Integer
        Dim Box As String()
        Dim Coll As New CollX
        Box = Input.Split("|")
        If Box.GetUpperBound(0) > -1 Then
            For i = 0 To Box.GetUpperBound(0)
                Coll.Assign(Box(i))
            Next
        End If
        Return Coll
    End Function

    Public Shared Function NewFromDataRow(ByRef dr As DataRow) As CollX
        Dim i As Integer
        Dim dt As DataTable
        Dim Coll As New CollX
        dt = dr.Table
        For i = 0 To dt.Columns.Count - 1
            Coll.Assign(dt.Columns(i).ColumnName, dr(i))
        Next
        Return Coll
    End Function

    Public Shared Function NewFromTable(ByRef dt As DataTable) As CollX
        Dim i As Integer
        Dim Coll As New CollX

        For i = 0 To dt.Rows.Count - 1
            Coll.Assign(dt.Rows(i)(0), dt.Rows(i)(1))
        Next
        Return Coll
    End Function


    Public Shared Function NewFromKeyValue(ByVal Input As String, ByVal RowDelimter As String, ByVal ColDelimter As String) As CollX
        Dim i As Integer
        Dim Box As String()
        Dim Box2 As String()
        Dim Coll As New CollX
        Box = Input.Split(RowDelimter)
        For i = 0 To Box.GetUpperBound(0)
            Box2 = Box(i).Split(ColDelimter)
            Coll.Assign(Box2(0), Box2(1))
        Next
        Return Coll
    End Function
#End Region

    Public Overloads Sub RemoveAt(ByVal Index As Integer)
        List.RemoveAt(Index)
    End Sub

    Public Overloads Sub Remove(ByVal Key As String)
        Dim i As Integer
        For i = 1 To List.Count - 1
            If List(i).Key.ToUpper = Key.ToUpper Then
                List.Remove(List(i))
                Exit For
            End If
        Next
    End Sub

    Public Shared Function GetDistinct(ByRef dt As DataTable, ByVal FieldName As String) As CollX
        Dim i As Integer
        Dim Coll As New CollX
        Dim FirstValue As Object = DBNull.Value

        If dt.Rows.Count = 0 Then
            Return Nothing
        Else
            Coll.Assign(dt.Rows(0)(FieldName))
            For i = 0 To dt.Rows.Count - 1
                If dt.Rows(0)(FieldName) <> Coll(Coll.Count) Then
                    Coll.Assign(dt.Rows(i)(FieldName))
                End If
            Next
        End If
        Return Coll
    End Function

    Public Function View() As String()
        Dim Output(Me.Count) As String
        Dim Val As String

        Try

            For i = 1 To List.Count - 1
                Try
                    Val = List(i).Value
                Catch ex As Exception
                    Val = "<object>"
                End Try
                Output(i) = List(i).Key & "|" & Val
            Next
            Return Output


        Catch ex As Exception
            Throw New CollXError("Error #3608: CallX.View. " & ex.Message)
        End Try
    End Function

    Public Shared Function Clone(ByVal InputColl As CollX) As CollX
        Dim i As Integer
        Dim OutputColl As New CollX

        Try

            For i = 1 To InputColl.Count
                OutputColl.Assign(InputColl.Key(i), InputColl(i))
            Next
            Return OutputColl

        Catch ex As Exception
            Throw New CollXError("Error #3609: CallX.Clone. " & ex.Message)
        End Try
    End Function

    Public Class KeyValuePair
        Private cKey As String
        Private cValue As Object

        Public Sub New(ByVal Key As String, ByVal Value As Object)
            cKey = Key
            cValue = Value
        End Sub
        Public Property Key() As String
            Get
                Return cKey
            End Get
            Set(ByVal value As String)
                cKey = value
            End Set
        End Property

        Public Property Value() As Object
            Get
                Return cValue
            End Get
            Set(ByVal value As Object)
                cValue = value
            End Set
        End Property
    End Class

    Public Class CollXError
        Inherits Exception
        Private cMessage As String

        Public Sub New(ByVal Message As String)
            cMessage = Message
        End Sub

        Public Overrides ReadOnly Property Message() As String
            Get
                Return cMessage
            End Get
        End Property

        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub
    End Class
End Class

Public Class Ajaxer
    Public Shared ClientName As String
    Public Shared ErrorMessage As String
    Public Shared AuthenticateError As Boolean
End Class