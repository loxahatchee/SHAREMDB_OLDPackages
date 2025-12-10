Attribute VB_Name = "VBAUAUTH"
'*********************************************************************
'
'Developer - The following is the profle of the Operators
'            information and Application permissions. All
'            data is available through the Class Module in
'            VBAAUTHEN.DLL.
'            All data in the DLL is expressed as STRING or
'            INTEGER with no actual lengths assigned. The
'            true length of the data being placed in the
'            strings and integers is in the comments below.
'
'User Profile Data
'********************************************************************
'Oper_ParticipantID                  Length 15
'Oper_ApplicationID                  Length 15
'Oper_WorksAtStation                 Length 3
'Oper_StationName                    Length 50
'Oper_File_Number                    Length 9
'Oper_Social_Security                Length 9
'Oper_First_Name                     Length 30
'Oper_Middle_Name                    Length 30
'Oper_Last_Name                      Length 30
'Oper_Suffix                         Length 3
'Oper_Job_Title                      Length 50
'Oper_Org_Code                       Length 50
'Oper_UnitArea_Location              Length 50
'Oper_Phone_Number_Area              Length 4
'Oper_Phone_Number                   Length 11
'Oper_Phone_Extension                Length 4
'Oper_Security_Officer               Length 1    "Y" or "N"
'Oper_Approve_Request                Length 50
'Oper_Application_Role               Length 50
'
'Oper_Number_POAS contains the number of POAS in the TBLPOAS string
'The limit is 20
'Oper_Number_POAS                    Length 2
'***Oper_TBLPOAS contains all POAS assigned to this User if the
'***User is a VSO and consists of the following:
'****POA Code                        Length 3
'Oper_TBLPOAS                        Length 60
'
'Oper_Access_Level                   Length 1
'Oper_BDN_Badge                      Length 4
'Oper_Security_Header                Length 30
'
'***Number_Operations contains the number of Operations in
'***Oper_TBLApplication_Operation  - Limit is 270
'Oper_Number_Operations              Length 3
'***Application Operation consists of the following:
'****Operation Title Text            Length 25
'****Operation Disabled              Length 1
'****Operation Assigned Value        Length 12
'****Operation ID                    Length 15
'
'Oper_TBLApplication_Operation       Length up to 14310
'Oper_Diagnostic_Suppression         Length 1
'Oper_Email_Address                  Length 100
'Oper_App_OutBased                   Length 1
'Oper_LocationID                     Length 15
'Oper_Jurisdiction_Station           Length 3
'Oper_Jurisdiction_ID                Length 15
'Oper_WEB_App_URL                    Length 255
'****************************************************************
'
'
'The following Access data is used/provided when Validating a User
'for another application while still in the current application
'or for Validating another User other than the one that is
'currently logged on for an application. A complete profile will
'be supplied for the User ID and application supplied.
'
'****************************************************************
'User Access Profile Data
'********************************************************************
'Access_ParticipantID                  Length 15
'Access_ApplicationID                  Length 15
'Access_WorksAtStation                 Length 3
'Access_StationName                    Length 50
'Access_File_Number                    Length 9
'Access_Social_Security                Length 9
'Access_First_Name                     Length 30
'Access_Middle_Name                    Length 30
'Access_Last_Name                      Length 30
'Access_Suffix                         Length 3
'Access_Job_Title                      Length 50
'Access_Org_Code                       Length 50
'Access_UnitArea_Location              Length 50
'Access_Phone_Number_Area              Length 4
'Access_Phone_Number                   Length 11
'Access_Phone_Extension                Length 4
'Access_Security_Officer               Length 1    "Y" or "N"
'Access_Approve_Request                Length 50
'Access_Application_Role               Length 50
'
'Access_Number_POAS contains the number of POAS in the TBLPOAS string
'The limit is 20
'Access_Number_POAS                    Length 2
'***Access_TBLPOAS contains all POAS assigned to this User if the
'***User is a VSO and consists of the following:
'****POA Code                          Length 3
'Access_TBLPOAS                        Length 60
'
'Access_Access_Level                   Length 1
'Access_BDN_Badge                      Length 4
'Access_Security_Header                Length 30
'
'***Number_Operations contains the number of Operations in
'***Access_TBLApplication_Operation  - Limit is 270
'Access_Number_Operations              Length 3
'***Application Operation consists of the following:
'****Operation Title Text              Length 25
'****Operation Disabled                Length 1
'****Operation Assigned Value          Length 12
'****Operation ID                      Length 15
'
'Access_TBLApplication_Operation       Length up to 14310
'Access_Diagnostic_Suppression         Length 1
'Access_Email_Address                  Length 100
'Access_App_OutBased                   Length 1
'Access_LocationID                     Length 15
'Access_Jurisdiction_Station           Length 3
'Access_Jurisdiction_ID                Length 15
'Access_WEB_App_URL                    Length 255
'****************************************************************

'********
'In addition VSO_Indication will be set to a "Y" or "N" based
'on the presence of POA codes
'********
'
'
'
'*********************************************************************
'
'Developer - If your application already uses a Sub Main
'            please rename your Sub Main and call it from the
'            appropriate place in the code below.
'
'*********************************************************************

Option Explicit
Public StationCount As Integer
Private Stations As String
Public tblROS() As String
'Declare function for self registration
Public Const GW_HWNDPREV = 3
Declare Function OpenIcon Lib "user32" (ByVal hwnd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Dim OldTitle As String
Dim PrevHndl As Long
Dim result As Long
Public CommandParms As New Collection
Public cs As VBAATHENAD.VA_TUX


Sub Main()

'Use the Common Services DLL to Logon to Tuxedo

    On Error GoTo Err_DLL_Not_Registered
    Set cs = New VA_TUX
    
'Commandline is used when Application is being shelled and uses _
'the DLL and another presentation of the Logon Screen is not desired._
'The parameter will contain the necessary parameter to pass authentication.
    
    Dim CommandLine As String
    Dim CmdLineLen As Integer
    Dim FileExists As String
    Dim RegMyDLLAttempted As Boolean
    CommandLine = Command()
    StatusFile = App.Path & "\StatusFile.txt"
    FileExists = Dir(StatusFile)
    If FileExists <> "" Then
        Open StatusFile For Append As #1
    Else
        Open StatusFile For Output As #1
    End If
    CmdLineLen = Len(CommandLine)
    
'*********************************************************************
'Developer - Supply the name of your application in place of
'            "XXXXXXXX". This should be the same name used when
'            building the Application Name in the Corporate Database.
'                Application name is a string * 30.
'*********************************************************************
    With cs
        .Function_Type = "WEBLOGON"
        .WEBPassword = "Austin1!"
        .WEBUser = "281LSMIT"
        .WEBNewPassword = ""
        .WEBVerifyPassword = ""
        .Application_Name = "SHAREMDBCREATESRVR"
        .WEBStation = "281"
        'Initialize Tuxedo and perform User Authentication
        .Return_Code = "OK"
        .Validate_Version = Format(App.Major, "0000") & "." & Format(App.Minor, "0000") & "." & Format(App.Revision, "0000")
       
        'Call the Security Method passing the
        'Function Type and Application Name
        
        .CSSTUX_Function
      
        'If the return from Authentication was unsuccessful - terminate
        
        If cs.Return_Code = "FATAL" Then
            .Function_Type = "TERMINATE"
            .CSSTUX_Function
            End
        Else
    '*********************************************************************
    'Developer - If you are using TPCLNT32.BAS in your project
    '            then the next line will set the values in TPCLNT32.BAS.
    '            If you are not using TPCLNT32.BAS then use CSSecurityDLL2.BAS.
    '            It includes everything but the TPCLIENT setup.
    '            The CSSecurityDLL2.BAS can be used for applications not
    '            using TPCLNT32.BAS.
    '*********************************************************************
            CSS_Return_SetUp
            On Error GoTo Err_In_Application
    
Testing_Skip:
    '*********************************************************************
    'Developer - Place the name of your first form here.
    '            You determine if Modal or Non-Modal.
    '            If you had a Sub Main in your application
    '            then you can rename your Sub Main and call
    '            the procedure here instead of placing the name
    '            of your first form here.
    '*********************************************************************
            
    '        EX: frmTest.Show, frmTest.Show 1, Your_Main
            ShareMDB_Main
        End If
    End With
On Error GoTo 0
    Exit Sub

Err_DLL_Not_Registered:
    
    If Err.Number = 430 Then
         StatusMsg = "1. This application has not been compiled against the version of the DLL that allows Binary Compatability." _
         & vbCrLf & "2. The application was compiled against an older version of the DLL than the one installed on this PC. " _
         & vbCrLf & "3. The registry is incorrect and the VBAAUTHEN.DLL should be un-installed and re-installed." _
         & vbCrLf & "Error is " & Err.Number & "." _
         & vbCrLf & "Description is - " & Err.Description
         Print #1, StatusMsg
    ElseIf Err.Number = 429 Then '19471
        StatusMsg = "Unexpected fatal error in VBAAUTHEN DLL caused by the launched Application. " & vbCrLf & "Error is " & Err.Number & "." _
        & vbCrLf & "Description is - " & Err.Description & vbCrLf & vbCrLf & "This is not a Common Security Error!!!" & vbCrLf _
        & "There is an OCX or DLL on this PC that is incompatible with the application being launched." & vbCrLf _
        & "One of three things can be done." & vbCrLf _
        & " 1. Re-install the application." & vbCrLf _
        & " 2. Determine the OCX or DLL in error through registry compares." & vbCrLf _
        & " 3. Re-stage the PC."
        Print #1, StatusMsg
    ElseIf Err.Number = 75 Or Err.Number = 52 Then
        StatusMsg = "Unexpected fatal error in VBAAUTHEN DLL call. " & vbCrLf & "Error is " & Err.Number & "." _
        & vbCrLf & "Description is - " & Err.Description & vbCrLf _
        & "Most likely reasons are:" & vbCrLf _
        & " 1. The user does not have Read, Write, Execute permissions for the TUX\ULOG directory" & vbCrLf _
        & " 2. The WsEnvFile is missing or the environment variable is incorrect." & vbCrLf _
        & " 3. The Stationfile.txt is not present or incorrect permissions on the Public Drive under VBAAUTHEN"
        Print #1, StatusMsg
    Else
        StatusMsg = "Unexpected fatal error in VBAAUTHEN DLL call. " & vbCrLf & "Error is " & Err.Number & "." _
        & vbCrLf & "Description is - " & Err.Description
        Print #1, StatusMsg
    End If
    
    Exit Sub
Err_In_Application:
    If Err.Number = 429 Then  '19471
        StatusMsg = "Unexpected fatal error in VBAAUTHEN DLL caused by the launched Application. " & vbCrLf & "Error is " & Err.Number & "." _
        & vbCrLf & "Description is - " & Err.Description & vbCrLf & vbCrLf & "This is not a Common Security Error!!!" & vbCrLf _
        & "There is an OCX or DLL on this PC that is incompatible with the application being launched." & vbCrLf _
        & "One of three things can be done." & vbCrLf _
        & " 1. Re-install the application." & vbCrLf _
        & " 2. Determine the OCX or DLL in error through registry compares." & vbCrLf _
        & " 3. Re-stage the PC."
        Print #1, StatusMsg
    ElseIf Err.Number = 75 Or Err.Number = 52 Then
        StatusMsg = "Unexpected fatal error in VBAAUTHEN DLL call. " & vbCrLf & "Error is " & Err.Number & "." _
        & vbCrLf & "Description is - " & Err.Description & vbCrLf _
        & "Most likely reasons are:" & vbCrLf _
        & " 1. The user does not have Read, Write, Execute permissions for the TUX\ULOG directory" & vbCrLf _
        & " 2. The WsEnvFile is missing or the environment variable is incorrect." & vbCrLf _
        & " 3. The Stationfile.txt is not present or incorrect permissions on the Public Drive under VBAAUTHEN"
        Print #1, StatusMsg
    Else
        StatusMsg = "Unexpected fatal error in " & cs.Application_Name & ". " & vbCrLf & "Error is " & Err.Number & "." _
            & vbCrLf & "Description is - " & Err.Description
        Print #1, StatusMsg
    End If
End Sub

Function CSS_TermRestart(ByVal TuxFunction)
    
    'The Termination of Tuxedo should be the last
    'thing done by an application before issuing and END statement
    'Its purpose is to terminate the Tuxedo thread normally but
    'the return code is not checked here because this is a normal
    'application termination process.
    
    'Initialize return code and Tuxedo function type call
    With cs
        .Return_Code = "OK"
    
        'Set the function type to the function passed by the _
        'calling application
    
        .Function_Type = TuxFunction
        
        'Call the Security Function
        .CSSTUX_Function
    End With
    
    If TuxFunction = "RESTART" Then
        CSS_Return_SetUp
    End If
    
End Function

Sub CSS_Return_SetUp()

    'Put the data in the TPClient returned
    'from the Common Services Security DLL
    With TP
        .Client_Tuxedo_Function = cs.Client_Tuxedo_Function
        .Client_TP_Status_Code = cs.Client_TP_Status_Code
        .Client_Appl_Data_Send_Len = cs.Client_Appl_Data_Send_Len
        .Client_Appl_Data_Recv_Len = cs.Client_Appl_Data_Recv_Len
        .Client_Client_Module_Name = cs.Client_Client_Module_Name
        .Client_Service_Name = cs.Client_Service_Name
        .Client_Route_Station = cs.Client_Route_Station
        .Client_Client_User_Name = cs.Client_Client_User_Name
        .Client_Client_Computer_Name = cs.Client_Client_Computer_Name
        .Client_TP_Environment_Ind = cs.Client_TP_Environment_Ind
        .Client_Conv_Control_Ind = cs.Client_Conv_Control_Ind
        .Client_Multi_Context_Addr = cs.Client_Multi_Context_Addr
    End With
         
End Sub

Public Sub CSS_Build_RO_Table()
     'After calling the following function to retrieve all RO's
    'and place them in a string named cs.TBLStationROS
    '
    'Call the Security Method passing the Function Type
    'Example: To build an RO table(tblROS)
    '   cs.Function_Type = "STATIONROS"
    '   cs.CSSTUX_Function
    
    '   If you wish all stations in corporate excluding NARA's the the function
    '   would be "STATIONALL"
    '   If you wish all stations in corporate including NARA's the the function
    '   would be "STATIONALN"
    '   If you wish all stations in corporate other than RO's and excluding NARA's
    '   the function would be "STATIONOTH"
    '   If you wish all stations in corporate that are only RO's and includes development
    '   sites then the function would be "STATIONROS"
    '   If you wish all National Archives Centers in corporate that are only RO's the
    '   function would be "STATIONARA"
    '   If you wish all stations in corporate that are only RO's and excludes development
    '   sites then the function would be "STATIONROO"
    
    'Developers can then call this procedure to build a table of station
    'numbers and station names from the string of stations in cs.TBLStationROS.
    'The format will be 3 digit station number followed by 50 character station
    'name and it will place them in tblROS. cs.TBLStationROSCnt contains the
    'count of stations in the table.
    
    With cs
        ReDim tblROS(.TBLStationROSCnt)
        Stations = .TBLStationROS
        For StationCount = 1 To .TBLStationROSCnt
            tblROS(StationCount) = Left(Stations, 53)
            If StationCount <> .TBLStationROSCnt Then
                Stations = Right(Stations, Len(Stations) - 53)
            End If
        Next StationCount
    End With
    
End Sub

Public Sub CSS_Change_Station()
    'Developers can call this function to change their station for
    'testing purposes. CS_Return_Setup will place the new values
    'in TP.Client
    With cs
        .Return_Code = "OK"
        
        .Function_Type = "STATIONCHG"
        
        .CSSTUX_Function
        If .Return_Code <> "FATAL" Then
            CSS_Return_SetUp
        End If
    End With
End Sub

Sub Extended_Check_For_Previous_Instance()

   
    'Save the title of the application.
    OldTitle = App.Title
    
    'Rename the title of this application so FindWindow
    'will not find this application instance.
    App.Title = "unwanted instance"

    'Attempt to get window handle using VB4 class name.
    PrevHndl = FindWindow("ThunderRTMain", OldTitle)

   'Check for no success.
    If PrevHndl = 0 Then
        PrevHndl = FindWindow("wndclass_desked_gsk", OldTitle)
        If PrevHndl = 0 Then
            'Attempt to get window handle using VB5 class name.
            PrevHndl = FindWindow("ThunderRT5Main", OldTitle)
            If PrevHndl = 0 Then
                'Attempt to get window handle using VB6 class name
                PrevHndl = FindWindow("ThunderRT6Main", OldTitle)
            End If
        End If
    End If
    
    'Check if running
    
    If PrevHndl = 0 Then
       'No previous instance found.
       App.Title = OldTitle
       Exit Sub
    Else
        StatusMsg = "Application " & cs.Application_Name & " is already running." & vbCrLf & "Processing is terminated."
        Print #1, StatusMsg
        End
    End If
    
    'Future Posible Code
    'Get handle to previous window.
    'PrevHndl = GetWindow(PrevHndl, GW_HWNDPREV)

    'Restore the program.
    'result = OpenIcon(PrevHndl)

    'Activate the application.
    'result = SetForegroundWindow(PrevHndl)

    'End the application.
    'End

End Sub
Private Function ParseCommand(ByRef fncCommandLine, ByRef fncCommandParms)
    Dim Where As Integer
    Dim StartAt As Integer
    Dim a As Integer
    Dim AddToLine As String
    Dim MoreParms As Boolean
    MoreParms = True
    a = 1
    StartAt = 1
    Do While MoreParms = True
        Where = InStr(StartAt, Trim(fncCommandLine), " ")
        If Where > 0 Then
            AddToLine = Mid(fncCommandLine, StartAt, Where - StartAt)
            fncCommandParms.Add (AddToLine)
            a = a + 1
            StartAt = Where + 1
        Else
            If StartAt < Len(Trim(fncCommandLine)) Then
                AddToLine = Mid(fncCommandLine, StartAt, (Len(Trim(fncCommandLine)) - StartAt) + 1)
                fncCommandParms.Add (AddToLine)
            End If
            MoreParms = False
        End If
    Loop
        
End Function
