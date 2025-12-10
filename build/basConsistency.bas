Attribute VB_Name = "basConsistency"
'VETNT00018229 DSG New EPC Consistency module 3/08/08 DSG
Global SendRec As Out_User_Record
Global ReplyRec As gt_User_Record
Type gt_User_Record
    Data As String * 26000
End Type
Type Out_User_Record
    Data As String * 6000
End Type


'VETNT00018229 DSG Breakdown gblHoldProcessEPCs
'     Benefit Type Code 12 bytes such as "010LCOMP"
'     Benefit Type Name 50 bytes
Global gblHoldProcessEPCs() As String * 62
Global gblEPCcount As Integer
Global StatusFile As String
Global StatusMsg As String
'VETNT00018229 DSG Breakdown gblSortProcessEPCs
'     Benefit Type Code 3 bytes such as "010"
'     Filler of " - " for 3 bytes
'     Benefit Type Name 50 bytes
'     Rest of Benefit Type Code 9 bytes such as "LCOMP"
Global gblSortProcessEPCs() As String * 65
Global gblSortEPCcount As Integer


'VETNT00018229 DSG Breakdown gblIncrementTable
'     Parent EPC 4 bytes
'     Actual EPC 4 bytes
'     Increment to EPC 4 bytes
'     Type Single 1 byte "Y" or "N" - "Y" not incrementable - "N" incrementable
'     Benefit Type Code 12 bytes - such as "CPL"
'     Form Name 50 bytes - such as frmShareSearch
Global gblIncrementTable() As String * 75
Global gblIncrementCount As Integer
Global gblIncCount As Integer

'VETNT00018229 DSG Breakdown gblHoldConsistency
'    Command index 2 byte
'    Benefit index 2 byte
'    Payee Range index 2 bytes
'    EPC Code index 3 bytes
'    Effective Date 8 bytes - not used for Share Processing
'    Disable Date 8 bytes - not used for Share Processing
Global gblHoldConsistency() As String * 25
Global gblConsistencyCount As Long
Global gblHaveConsistency As Boolean

'VETNT00018229 DSG Breakdown gblHoldCommands
'    Command BDN representation 12 bytes such as "CEST" or "PCLR"
'    Command description 30 bytes
Global gblHoldCommands() As String * 42
Global gblCommandsCount As Integer
    
'VETNT00018229 DSG Breakdown gblHoldBenefitCodes
'     Benefit codes 12 bytes such as "CPL" or "CPD"
'     Benefit Description 50 bytes such as "Compensation and Pension Live"
Global gblHoldBenefitCodes() As String * 62
Global gblBenefitsCount As Integer

'Breakdown gblCommandBenefit for holding Benefits for a specific command
'    Benefit Type Codes 12 bytes
'    Benefit Description 50 bytes
Global gblCommandBenefit() As String * 62
Global gblCmdBeneCount As Integer


'VETNT00018229 DSG Breakdown gblHoldPayees
'    Payee Ranges 50 bytes
Global gblHoldPayeeRanges() As String * 50
Global gblPayeeCount As Integer
Global dbEPC As Database
Global DatabasePath As String
Global fCriteria As String
Global Benefit_Type_Table As Recordset
Global Refresh_Date_Table As Recordset
Global gblDBContensionError As Boolean
Dim MoreRecords As String * 1
Dim EPCPageCount As Integer
Dim inArray2() As String
Dim SwapValues As Boolean
Dim EPCCount As Integer
Dim IncPageCount As Integer
Dim IncCount As Integer

Public Sub ShareMDB_Main()
    Dim CheckDateDB As String * 10
    Dim result As String
    Dim FileExists As String
    
    TP.Client_Tuxedo_Function = 98
    TP.Client_Client_Module_Name = "CSSDLL"
    TP.Client_Appl_Data_Send_Len = 16
    TP.Client_Appl_Data_Recv_Len = 5000
    CSS_InitializeTuxedo

    CheckDateDB = ""
    DatabasePath = App.Path & "\Consistency.MDB"
    FileExists = Dir(DatabasePath)
    If Trim(FileExists) = "" Then
        StatusMsg = "The Share Consistency MDB does not exist - run terminating!"
        Print #1, StatusMsg
        End
    End If
    Set dbEPC = OpenDatabase(DatabasePath, False, True, ";pwd=consistdsg")
    fCriteria = "Select * FROM Refresh_Date_Table"
    Set Refresh_Date_Table = dbEPC.OpenRecordset(fCriteria)
    If Not Refresh_Date_Table.EOF Then
        With Refresh_Date_Table
            .MoveFirst
            CheckDateDB = !RefreshDate
        End With
        Refresh_Date_Table.Close
    End If
    dbEPC.Close
    result = Clear_And_Reload_Consistency_From_MDB
    End
End Sub
Public Sub Get_Data_For_Consistency()
'VETNT00018229 DSG Get all the data required for consistency checks
'This procedure is called from the Splash Screen and Global Tables loaded
    Get_Consistency_Table
    Get_Increment
    Get_All_EPCs
    Get_Commands
End Sub

Public Sub Get_Increment()
'VETNT00018229 DSG Get the Incrementals for the End Product Codes and load
'them into a table
'    SendRec.Data = ""
'    TP.Client_Service_Name = "shrinqepc"
'    TP.Client_Appl_Data_Send_Len = 6
'    TP.Client_Appl_Data_Recv_Len = 26000
'    SendRec.Data = "GETINC"
'    DoEvents
'    SHARE_CallTuxedo
'    If gReturnCode <> "FATAL" Then
'        Load_Increment
'    End If
'    If gblIncrementCount = 0 Then
'        StatusMsg = "The Increment Count was 0 - there is no data on the Corporate Database!"
'        Print #1, StatusMsg
'        End
'    End If

    Dim b As Integer
    
    IncCount = 0
    IncPageCount = 0
    b = 1
    
    SendRec.Data = ""
    
    MoreRecords = "Y"
    Do Until MoreRecords = "N"
        IncPageCount = IncPageCount + 1
        TP.Client_Service_Name = "shrinqepc"
        TP.Client_Appl_Data_Send_Len = 9
        TP.Client_Appl_Data_Recv_Len = 26000
        SendRec.Data = "GETINC" & Format(IncPageCount, "000")
        DoEvents
        SHARE_CallTuxedo
        If gReturnCode <> "FATAL" Then
            If IncPageCount > 1 Then
                b = b + 300
            End If
            
            Load_Increment (b)
        Else
            Exit Sub
        End If
    Loop
End Sub
Public Sub Load_Increment(b As Integer)
'VETNT00018229 DSG Load the Incrementals to the table
'    If Left(ReplyRec.Data, 9) = "SECU00000" Then
'        ReplyRec.Data = Right(ReplyRec.Data, Len(ReplyRec.Data) - 9)
'        gblIncrementCount = Trim(Left(ReplyRec.Data, 3))
'        ReplyRec.Data = Right(ReplyRec.Data, Len(ReplyRec.Data) - 3)
'        ReDim gblIncrementTable(gblIncrementCount + 1)
'        For a = 1 To gblIncrementCount
'            gblIncrementTable(a) = Left(ReplyRec.Data, 75)
'            ReplyRec.Data = Right(ReplyRec.Data, Len(ReplyRec.Data) - 75)
'        Next a
'    End If

    gblIncCount = 0
    
    If Left(ReplyRec.Data, 9) = "SECU00000" Then
        ReplyRec.Data = Right(ReplyRec.Data, Len(ReplyRec.Data) - 9)
        MoreRecords = Trim(Left(ReplyRec.Data, 1))
        ReplyRec.Data = Right(ReplyRec.Data, Len(ReplyRec.Data) - 1)
        gblIncrementCount = Trim(Left(ReplyRec.Data, 4))
        ReplyRec.Data = Right(ReplyRec.Data, Len(ReplyRec.Data) - 4)
        
        If gblIncrementCount = 0 Then
            StatusMsg = "The Increment Count was 0 - there is no data on the Corporate Database!"
            Print #1, StatusMsg
            End
        End If
        
        If IncPageCount = 1 Then
            ReDim gblIncrementTable(1)
            DoEvents
            ReDim gblIncrementTable(gblIncrementCount + 1)
        ElseIf IncPageCount > 1 Then
            If b > 300 Then
                gblIncrementCount = gblIncrementCount + b
            End If
            ReDim Preserve gblIncrementTable(gblIncrementCount + 1)
        End If
        
        For a = b To gblIncrementCount
            gblIncCount = a
            gblIncrementTable(gblIncCount) = Left(ReplyRec.Data, 75) '& Space(12) & Space(50)
            ReplyRec.Data = Right(ReplyRec.Data, Len(ReplyRec.Data) - 75)
        Next a
    End If
End Sub
Public Sub Get_Commands()
'VETNT00018229 DSG Get the commands involved in Consistency checking and load them into a table
'This function will also return the Database Benefit Types eg. "CPL", "CPD" and load them into a table
'It will also return the Payee Groupings used in Consistency and load them into a table
   
    TP.Client_Service_Name = "shrinqepc"
    TP.Client_Appl_Data_Send_Len = 6
    TP.Client_Appl_Data_Recv_Len = 26000
    SendRec.Data = "GETCMD"
'    DisplayReceive = True
    SHARE_CallTuxedo
'    DisplayReceive = False
    
    If gReturnCode <> "FATAL" Then
        Load_Commands
    End If
End Sub
Public Sub Load_Commands()
'VETNT00018229 DSG Parse the Commands, Benefits, and Payees and load them into tables for Table Driven EPC's
    If Left(ReplyRec.Data, 9) = "SHAR00000" Then
        ReplyRec.Data = Right(ReplyRec.Data, Len(ReplyRec.Data) - 9)
        gblCommandsCount = Trim(Left(ReplyRec.Data, 2))
        ReplyRec.Data = Right(ReplyRec.Data, Len(ReplyRec.Data) - 2)
        If gblCommandsCount = 0 Then
            StatusMsg = "The Command Count was 0 - there is no data on the Corporate Database!"
            Print #1, StatusMsg
            End
        End If
        ReDim gblHoldCommands(gblCommandsCount + 1)
        'Load The Commands
        For a = 1 To gblCommandsCount
            gblHoldCommands(a) = Left(ReplyRec.Data, 42)
            ReplyRec.Data = Right(ReplyRec.Data, Len(ReplyRec.Data) - 42)
        Next a
      
        gblBenefitsCount = Trim(Left(ReplyRec.Data, 2))
        ReplyRec.Data = Right(ReplyRec.Data, Len(ReplyRec.Data) - 2)
        If gblBenefitsCount = 0 Then
            StatusMsg = "The Benefit Code Count was 0 - there is no data on the Corporate Database!"
            Print #1, StatusMsg
            End
        End If
        ReDim gblHoldBenefitCodes(gblBenefitsCount + 1)
        'Load the Benefits
        For a = 1 To gblBenefitsCount
            gblHoldBenefitCodes(a) = Left(ReplyRec.Data, 62)
            ReplyRec.Data = Right(ReplyRec.Data, Len(ReplyRec.Data) - 62)
        Next a
        
        'Load Payees
        gblPayeeCount = Trim(Left(ReplyRec.Data, 2))
        If gblPayeeCount = 0 Then
            StatusMsg = "The Payee Count was 0 - there is no data on the Corporate Database!"
            Print #1, StatusMsg
            End
        End If
        ReplyRec.Data = Right(ReplyRec.Data, Len(ReplyRec.Data) - 2)
        ReDim gblHoldPayeeRanges(gblPayeeCount + 1)
        For a = 1 To gblPayeeCount
            gblHoldPayeeRanges(a) = Left(ReplyRec.Data, 50)
            ReplyRec.Data = Right(ReplyRec.Data, Len(ReplyRec.Data) - 50)
        Next a
    End If

End Sub
Public Sub Get_All_EPCs()
'VETNT00018229 DSG Get all the available End Product Codes and their descriptions
'and load them into a table
    EPCCount = 0
    gblEPCcount = 0
    EPCPageCount = 0
    MoreRecords = "Y"
    Do Until MoreRecords = "N"
        EPCPageCount = EPCPageCount + 1
        TP.Client_Service_Name = "shrinqepc"
        TP.Client_Appl_Data_Send_Len = 9
        TP.Client_Appl_Data_Recv_Len = 26000
        SendRec.Data = "GETEPC" & Format(EPCPageCount, "00")
        SHARE_CallTuxedo
        If gReturnCode <> "FATAL" Then
            Load_All_EPCs_Cmpr
        Else
            Exit Do
        End If
    Loop
    If EPCCount > gblEPCcount Then
        ReDim Preserve gblHoldProcessEPCs(gblEPCcount + 1)
    End If
    If gblEPCcount = 0 Then
        StatusMsg = "The EPC Count was 0 - there is no data on the Corporate Database!"
        Print #1, StatusMsg
        End
    End If
End Sub
Public Sub Load_All_EPCs_Cmpr()
    Dim NotNumeric As Integer
    NotNumeric = 0
'VETNT00018229 DSG Parse and Load EPC's into a table
    If Left(ReplyRec.Data, 9) = "SECU00000" Then
        ReplyRec.Data = Right(ReplyRec.Data, Len(ReplyRec.Data) - 9)
        MoreRecords = Trim(Left(ReplyRec.Data, 1))
        ReplyRec.Data = Right(ReplyRec.Data, Len(ReplyRec.Data) - 1)
        If EPCPageCount = 1 Then
            EPCCount = Trim(Left(ReplyRec.Data, 4))
            ReDim gblHoldProcessEPCs(EPCCount + 1)
        End If
        ReplyRec.Data = Right(ReplyRec.Data, Len(ReplyRec.Data) - 4)
        For a = 1 To 401
            If IsNumeric(Left(ReplyRec.Data, 3)) Then
                gblEPCcount = gblEPCcount + 1
                gblHoldProcessEPCs(gblEPCcount) = Trim(Left(ReplyRec.Data, 62))
                ReplyRec.Data = Right(ReplyRec.Data, Len(ReplyRec.Data) - 62)
            Else
                ReplyRec.Data = Right(ReplyRec.Data, Len(ReplyRec.Data) - 62)
                NotNumeric = NotNumeric + 1
            End If
        Next a
    End If
End Sub
Public Sub Get_Consistency_Table()
'VETNT00018229 DSG Get the Consistency table values to load them into a table
    EPCPageCount = 0
    gblConsistencyCount = 0
    Do Until MoreRecords = "N"
        EPCPageCount = EPCPageCount + 1
        TP.Client_Service_Name = "shrinqepc"
         TP.Client_Appl_Data_Send_Len = 9
        TP.Client_Appl_Data_Recv_Len = 26000
        SendRec.Data = "GETMNG" & Format(EPCPageCount, "000")
        SHARE_CallTuxedo
        If gReturnCode <> "FATAL" Then
            Load_Consistency_Table
        Else
            Exit Do
        End If
    Loop
    If gblConsistencyCount = 0 Then
        StatusMsg = "The Consistency Count was 0 - there is no data on the Corporate Database!"
        Print #1, StatusMsg
        End
    End If
   
End Sub
Public Sub Load_Consistency_Table()
'VETNT00018229 DSG Parse and Load Consistency values to table
    Dim ConsistencyCount As Integer
    If Left(ReplyRec.Data, 9) = "SECU00000" Then
        ReplyRec.Data = Right(ReplyRec.Data, Len(ReplyRec.Data) - 9)
        MoreRecords = Trim(Left(ReplyRec.Data, 1))
        ReplyRec.Data = Right(ReplyRec.Data, Len(ReplyRec.Data) - 1)
        ConsistencyCount = Trim(Left(ReplyRec.Data, 4))
        ReplyRec.Data = Right(ReplyRec.Data, Len(ReplyRec.Data) - 4)
        If EPCPageCount = 1 Then
            ReDim gblHoldConsistency(ConsistencyCount + 1)
        Else
            ReDim Preserve gblHoldConsistency(ConsistencyCount + gblConsistencyCount)
        End If
        For a = 1 To ConsistencyCount
            If Trim(Left(ReplyRec.Data, 25)) = "" Then
                Exit For
            Else
                gblConsistencyCount = gblConsistencyCount + 1
                gblHoldConsistency(gblConsistencyCount) = Left(ReplyRec.Data, 25)
                ReplyRec.Data = Right(ReplyRec.Data, Len(ReplyRec.Data) - 25)
            End If
        Next a
    End If
End Sub

Public Function Clear_And_Reload_Consistency_From_MDB()
    On Error GoTo ErrorDB
    Set dbEPC = OpenDatabase(DatabasePath, True, False, ";pwd=consistdsg")
    fCriteria = "Select * FROM Refresh_Date_Table"
    Set Refresh_Date_Table = dbEPC.OpenRecordset(fCriteria)
    If Not Refresh_Date_Table.EOF Then
        Refresh_Date_Table.Delete
    End If
    Refresh_Date_Table.Close
    fCriteria = "Select * FROM Command_Table"
    Set Command_Table = dbEPC.OpenRecordset(fCriteria)
    If Not Command_Table.EOF Then
        With Command_Table
            .MoveFirst
            Do While Not .EOF
                .Delete
                .MoveNext
            Loop
        End With
    End If
    Command_Table.Close
    fCriteria = "Select * FROM Benefit_Type_Table"
    Set Benefit_Type_Table = dbEPC.OpenRecordset(fCriteria)
    If Not Benefit_Type_Table.EOF Then
        With Benefit_Type_Table
            .MoveFirst
            Do While Not .EOF
                .Delete
                .MoveNext
            Loop
        End With
    End If
    Benefit_Type_Table.Close
    fCriteria = "Select * FROM Payee_Range_Table"
    Set Payee_Range_Table = dbEPC.OpenRecordset(fCriteria)
    If Not Payee_Range_Table.EOF Then
        With Payee_Range_Table
            .MoveFirst
            Do While Not .EOF
                .Delete
                .MoveNext
            Loop
        End With
    End If
    Payee_Range_Table.Close
    fCriteria = "Select * FROM Consistency_Table"
    Set Consistency_Table = dbEPC.OpenRecordset(fCriteria)
    If Not Consistency_Table.EOF Then
        With Consistency_Table
            .MoveFirst
            Do While Not .EOF
                .Delete
                .MoveNext
            Loop
        End With
    End If
    Consistency_Table.Close
    fCriteria = "Select * FROM Increment_Table"
    Set Increment_Table = dbEPC.OpenRecordset(fCriteria)
    If Not Increment_Table.EOF Then
        With Increment_Table
            .MoveFirst
            Do While Not .EOF
                .Delete
                .MoveNext
            Loop
        End With
    End If
    Increment_Table.Close
    fCriteria = "Select * FROM EPC_Table"
    Set EPC_Table = dbEPC.OpenRecordset(fCriteria)
    If Not EPC_Table.EOF Then
        With EPC_Table
            .MoveFirst
            Do While Not .EOF
                .Delete
                .MoveNext
            Loop
        End With
    End If
    EPC_Table.Close
    Get_Data_For_Consistency
    result = Save_Consistency_To_MDB
    Exit Function
ErrorDB:
    Dim Msg As String
    If Err <> 0 Then
        gblDBContensionError = True
        dbEPC.Close
        End
    End If
End Function

Public Function Save_Consistency_To_MDB()
    On Error GoTo ErrorDB
    Dim EnvironmentDB As String
    Dim HoldEnvID As Long
    Dim WorkArea As String
    Dim WorkArea2 As String

'    passform.ctlProgress.Min = 0
'    passform.ctlProgress.Max = 100
    
'    Set dbEPC = OpenDatabase(DatabasePath, True, False, ";pwd=consistdsg")
'    passform.ctlProgress.Value = 15
   
    fCriteria = "Select * FROM Command_Table"
    Set Command_Table = dbEPC.OpenRecordset(fCriteria)
    If Command_Table.EOF Then
        For a = 1 To gblCommandsCount
            Command_Table.AddNew
            With Command_Table
                !Command = Left(gblHoldCommands(a), 12)
                !Command_Descr = Right(gblHoldCommands(a), 30)
                Command_Table.Update
            End With
        Next a
    End If
    Command_Table.Close
'    passform.ctlProgress.Value = passform.ctlProgress.Value + 15
    fCriteria = "Select * FROM Benefit_Type_Table"
    Set Benefit_Type_Table = dbEPC.OpenRecordset(fCriteria)
    If Benefit_Type_Table.EOF Then
        For a = 1 To gblBenefitsCount
            Benefit_Type_Table.AddNew
            With Benefit_Type_Table
                !Benefit_Type = Left(gblHoldBenefitCodes(a), 12)
                !Benefit_Descr = Right(gblHoldBenefitCodes(a), 50)
                Benefit_Type_Table.Update
            End With
        Next a
    End If
    Benefit_Type_Table.Close
'    passform.ctlProgress.Value = passform.ctlProgress.Value + 15
    fCriteria = "Select * FROM Payee_Range_Table"
    Set Payee_Range_Table = dbEPC.OpenRecordset(fCriteria)
    If Payee_Range_Table.EOF Then
        For a = 1 To gblPayeeCount
            Payee_Range_Table.AddNew
            With Payee_Range_Table
               !Payee_Range = Left(gblHoldPayeeRanges(a), 12)
                Payee_Range_Table.Update
            End With
        Next a
    End If
    Payee_Range_Table.Close
'    passform.ctlProgress.Value = passform.ctlProgress.Value + 15
    fCriteria = "Select * FROM Consistency_Table"
    Set Consistency_Table = dbEPC.OpenRecordset(fCriteria)
    If Consistency_Table.EOF Then
        For a = 1 To gblConsistencyCount
            Consistency_Table.AddNew
            With Consistency_Table
                !Command_Ref = Left(gblHoldConsistency(a), 2)
                !Benefit_Ref = Mid(gblHoldConsistency(a), 3, 2)
                !Payee_Ref = Mid(gblHoldConsistency(a), 5, 2)
                !EPC_Ref = Mid(gblHoldConsistency(a), 7, 3)
                !Effective_Date = Mid(gblHoldConsistency(a), 10, 8)
                !Disable_Date = Mid(gblHoldConsistency(a), 18, 8)
                Consistency_Table.Update
            End With
        Next a
    End If
    Consistency_Table.Close
'    passform.ctlProgress.Value = passform.ctlProgress.Value + 15
    fCriteria = "Select * FROM Increment_Table"
    Set Increment_Table = dbEPC.OpenRecordset(fCriteria)
    If Increment_Table.EOF Then
        For a = 1 To gblIncrementCount
            Increment_Table.AddNew
            With Increment_Table
                !Parent_EPC = Left(gblIncrementTable(a), 4)
                !Actual_EPC = Mid(gblIncrementTable(a), 5, 4)
                !Increment_To_EPC = Mid(gblIncrementTable(a), 9, 4)
                !Type_Single = Mid(gblIncrementTable(a), 13, 1)
                !Benefit_Type = Mid(gblIncrementTable(a), 14, 12)
                !Form_Name = Mid(gblIncrementTable(a), 26, 30)
                Increment_Table.Update
            End With
        Next a
    End If
    Increment_Table.Close
'    passform.ctlProgress.Value = passform.ctlProgress.Value + 15
    fCriteria = "Select * FROM EPC_Table"
    Set EPC_Table = dbEPC.OpenRecordset(fCriteria)
    If EPC_Table.EOF Then
        For a = 1 To gblEPCcount
            EPC_Table.AddNew
            With EPC_Table
                !EPC_Code = Left(gblHoldProcessEPCs(a), 12)
                !EPC_Descr = Right(gblHoldProcessEPCs(a), 50)
                EPC_Table.Update
            End With
        Next a
    End If
    EPC_Table.Close
    fCriteria = "Select * FROM Refresh_Date_Table"
    Set Refresh_Date_Table = dbEPC.OpenRecordset(fCriteria)
    If Refresh_Date_Table.EOF Then
        Refresh_Date_Table.AddNew
        With Refresh_Date_Table
            !RefreshDate = Format(Now, "mm/dd/yyyy hh:mm:ss AMPM")
            !Source = Trim(cs.DatabaseName)
            Refresh_Date_Table.Update
        End With
    End If
    Refresh_Date_Table.Close
    dbEPC.Close
    CompactDatabase (DatabasePath), DatabasePath & ".tmp", , , ";pwd=consistdsg"
    Kill DatabasePath
    Name DatabasePath & ".tmp" As (DatabasePath)
    StatusMsg = "Success - the MDB was cleared, reloaded, and compacted."
    Print #1, StatusMsg
    Exit Function
ErrorDB:
    If Err <> 0 Then
        End
    End If
End Function
Public Function SHARE_CallTuxedo() As Integer

On Error GoTo Err_SHARE_CallTuxedo

    'Set the Mouse to Hourglass during the call to let us know
    'we're busy during the call
    Dim SavePointer As Integer
    SavePointer = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    gReturnCode = "OK"

    'Tell the Tuxedo program to do the Call function.  In this case
    'send application name and new password if any

    ReplyRec.Data = ""
    '08/16/2006 mlc cmted out fro removal of string disp
    'Display_Transaction_Send

    TP.Client_Tuxedo_Function = TPClient_Do_TPCall
    'call tuxedo wrapper
    Call TPCLIENT(TP, SendRec, ReplyRec)

    '08/16/2006 mlc cmted out with removal of string display
    'Display_Transaction_Receive
    'check tuxedo status
        If TP.Client_TP_Status_Code = TPClient_TP_Status_Sys_Error Then
        gReturnCode = "FATAL"
        If Left(ReplyRec.Data, 9) = "TUX-10216" Or Left(ReplyRec.Data, 9) = "TUX-10215" Then
            StatusMsg = "SHR CMN02: The Tuxedo Server is down. Please notify your Administrator."
            Print #1, StatusMsg
            cs.Function_Type = "RESTART"
            cs.CSSTUX_Function
'            OutOfHere
        ElseIf Mid(ReplyRec.Data, 1, 9) = "TUX-20302" Then
                StatusMsg = "SHR CMN03: Communication Lines are Down." & vbCr & _
                   "Try Again Later or Contact Help Desk."
                Print #1, StatusMsg
                OutOfHere
        ElseIf Mid(ReplyRec.Data, 1, 9) = "TUX-20307" Then
                StatusMsg = "SHR CMN04: Line Timeout." & vbCr & _
                    "Possible cause may be that a Service is down." & vbCr & _
                   "Try Again Later or Contact Help Desk."
                Print #1, StatusMsg
        ElseIf Mid(ReplyRec.Data, 1, 9) = "TUX-20306" Then
                StatusMsg = "SHR CMN05: No entry Point for Service." & vbCr & _
                    "Possible cause may be that a Service is down or" & vbCr & _
                    "the Service Name is misspelled." & vbCr & _
                    "Try Again Later or Contact Help Desk."
                Print #1, StatusMsg
        Else
                StatusMsg = "SHR CMN06: Tuxedo Error." & vbCr & _
                   Left(ReplyRec.Data, Len(ReplyRec.Data))
                Print #1, StatusMsg
        End If
        GoTo Exit_SHARE_CallTuxedo
    Else
'03/09/2004 - mlc VBA0000DSG - added if stmt for application error and added "ORA-" to second if stmt
        If TP.Client_TP_Status_Code = TPClient_TP_Status_App_Error Then
            gReturnCode = "FATAL"
            If Trim(ReplyRec.Data) = "GUIE50007" Then 'VETNT00019834 & 20221 VR2 09/19/2007
            Else                                      'VETNT00019834 & 20221 VR2 09/19/2007
            If Trim(ReplyRec.Data) = "GUIE50008" Then 'VETNT00019834 & 20221 VR2 09/19/2007
            Else                                      'VETNT00019834 & 20221 VR2 09/19/2007
                StatusMsg = "SHR CMN08: Application Error." & vbCrLf & Trim(ReplyRec.Data) & vbCrLf & "Service is " & TP.Client_Service_Name & vbCrLf & "Please notify your Administrator."
                Print #1, StatusMsg
            End If                                    'VETNT00019834 & 20221 VR2 09/19/2007
            End If                                    'VETNT00019834 & 20221 VR2 09/19/2007
        Else
            If Left(ReplyRec.Data, 4) = "ORAC" Or Left(ReplyRec.Data, 4) = "ORA-" Then
                gReturnCode = "FATAL"
                StatusMsg = "SHR CMN07: The Oracle Error." & vbCrLf & Trim(ReplyRec.Data) & vbCrLf & "Please notify your Administrator."
                Print #1, StatusMsg
                OutOfHere
            End If
        End If
    End If
    'clear hourglass
    Screen.MousePointer = SavePointer




Exit_SHARE_CallTuxedo:
    'SHARE_CallTuxedo = CSS_InitStatus
    Screen.MousePointer = SavePointer
    Exit Function

Err_SHARE_CallTuxedo:
    StatusMsg = Error$
    Print #1, StatusMsg
    Resume Exit_SHARE_CallTuxedo

End Function
Public Sub OutOfHere()
'02/16/2005 mlc - VBA00002314 added 'End'


    cs.Function_Type = "TERMINATE"
    cs.CSSTUX_Function

    End

End Sub
Public Function CSS_InitializeTuxedo() As Integer

On Error GoTo Err_CSS_InitializeTuxedo


'    Before_Init
    'Set default to good init
    TP.Client_Appl_Data_Send_Len = 62
    TP.Client_Appl_Data_Recv_Len = 5000

    'Set the Mouse to vbhourglass during the call to let us know
    'we're busy during the call
'    Screen.MousePointer = vbHourglass
    'Initialize reply data
    ReplyRec.Data = ""
    'Tell the Tuxedo program to do the Initialize function.  In this case
    'call the service and start a session
    'call tuxedo wrapper


    Call TPCLIENT(TP, SendRec, ReplyRec)

    'check tuxedo status
    If TP.Client_Tuxedo_Function = 98 Then
        Exit Function
    End If
Exit_CSS_InitializeTuxedo:
    CSS_InitializeTuxedo = CSS_InitStatus
    Exit Function

Err_CSS_InitializeTuxedo:
    StatusMsg = Error$
    Print #1, StatusMsg
    CSSTUX_ReturnCode = "FATAL"
    CSS_InitStatus = False
    Resume Exit_CSS_InitializeTuxedo

End Function

