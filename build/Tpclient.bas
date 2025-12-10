Attribute VB_Name = "tpclnt32"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Tpclient.bas on Mon 12/29/03 @ 12:01 
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Tpclient.bas on Tue 11/4/03 @ 12:58 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Public Constants (Removed)                                                             *
'*  TPClient_Do_TPInitialize      TPClient_Do_TPTerm            TPClient_Do_TPUserlog     *
'*  TPClient_Do_Conv_Send         TPClient_Do_Conv_Recv         TPClient_Do_Conv_Abort    *
'*  TPClient_Do_TPAcall_With_Reply                              TPClient_Do_TPGetreply    *
'*  TPClient_Do_TPAcall_No_Reply  TPClient_Do_TPBegin           TPClient_Do_TPCommit      *
'*  TPClient_Do_TPAbort           TPClient_Do_TPBroadcast       TPClient_Do_TPChkunsol    *
'*  TPClient_Do_TPGetreply_No_Block                             TPClient_Do_MC_TPInit     *
'*  TPClient_Do_MC_TPInit_For_Web TPClient_Do_MC_TPInit_Diff_UID                          *
'*  TPClient_Do_TPInit_For_Web    TPClient_Do_TPInit_Diff_UID   TPClient_Get_Userid       *
'*  TPClient_Get_VBA_Version      TPClient_TP_Env_Devl          TPClient_TP_Env_Integrated*
'*  TPClient_TP_Env_Cert          TPClient_TP_Env_Performance   TPClient_TP_Env_Academy   *
'*  TPClient_TP_Env_Prod          TPClient_TP_Env_TP_SoftTest   TPClient_TP_Env_Unknown   *
'*  TPClient_Client_Has_Send_Cntrl                              TPCLient_Server_Has_Send_Cntrl
'*  TPClient_Server_Ended_Conv                                                            *
'******************************************************************************************


Option Explicit                                    'Version 8.8  06/2003

Type Client_Admin_Info_Record

     Client_Tuxedo_Function         As Long         'Function TPClient is to do
     Client_TP_Status_Code          As Long         '0=OK 1=SysError 2=AppError
     Client_Appl_Data_Send_Len      As Long         'Length of outgoing message
     Client_Appl_Data_Recv_Len      As Long         'Length of incoming message
     Client_Client_Module_Name      As String * 15  'Calling Form name or module name
     Client_Service_Name            As String * 15  'Service being requested on TPCALL
     Client_Route_Station           As String * 3   'Station # to send message to
     Client_Client_User_Name        As String * 15  'User's Lan login id
     Client_Client_Computer_Name    As String * 15  'User's Lan computer name
     Client_TP_Environment_Ind      As String * 1   'Indicates TP environment
     Client_Conv_Control_Ind        As String * 1   'Indicates status of conversation
     Client_Multi_Context_Addr      As Long         'Multi Context Connection Address PTR 
End Type

'define the admin record that will be passed to TPCLIENT
Global TP As Client_Admin_Info_Record

'define constants that will be used to set and check variables
Global Const TPClient_Do_TPCall = 2

'Multi Context tpinit functions:

'Alternate single context tpinit functions:


Global Const TPClient_TP_Status_OK = 0
Global Const TPClient_TP_Status_Sys_Error = 1
Global Const TPClient_TP_Status_App_Error = 2




'declare the TPCLIENT subroutine call within the tpclnt32.dll
Declare Sub TPCLIENT Lib "tpclnt32.dll" (Admin As Client_Admin_Info_Record, TPClient_Appl_Data_Send_Area As Any, TPClient_Appl_Data_Recv_Area As Any)
