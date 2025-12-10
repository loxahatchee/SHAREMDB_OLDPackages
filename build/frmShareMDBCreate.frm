VERSION 5.00
Begin VB.Form frmShareCreate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create Share MDB"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   750
      Left            =   1260
      TabIndex        =   1
      Top             =   1995
      Width           =   1770
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Share MDB"
      Height          =   750
      Left            =   1290
      TabIndex        =   0
      Top             =   705
      Width           =   1770
   End
End
Attribute VB_Name = "frmShareCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClear_Click()
    Public Function Clear_And_Reload_Consistency_From_MDB(passform As Form)
    Set dbEPC = OpenDatabase(DatabasePath, True, False, ";pwd=consistdsg")
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
    fCriteria = "Select * FROM Refresh_Date_Table"
    Set Refresh_Date_Table = dbEPC.OpenRecordset(fCriteria)
    If Not Refresh_Date_Table.EOF Then
        Refresh_Date_Table.Delete
    End If
    Refresh_Date_Table.Close
    epcdb.Close
End Function

Private Sub cmdExit_Click()
    Unload Me
End Sub
