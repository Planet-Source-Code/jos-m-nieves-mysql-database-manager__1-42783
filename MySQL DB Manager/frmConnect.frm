VERSION 5.00
Begin VB.Form frmConnect 
   Caption         =   "Connect To The MySQL Database Server"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   9975
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   6600
      TabIndex        =   23
      Top             =   5160
      Width           =   3255
   End
   Begin VB.Frame frmWorkDB 
      Caption         =   "Working With: "
      Enabled         =   0   'False
      Height          =   3855
      Left            =   0
      TabIndex        =   16
      Top             =   1920
      Width           =   6375
      Begin VB.CommandButton cmdDataEntry 
         Caption         =   "Data Entry"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2760
         TabIndex        =   22
         Top             =   3480
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton cmdDropTable 
         Caption         =   "Drop Table"
         Height          =   255
         Left            =   4560
         TabIndex        =   21
         Top             =   3480
         Width           =   1695
      End
      Begin VB.CommandButton cmdEditTable 
         Caption         =   "Edit Table"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2760
         TabIndex        =   20
         Top             =   2760
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton cmdCreateTable 
         Caption         =   "Create Table"
         Height          =   255
         Left            =   4560
         TabIndex        =   19
         Top             =   3120
         Width           =   1695
      End
      Begin VB.ListBox lstTables 
         Appearance      =   0  'Flat
         Height          =   3150
         ItemData        =   "frmConnect.frx":0000
         Left            =   120
         List            =   "frmConnect.frx":0002
         TabIndex        =   18
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Tables"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdCreateDB 
      Caption         =   "Create Database"
      Height          =   255
      Left            =   6720
      TabIndex        =   15
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton cmdDropDB 
      Caption         =   "Drop Database"
      Height          =   255
      Left            =   6720
      TabIndex        =   14
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton cmdSelectDB 
      Caption         =   "Select Database"
      Height          =   255
      Left            =   6720
      TabIndex        =   13
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Frame frmDatabases 
      Caption         =   "Databases"
      Enabled         =   0   'False
      Height          =   3855
      Left            =   6600
      TabIndex        =   11
      Top             =   120
      Width           =   3255
      Begin VB.ListBox lstDatabases 
         Appearance      =   0  'Flat
         Height          =   2370
         ItemData        =   "frmConnect.frx":0004
         Left            =   120
         List            =   "frmConnect.frx":0006
         TabIndex        =   12
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.CommandButton cmdDisconect 
      Caption         =   "Disconnect"
      Height          =   255
      Left            =   5040
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtPass 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtUser 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtHost 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Database Server Connection"
      Height          =   1695
      Left            =   0
      TabIndex        =   5
      Top             =   120
      Width           =   6375
      Begin VB.Label lblStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3120
         TabIndex        =   10
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label Label5 
         Caption         =   "Connection Status:"
         Height          =   255
         Left            =   3120
         TabIndex        =   9
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Password"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "User Name"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Server/Host/IP"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This program was made by José M. Nieves (nievesj@prtc.net)
'You can use it in any way you see it usefull, just give me some credit.
'I made this because I were bored so dont expect the code to be elegant or professional :), it just works.

Option Explicit
Dim connected As Boolean

Private Sub cmdConnect_Click()
strHost = txtHost
strDatabase = "" 'txtDatabase
strUser = txtUser
strPass = txtPass
On Error GoTo conERROR
AbrirConeccion
lblStatus.Caption = "Connected!!!"
connected = True
frmDatabases.Enabled = True
PopulateDBList
Exit Sub
conERROR:
    Call MsgBox("ERROR: " & DBCon.Errors(0), vbCritical + vbOKOnly, Me.Caption)
    Set DBCon = Nothing
End Sub

Private Sub cmdCreateDB_Click()
frmCreateDatabase.Show vbModal, Me
End Sub

Private Sub cmdCreateTable_Click()
frmCreateTable.Show vbModal
End Sub

Private Sub cmdDisconect_Click()
CerrarConeccion
lblStatus.Caption = "Disconnected"
connected = False
frmDatabases.Enabled = False
lstDatabases.Clear
lstTables.Clear
End Sub

Private Sub cmdDropDB_Click()
Dim strDatabaseName As String
Dim responce As VbMsgBoxResult

strDatabaseName = lstDatabases.List(lstDatabases.ListIndex)

Select Case strDatabaseName
    Case "mysql"
        Call MsgBox("Cannot delete database " & strDatabaseName, vbInformation + vbOKOnly, Me.Caption)
    Case ""
        Call MsgBox("Please select a database to Drop.", vbInformation + vbOKOnly, Me.Caption)
    Case Else
        responce = MsgBox("Are you sure you want to drop the database " & strDatabaseName & "?", vbQuestion + vbYesNo, Me.Caption)
        If responce = vbYes Then
            DBCon.Execute ("DROP DATABASE " & strDatabaseName)
            PopulateDBList
        End If
End Select
End Sub

Private Sub cmdDropTable_Click()
Dim responce As VbMsgBoxResult

On Error GoTo Nada
If lstTables.ListCount > 0 Then
    responce = MsgBox("Are you sure you want to drop" & _
            lstTables.List(lstTables.ListIndex), vbQuestion + vbYesNo, Me.Caption)
    If responce = vbYes Then
        DBCon.Execute "DROP TABLE " & lstTables.List(lstTables.ListIndex)
        PopulatelstTables
    End If
End If
Exit Sub
Nada:
Exit Sub
End Sub

Private Sub cmdExit_Click()
SalirDelPrograma
End Sub

Private Sub cmdSelectDB_Click()
strDatabase = lstDatabases.List(lstDatabases.ListIndex)
PopulatelstTables
End Sub

Public Sub PopulatelstTables()
Dim tmpstrDBName As String

tmpstrDBName = lstDatabases.List(lstDatabases.ListIndex)
If tmpstrDBName = "" Then
    Call MsgBox("Select a database to work with.", vbInformation + vbOKOnly, Me.Caption)
Else
    DBCon.Execute ("USE " & tmpstrDBName)
    Set RecSet = DBCon.Execute("SHOW Tables")
    lstTables.Clear
    Do While Not RecSet.EOF
        lstTables.AddItem RecSet.Fields(0).Value
        RecSet.MoveNext
        DoEvents
        Loop
    frmWorkDB.Enabled = True
End If
End Sub

Public Sub PopulateDBList()
Set RecSet = DBCon.Execute("SHOW Databases")
lstDatabases.Clear
Do While Not RecSet.EOF
  lstDatabases.AddItem RecSet.Fields(0).Value
  RecSet.MoveNext
  DoEvents
Loop
End Sub

Private Sub Form_Load()
connected = False
End Sub
