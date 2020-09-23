VERSION 5.00
Begin VB.Form frmCreateDatabase 
   Caption         =   "Create Database"
   ClientHeight    =   1170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   4860
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton cmdCreateDatabase 
      Caption         =   "Create Database"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox txtNewDatabaseName 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "New database name:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmCreateDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This program was made by José M. Nieves (nievesj@prtc.net)
'You can use it in any way you see it usefull, just give me some credit.
'I made this because I were bored so dont expect the code to be elegant or professional :), it just works.

Private Sub cmdCancel_Click()
txtNewDatabaseName = ""
Me.Hide
End Sub

Private Sub cmdCreateDatabase_Click()
Dim strNewDBName As String
Dim intLindex As Integer
Dim RepeatedDBName As Boolean

For intLindex = 0 To frmConnect.lstDatabases.ListCount
    If txtNewDatabaseName = frmConnect.lstDatabases.List(intLindex) Then
        'there are two columns with the same name
        RepeatedDBName = True
        intLindex = frmConnect.lstDatabases.ListCount 'To exit the For loop
    End If
    DoEvents
Next intLindex
If RepeatedDBName = True Then
    Call MsgBox(txtNewDatabaseName & " is already a database.", vbInformation + vbOKOnly, Me.Caption)
Else
    DBCon.Execute "CREATE DATABASE " & txtNewDatabaseName
    frmConnect.PopulateDBList
End If
End Sub

Private Sub Form_Deactivate()
frmConnect.PopulateDBList
End Sub

