VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCreateTable 
   Caption         =   "CREATE TABLE"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   10185
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close This Window"
      Height          =   375
      Left            =   5040
      TabIndex        =   26
      Top             =   7920
      Width           =   2295
   End
   Begin VB.CommandButton cmdApplyChanges 
      Caption         =   "Apply Changes To Database"
      Height          =   375
      Left            =   7800
      TabIndex        =   24
      Top             =   7920
      Width           =   2295
   End
   Begin VB.Frame Frame4 
      Caption         =   "Data Type Explanations"
      Height          =   3615
      Left            =   5040
      TabIndex        =   23
      Top             =   4200
      Width           =   5055
      Begin RichTextLib.RichTextBox rtbDTExplanations 
         Height          =   3255
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   5741
         _Version        =   393217
         BackColor       =   -2147483648
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmCreateTable.frx":0000
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Set Primary Key"
      Height          =   1215
      Left            =   120
      TabIndex        =   20
      Top             =   7080
      Width           =   4815
      Begin VB.ComboBox cboSelectPrimaryKey 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label6 
         Caption         =   "Select Primary Key"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.TextBox txtNewTableName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   12
      Top             =   120
      Width           =   2895
   End
   Begin VB.Frame Frame2 
      Caption         =   "Add Column"
      Height          =   2775
      Left            =   120
      TabIndex        =   3
      Top             =   4200
      Width           =   4815
      Begin VB.CommandButton cmdAddCol 
         Caption         =   "Add Column"
         Height          =   255
         Left            =   3240
         TabIndex        =   19
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox txtDefault 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         TabIndex        =   18
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CheckBox chkAutoIncrement 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Auto Increment"
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   3600
         TabIndex        =   16
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CheckBox chkBinary 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Binary"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3600
         TabIndex        =   15
         Top             =   1440
         Width           =   855
      End
      Begin VB.CheckBox chkZeroFill 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Zero Fill"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3600
         TabIndex        =   14
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CheckBox chkUnsigned 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Unsigned"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3600
         TabIndex        =   13
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox chkNotNull 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Not Null"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3600
         TabIndex        =   10
         Top             =   360
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.TextBox txtLength 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         TabIndex        =   9
         Top             =   1320
         Width           =   1335
      End
      Begin VB.ComboBox cboDataType 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmCreateTable.frx":00D5
         Left            =   1320
         List            =   "frmCreateTable.frx":011E
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox txtColumnName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Default"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Lenght or Enum/ Set Elements"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Data Type"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Column Name"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Column List"
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   9975
      Begin VB.CommandButton cmdRemoveColumn 
         Caption         =   "Remove Column"
         Height          =   375
         Left            =   8400
         TabIndex        =   2
         Top             =   3000
         Width           =   1455
      End
      Begin VB.ListBox lstColumnList 
         Appearance      =   0  'Flat
         Height          =   2565
         ItemData        =   "frmCreateTable.frx":01DC
         Left            =   120
         List            =   "frmCreateTable.frx":01DE
         TabIndex        =   1
         Top             =   240
         Width           =   9735
      End
   End
   Begin VB.Label lblWorkingWithDB 
      Height          =   255
      Left            =   6840
      TabIndex        =   28
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label7 
      Caption         =   "Working With Database:"
      Height          =   255
      Left            =   4800
      TabIndex        =   27
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Table Name"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmCreateTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This program was made by José M. Nieves (nievesj@prtc.net)
'You can use it in any way you see it usefull, just give me some credit.
'I made this because I were bored so dont expect the code to be elegant or professional :), it just works.

Option Explicit
Dim strNotNull As String
Dim strUnsigned As String
Dim strZeroFill As String
Dim strBinary As String
Dim strAuto_Increment As String

Private Sub cboDataType_Click()
Dim intDataTypeIndex As Integer
Dim strDataTypeString As String

intDataTypeIndex = cboDataType.ListIndex
strDataTypeString = cboDataType.List(intDataTypeIndex)
rtbDTExplanations.FileName = App.Path & "\" & strDataTypeString & ".rtf"
End Sub

Private Sub cboSelectPrimaryKey_Click()
'Label7.Caption = cboSelectPrimaryKey.ListIndex
End Sub

Private Sub chkAutoIncrement_Click()
If chkAutoIncrement.Value = vbChecked Then
    strAuto_Increment = " AUTO_INCREMENT"
Else
    strAuto_Increment = ""
End If
End Sub

Private Sub chkBinary_Click()
If chkBinary.Value = vbChecked Then
    strBinary = " BINARY"
Else
    strBinary = ""
End If
End Sub

Private Sub chkNotNull_Click()
If chkNotNull.Value = vbChecked Then
    strNotNull = " NOT NULL"
Else
    strNotNull = ""
End If
End Sub

Private Sub chkUnsigned_Click()
If chkUnsigned.Value = vbChecked Then
    strUnsigned = " UNSIGNED"
Else
    strUnsigned = ""
End If
End Sub

Private Sub chkZeroFill_Click()
If chkZeroFill.Value = vbChecked Then
    strZeroFill = " ZEROFILL"
Else
    strZeroFill = ""
End If
End Sub

Private Sub cmdAddCol_Click()
Dim strDataType As String
Dim RepeatedColName As Boolean
Dim Lindex As Integer

For Lindex = 0 To cboSelectPrimaryKey.ListCount
    If txtColumnName = cboSelectPrimaryKey.List(Lindex) Then
        'there are two columns with the same name
        RepeatedColName = True
        Lindex = cboSelectPrimaryKey.ListCount 'To exit the For loop
    End If
    DoEvents
Next Lindex

If RepeatedColName = False Then
    strDataType = cboDataType.List(cboDataType.ListIndex)

    Select Case strDataType
        Case "CHAR"
            CheckFieldLength
        Case "VARCHAR"
            CheckFieldLength
        Case "ENUM"
            CheckFieldLength
        Case "SET"
            CheckFieldLength
        Case Else
            AddColumn
    End Select
Else
    Call MsgBox("There is already a column named " _
            & txtColumnName & ", please enter another Column Name.", vbExclamation + vbOKOnly, Me.Caption)
    txtColumnName = ""
    txtColumnName.SetFocus
End If
CleanColProperties
End Sub

Private Sub CheckFieldLength()
If txtLength = "" Then
    Call MsgBox("You need to put the lenght or elements for " _
            & cboDataType.List(cboDataType.ListIndex) & ".", _
              vbOKOnly + vbExclamation, Me.Caption)
    txtLength.SetFocus
Else
    AddColumn
End If
End Sub

Private Sub cmdApplyChanges_Click()
Dim strSQLtoDB As String
Dim ColListIndex As Integer
Dim responce As VbMsgBoxResult

On Error GoTo createERROR
If txtNewTableName <> "" Then
    strSQLtoDB = "CREATE TABLE " & txtNewTableName & "("
    For ColListIndex = 0 To lstColumnList.ListCount - 1
        strSQLtoDB = strSQLtoDB & " " & lstColumnList.List(ColListIndex) & ", "
        DoEvents
    Next ColListIndex
    strSQLtoDB = Left(strSQLtoDB, Len(strSQLtoDB) - 2) 'para quitar la ultima ", "
    
    If cboSelectPrimaryKey.List(cboSelectPrimaryKey.ListIndex) = "" Then
        responce = MsgBox("There is no Primary Key, Do you wish to select one?, If you select NO, the table will be created without a Primary Key.", vbQuestion + vbYesNo, Me.Caption)
    End If
    
    If responce = vbYes Then
        cboSelectPrimaryKey.SetFocus
        Exit Sub
    End If
    strSQLtoDB = strSQLtoDB & ", PRIMARY KEY(" & cboSelectPrimaryKey.List(cboSelectPrimaryKey.ListIndex) & ")"
    strSQLtoDB = strSQLtoDB & ")"
    DBCon.Execute strSQLtoDB
    frmConnect.PopulatelstTables
    responce = MsgBox("The Database was updated with the new table, do you wish to add another table?, If you answer NO this wildow will close.", vbQuestion + vbYesNo, Me.Caption)
    If responce = vbNo Then
        CleanForm
        Me.Hide
    Else
        CleanForm
    End If
Else
    Call MsgBox("You need to put a table name!!!", vbExclamation + vbOKOnly, Me.Caption)
    txtNewTableName.SetFocus
End If
Exit Sub
createERROR:
    Call MsgBox("ERROR: " & DBCon.Errors(0), vbCritical + vbOKOnly, Me.Caption)
End Sub

Private Sub cmdClose_Click()

Me.Hide
End Sub

Private Sub cmdRemoveColumn_Click()
If lstColumnList.ListCount > 0 Then
cboSelectPrimaryKey.ListIndex = lstColumnList.ListIndex
lstColumnList.RemoveItem (lstColumnList.ListIndex)
cboSelectPrimaryKey.RemoveItem (cboSelectPrimaryKey.ListIndex)
End If
End Sub

Private Sub Form_Load()
lblWorkingWithDB = strDatabase
cboDataType.ListIndex = 1
strNotNull = " NOT NULL"
End Sub

Private Sub lstColumnList_Click()
'lblWorkingWithDB.Caption = lstColumnList.ListIndex
End Sub

Private Sub txtColumnName_Change()
If txtColumnName = " " Then
    txtColumnName = ""
End If
End Sub

Private Sub AddColumn()
Dim strAddCol As String
Dim intDataTypeIndex As Integer
Dim strDataTypeString As String

If txtColumnName <> "" Then
    intDataTypeIndex = cboDataType.ListIndex
    strDataTypeString = cboDataType.List(intDataTypeIndex)
    strAddCol = txtColumnName & " " & strDataTypeString

    If txtLength <> "" Then
        strAddCol = strAddCol & "(" & txtLength & ")"
    End If
    strAddCol = strAddCol & strBinary & strUnsigned & strZeroFill & strNotNull & strAuto_Increment
    
    lstColumnList.AddItem (strAddCol)
    cboSelectPrimaryKey.AddItem (txtColumnName)
Else
    Call MsgBox("The column name cant be blank!", vbOKOnly + vbExclamation, Me.Caption)
End If
End Sub

Private Sub CleanColProperties()
txtColumnName = ""
txtLength = ""
txtDefault = ""
chkUnsigned.Value = vbUnchecked
chkZeroFill.Value = vbUnchecked
chkBinary.Value = vbUnchecked
chkAutoIncrement.Value = vbUnchecked
chkNotNull.Value = vbChecked
cboDataType.ListIndex = 1
End Sub

Private Sub CleanForm()
txtNewTableName = ""
cboSelectPrimaryKey.Clear
lstColumnList.Clear
End Sub
