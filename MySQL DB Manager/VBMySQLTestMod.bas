Attribute VB_Name = "VBMySQLTestMod"
'This program was made by José M. Nieves (nievesj@prtc.net)
'You can use it in any way you see it usefull, just give me some credit.
'I made this because I were bored so dont expect the code to be elegant or professional :), it just works.

Option Explicit
Public DBCon As Connection 'connection to DB
Public RecSet As New Recordset 'Recordset

Public strHost As String
Public strDatabase As String
Public strUser As String
Public strPass As String
Public DBConErrorIndex As Integer

Public Sub AbrirConeccion()
Dim strConString As String
'close open conection
If Not (DBCon Is Nothing) Then
    DBCon.Close
    Set DBCon = Nothing
End If
'create new conection
Set DBCon = New Connection
strConString = "DRIVER={MySQL ODBC 3.51 Driver};" _
                        & "SERVER=" & strHost & ";" _
                        & "DATABASE=" & strDatabase & ";" _
                        & "UID=" & strUser & ";PWD=" & strPass & "; OPTION=35"
'On Error GoTo conERROR
With DBCon
    .ConnectionString = strConString
    .Open
End With
'Exit Sub
conERROR:
'Call MsgBox("ERROR: " & DBCon.Errors(0), vbCritical + vbOKOnly, "Connection Error")
End Sub

Public Sub CerrarConeccion()
On Error GoTo NoHacerNada
DBCon.Close
Set DBCon = Nothing
Exit Sub
NoHacerNada:
Exit Sub
End Sub

Public Sub SalirDelPrograma()
'this sub takes all forms out of memory
Dim Form As Form
   For Each Form In Forms
      Unload Form
      Set Form = Nothing
   Next Form
End
End Sub

