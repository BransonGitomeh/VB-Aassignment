VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6795
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   8430
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox LilstOneCol2 
      Height          =   1425
      Left            =   1080
      TabIndex        =   8
      Top             =   4920
      Width           =   5535
   End
   Begin VB.TextBox txtSearch 
      Height          =   495
      Left            =   2160
      TabIndex        =   7
      Top             =   3720
      Width           =   3615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "search"
      Height          =   495
      Left            =   2400
      TabIndex        =   6
      Top             =   4320
      Width           =   2655
   End
   Begin VB.ListBox lstOneCol 
      Height          =   1815
      Left            =   840
      TabIndex        =   5
      Top             =   1680
      Width           =   5535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Insert"
      Height          =   855
      Left            =   6600
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Execute"
      Height          =   855
      Left            =   4680
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Height          =   855
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox txtPassword 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   600
      Width           =   2535
   End
   Begin VB.TextBox txtUsername 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset


Private Sub Command1_Click()
    
    txtUsername = rs.Fields("username")
    txtPassword = rs.Fields("password")

End Sub

Private Sub Command2_Click()
Dim cmd As New ADODB.Command
cmd.CommandText = "CREATE TABLE test(PersonID int,LastName varchar(255));"
cmd.ActiveConnection = con
cmd.Execute
End Sub


Private Sub Command3_Click()
Dim cmd As New ADODB.Command
cmd.CommandText = "INSERT INTO test VALUES (1, 'branson gitomeh');"
cmd.ActiveConnection = con
cmd.Execute
End Sub

Private Sub Command4_Click()
'read from txtSearch'
'construct sql query and send'
'populate listbox with result'
Dim text As String
text = txtSearch.text

Dim myQuery As String
myQuery = "SELECT id,username From users WHERE username = '" + text + "';"
MsgBox (myQuery)

Call OpenConnection

With rs
.ActiveConnection = con
.CursorLocation = adUseClient
.LockType = adLockOptimistic
.Open myQuery
End With

If Not (rs.EOF And rs.BOF) Then
    rs.MoveFirst
    Do While Not rs.EOF
        strID = rs.Fields(0).Value
        strUsername = rs.Fields(1).Value
        'MsgBox (strID)'
        'MsgBox (strUsername)'
        Call AddToListBox2(strID, strUsername)
        rs.MoveNext
    Loop
Else
    MsgBox "No records found."
End If


End Sub

Private Sub form_load()

Call OpenConnection

With rs
.ActiveConnection = con
.CursorType = adOpenDynamic
.CursorLocation = adUseClient
.LockType = adLockOptimistic
.Open "SELECT * FROM users"
End With

If Not (rs.EOF And rs.BOF) Then
    rs.MoveFirst
    Do While Not rs.EOF
        strID = rs.Fields(0).Value
        strName = rs.Fields(1).Value
        strPassword = rs.Fields(2).Value
        Call AddToListBox(strID, strName, strPassword)
        rs.MoveNext
    Loop
Else
    MsgBox "No records found."
End If


End Sub

Private Sub OpenConnection()

    If con.State = 1 Then
        con.Close
    End If
       
    con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & App.Path & "\myAccessFile.accdb"
    con.Open

End Sub

Private Sub AddToListBox(ByVal strID As String, ByVal strName As String, ByVal strPassword As String)
    lstOneCol.AddItem strID + "        " + strName + "         " + strPassword
End Sub

Private Sub AddToListBox2(ByVal strID As String, ByVal strUsername As String)
    LilstOneCol2.AddItem strID + "          " + strUsername
End Sub

