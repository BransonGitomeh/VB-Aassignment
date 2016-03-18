VERSION 5.00
Begin VB.Form outgoing 
   Caption         =   "Form2"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "back"
      Height          =   495
      Left            =   9480
      TabIndex        =   2
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "check this car out"
      Height          =   495
      Left            =   7680
      TabIndex        =   1
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox txtPlate 
      Height          =   1095
      Left            =   5520
      TabIndex        =   0
      Top             =   1800
      Width           =   5895
   End
End
Attribute VB_Name = "outgoing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub Command2_Click()
Unload Me
switchboard.Show
End Sub

Private Sub Command1_Click()

Call OpenConnection

Dim cmd As New ADODB.Command
Dim plate As String
Dim query As String
plate = txtPlate.text
query = "SELECT plate,timeIn From incoming WHERE plate = '" + plate + "';"
'MsgBox (query)'

With rs
.ActiveConnection = con
.CursorLocation = adUseClient
.LockType = adLockOptimistic
.Open query
End With

Dim lMinutes As Long
lMinutes = DateDiff("n", rs.Fields("timeIn"), TimeValue(Now))

'multiply this with cost per minute'

MsgBox (lMinutes)

Unload Me
switchboard.Show

End Sub
'open the db connection'
Private Sub OpenConnection()

    If con.State = 1 Then
        con.Close
    End If
       
    con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & App.Path & "\myAccessFile.accdb"
    con.Open

End Sub


