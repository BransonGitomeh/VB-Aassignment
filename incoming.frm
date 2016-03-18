VERSION 5.00
Begin VB.Form incoming 
   Caption         =   "incoming"
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
      Left            =   12480
      TabIndex        =   3
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "save record"
      Height          =   975
      Left            =   6840
      TabIndex        =   2
      Top             =   4680
      Width           =   4815
   End
   Begin VB.TextBox txtPlate 
      Height          =   1215
      Left            =   4680
      TabIndex        =   0
      Top             =   1560
      Width           =   9015
   End
   Begin VB.Label Label1 
      Caption         =   "enter licence plate"
      Height          =   1095
      Left            =   5640
      TabIndex        =   1
      Top             =   3120
      Width           =   6975
   End
End
Attribute VB_Name = "incoming"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub Command1_Click()

Call OpenConnection

Dim cmd As New ADODB.Command
Dim plate As String
Dim query As String
plate = txtPlate.text
query = "INSERT INTO incoming VALUES ('" + plate + "', TimeValue(Now))"
'MsgBox (query)'

cmd.CommandText = query
cmd.ActiveConnection = con
cmd.Execute

MsgBox ("car registered successfully")

Unload Me
switchboard.Show

End Sub

Private Sub Command2_Click()
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

