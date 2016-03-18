VERSION 5.00
Begin VB.Form switchboard 
   Caption         =   "Form2"
   ClientHeight    =   8820
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15660
   LinkTopic       =   "Form2"
   ScaleHeight     =   8820
   ScaleWidth      =   15660
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "checkout registered car"
      Height          =   855
      Left            =   9000
      TabIndex        =   1
      Top             =   5520
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "register incoming car"
      Height          =   855
      Left            =   3840
      TabIndex        =   0
      Top             =   5520
      Width           =   4095
   End
End
Attribute VB_Name = "switchboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
incoming.Show

End Sub

Private Sub Command2_Click()
Unload Me
outgoing.Show
End Sub
