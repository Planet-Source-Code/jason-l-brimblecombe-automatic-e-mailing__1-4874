VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2625
   LinkTopic       =   "Form1"
   ScaleHeight     =   825
   ScaleWidth      =   2625
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Email Me!"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Downloaded from Jason's VB Page!
' http://get.to/jasonsvbpage
' Thankyou for visiting and please come again :)

Private Sub Command1_Click()
On Error GoTo 100

    Dim RetVal As Long
    
    RetVal = Shell("start mailto:jbrimble@hotmail.com", 0)
    
100
Exit Sub

End Sub
