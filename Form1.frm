VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1935
   LinkTopic       =   "Form1"
   ScaleHeight     =   1380
   ScaleWidth      =   1935
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      MaxLength       =   5
      TabIndex        =   1
      Text            =   "11768"
      Top             =   360
      Width           =   1215
   End
   Begin Project1.JaysWeather JaysWeather1 
      Left            =   480
      Top             =   960
      _ExtentX        =   1667
      _ExtentY        =   1879
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Zip Code:"
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Command1.Enabled = False
    MsgBox JaysWeather1.GetTemp(Text1)
    Command1.Enabled = True
End Sub
