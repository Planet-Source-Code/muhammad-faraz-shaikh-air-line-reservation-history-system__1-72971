VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Login Form"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5640
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Log.frx":0000
   ScaleHeight     =   3225
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Height          =   615
      Left            =   4440
      Picture         =   "Log.frx":4220
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   615
      Left            =   3240
      Picture         =   "Log.frx":5038
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      DataField       =   "pass"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2640
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1680
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      DataField       =   "user"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   600
      Width           =   2895
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   0
      Picture         =   "Log.frx":5C77
      ScaleHeight     =   3375
      ScaleWidth      =   2415
      TabIndex        =   0
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = Form17.Text6.Text And Text2.Text = Form17.Text5.Text Then
Form3.Show
Form2.Hide
Else
a = a + 1
MsgBox "Invalid User Name Or Password", vbInformation, "Error"
Text1.Text = ""
Text2.Text = ""
Text1.SetFocus
End If
If a = 4 Then End
End Sub


Private Sub Command2_Click()
End
End Sub
