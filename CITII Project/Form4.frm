VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9810
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   9810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Next >>"
      Height          =   375
      Left            =   8400
      TabIndex        =   7
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   8400
      TabIndex        =   6
      ToolTipText     =   "If You Have Understood Then Click Here "
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label12 
      Caption         =   "Yes !...."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      TabIndex        =   13
      Top             =   5160
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Label Label9 
      Caption         =   "3. In Searching Data Write Correct Data."
      Height          =   255
      Left            =   3120
      TabIndex        =   10
      Top             =   3000
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.Label Label3 
      Caption         =   $"Form4.frx":030A
      Height          =   735
      Left            =   3120
      TabIndex        =   2
      Top             =   1800
      Width           =   5055
   End
   Begin VB.Label Label11 
      Caption         =   "Welcome in Airline Reservation System : "
      Height          =   495
      Left            =   3120
      TabIndex        =   12
      Top             =   960
      Width           =   6135
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"Form4.frx":03AE
      Height          =   735
      Left            =   3120
      TabIndex        =   5
      Top             =   5160
      Width           =   5055
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"Form4.frx":0470
      Height          =   1455
      Left            =   3120
      TabIndex        =   4
      Top             =   3360
      Width           =   6135
   End
   Begin VB.Label Label4 
      Caption         =   "This Project is Licensed To : "
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   2880
      Width           =   6135
   End
   Begin VB.Label Label1 
      Caption         =   "Reservation System ( SAKI ) 2009."
      Height          =   255
      Left            =   3120
      TabIndex        =   0
      Top             =   240
      Width           =   6135
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  'Fixed Single
      Height          =   1830
      Left            =   120
      Picture         =   "Form4.frx":070A
      Top             =   2160
      Width           =   2730
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1815
      Left            =   120
      Picture         =   "Form4.frx":12AB
      Top             =   4200
      Width           =   2775
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1875
      Left            =   120
      Picture         =   "Form4.frx":2040
      Top             =   120
      Width           =   2760
   End
   Begin VB.Label Label7 
      Caption         =   "To Over Come The Problems You Should Perform Following Step : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   8
      Top             =   240
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.Label Label2 
      Caption         =   $"Form4.frx":2EEC
      Height          =   735
      Left            =   3120
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Label Label8 
      Caption         =   "2. Becarefull While Typing Data You Got i Wrong then You Have To Remove The            Whole Data"
      Height          =   495
      Left            =   3120
      TabIndex        =   9
      Top             =   2040
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.Label Label10 
      Caption         =   "Thank You! For Joining Us. Have a Nice Day."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   11
      Top             =   4320
      Visible         =   0   'False
      Width           =   6375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000011&
      BorderWidth     =   2
      X1              =   0
      X2              =   9840
      Y1              =   4920
      Y2              =   4920
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
Form3.Show
End Sub

Private Sub Command2_Click()
Label1.Visible = False
Label11.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label7.Visible = True
Label2.Visible = True
Label9.Visible = True
Label10.Visible = True
Label8.Visible = True
Label12.Visible = True
Command2.Enabled = False
End Sub


