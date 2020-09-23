VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form3 
   Caption         =   "Main Form"
   ClientHeight    =   8895
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10635
   DrawStyle       =   2  'Dot
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form3"
   Moveable        =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   10635
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3720
      TabIndex        =   11
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   4935
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   3135
      Begin VB.Timer Timer2 
         Left            =   1320
         Top             =   4440
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   4335
         Left            =   120
         ScaleHeight     =   4305
         ScaleWidth      =   2745
         TabIndex        =   4
         Top             =   840
         Width           =   2775
         Begin MSComCtl2.MonthView MonthView1 
            Height          =   2370
            Left            =   0
            TabIndex        =   6
            Top             =   240
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   4180
            _Version        =   393216
            ForeColor       =   -2147483630
            BackColor       =   -2147483633
            Appearance      =   1
            StartOfWeek     =   59899905
            CurrentDate     =   39835
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Label2"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   7
            Top             =   3000
            Width           =   2535
         End
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Task Pane"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Timer Timer1 
      Left            =   2160
      Top             =   4800
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   8640
      Width           =   10635
      _ExtentX        =   18759
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
      MousePointer    =   2
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   480
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":15ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":92C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":14F63
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":17E59
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":42F94
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":46B0C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   780
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   30000
      _ExtentX        =   52917
      _ExtentY        =   1376
      ButtonWidth     =   1693
      ButtonHeight    =   1376
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Aircraft"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Open"
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Flight"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Passengers"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Flight Crew"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Route"
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Route"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Show / Hide"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "* Put the Data Base of Air-Reservation (P)(mdb) File in C:\Program Files\Microsoft visual studio\VB98\"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1575
      Left            =   3360
      TabIndex        =   9
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label rty 
      BackStyle       =   0  'Transparent
      Caption         =   "Note :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3360
      TabIndex        =   8
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "* Others wise the database function will not work !!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   3360
      TabIndex        =   10
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Image Picture2 
      Height          =   11805
      Left            =   3240
      Picture         =   "MainForm.frx":47CB6
      Stretch         =   -1  'True
      Top             =   840
      Width           =   15360
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Index           =   1
      Begin VB.Menu Pass 
         Caption         =   "Change Password"
         Index           =   2
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit Application"
         Index           =   4
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
      Index           =   5
      Begin VB.Menu About 
         Caption         =   "About"
         Index           =   6
      End
      Begin VB.Menu Faraz 
         Caption         =   "-"
         Index           =   7
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' ''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Originaly Designed and Created By M.Faraz Shaikh a Student of 9th Standard Sukkur Pakistan '''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub About_Click(Index As Integer)
Form4.Show
End Sub

Private Sub Command1_Click()
Frame1.Visible = False
Picture2.Left = 0
Picture2.Width = 15360
Picture2.Stretch = True
End Sub

Private Sub Command2_Click()
rty.Visible = False
Label3.Visible = False
Label4.Visible = False
Command2.Visible = False

End Sub

Private Sub Exit_Click(Index As Integer)
End
End Sub
Private Sub Form_Load()

Toolbar1.Width = Me.Width
Frame1.Height = Me.Height - 1900
Timer2.Interval = 100
If MsgBox("Warning: This Form contain some tool which can make you confuse to over come this problem Please Go to Menu-Help-About. What Would You Like To Do Now?", _
              vbYesNo + vbQuestion, _
              "Warning") = vbNo Then
        'Do nothing
        Else
        Form4.Show
        Form3.Hide
        End If
End Sub

Private Sub Form_Resize()
Toolbar1.Width = Me.Width
Frame1.Height = Me.Height - 1550
Picture1.Height = Me.Frame1.Height - 1000
Picture2.Height = Me.Height - 1600
Picture2.Width = Me.Width
End Sub

Private Sub Pass_Click(Index As Integer)
Form17.Show
End Sub



Private Sub Timer2_Timer()
Label2.Caption = Format(Now(), "HH:MM:SS")
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Index = 1 Then
Form6.Show
End If
If Button.Index = 4 Then
Form6.Show
End If
If Button.Index = 5 Then
Form9.Show
End If
If Button.Index = 3 Then
Form10.Show
End If
If Button.Index = 7 Then
Form7.Show
End If
If Button.Index = 9 Then
Form8.Show
End If
If Button.Index = 13 Then
End
End If

If Button.Index = 11 Then
If Frame1.Visible = False Then
    Frame1.Visible = True
    Picture2.Left = 3200
    Picture2.Width = 12160
Else
    Frame1.Visible = False
    Picture2.Left = 0
    Picture2.Width = 15360
End If
End If

End Sub
