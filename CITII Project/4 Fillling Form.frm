VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form7 
   BorderStyle     =   0  'None
   Caption         =   "Form7"
   ClientHeight    =   7170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10005
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form7"
   Picture         =   "4 Fillling Form.frx":0000
   ScaleHeight     =   7170
   ScaleWidth      =   10005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataField       =   "date"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5040
      Locked          =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   28
      ToolTipText     =   "Enter The Capacity of Passengers : "
      Top             =   2040
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      MaskColor       =   &H00C0FFFF&
      Picture         =   "4 Fillling Form.frx":1D15E2
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Click Here To Add Records"
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      MaskColor       =   &H00C0FFFF&
      Picture         =   "4 Fillling Form.frx":1D43B8
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Click Here To Save Records"
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7080
      MaskColor       =   &H00C0FFFF&
      Picture         =   "4 Fillling Form.frx":1D7313
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Exit Application"
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      MaskColor       =   &H00C0FFFF&
      Picture         =   "4 Fillling Form.frx":1D9F15
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Go Through The Main Form"
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      MaskColor       =   &H00C0FFFF&
      Picture         =   "4 Fillling Form.frx":1DCD80
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Delete"
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      MaskColor       =   &H00C0FFFF&
      Picture         =   "4 Fillling Form.frx":1DFB8D
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5520
      Width           =   1095
   End
   Begin VB.TextBox Text9 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataField       =   "Hostess4"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7080
      Locked          =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      ToolTipText     =   "Enter 4th Air-Hostess No : "
      Top             =   4680
      Width           =   2055
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataField       =   "Crid"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2520
      Locked          =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      ToolTipText     =   "Enter The Given Crew ID : "
      Top             =   2760
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataField       =   "Pilot"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2520
      Locked          =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      ToolTipText     =   "Enter Piolot No According To Database of Fares : "
      Top             =   3240
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataField       =   "CoPilot"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2520
      Locked          =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      ToolTipText     =   "EnterCo- Piolot No According To Database of Fares : "
      Top             =   3720
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataField       =   "Fno"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2520
      Locked          =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      ToolTipText     =   "Enter Filght No : "
      Top             =   4200
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataField       =   "Hostess2"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7080
      Locked          =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      ToolTipText     =   "Enter 2nd Air-Hostess No : "
      Top             =   3720
      Width           =   2055
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataField       =   "Hostess1"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7080
      Locked          =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      ToolTipText     =   "Enter 1st Air-Hostess No : "
      Top             =   3240
      Width           =   2055
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataField       =   "Hostess3"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7080
      Locked          =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      ToolTipText     =   "Enter 3rd Air-Hostess No : "
      Top             =   4200
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9480
      Picture         =   "4 Fillling Form.frx":1E284F
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9000
      Picture         =   "4 Fillling Form.frx":1E2F91
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   495
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      DataField       =   "Depdate"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2520
      TabIndex        =   27
      ToolTipText     =   "Enter The Time of Departure : "
      Top             =   4680
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   20709377
      CurrentDate     =   39836
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   450
      Left            =   1680
      ToolTipText     =   "Click Arrows To Move Records"
      Top             =   6240
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   794
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   16744576
      ForeColor       =   -2147483643
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"4 Fillling Form.frx":1E346B
      OLEDBString     =   $"4 Fillling Form.frx":1E350B
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "FlightCrew"
      Caption         =   "             CONNECTION STRING"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Date ( MM/DD/YY ) : "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   29
      Top             =   2040
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Air-Hostess No :  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5040
      TabIndex        =   20
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Departure Date : "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   18
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Air-Hostess No :  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5040
      TabIndex        =   17
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Air-Hostess No :  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5040
      TabIndex        =   16
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Air-Hostess No :  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5040
      TabIndex        =   15
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "FLIGHT CREW"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3600
      TabIndex        =   13
      Top             =   120
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   1350
      Left            =   7680
      Picture         =   "4 Fillling Form.frx":1E35AB
      Top             =   960
      Width           =   1350
   End
   Begin VB.Image Image5 
      Height          =   705
      Left            =   0
      Picture         =   "4 Fillling Form.frx":1E3E8F
      Top             =   0
      Width           =   10020
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "International Air Transport Association "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   720
      TabIndex        =   12
      Top             =   1080
      Width           =   6975
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Co-Pilot No  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Crew ID :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pilot No :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   9
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Flight No :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   4200
      Width           =   2055
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim B As Integer

Private Sub Command1_Click()
Form7.WindowState = 1
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
If MsgBox("Are you sure you want to exit form this Software?", _
              vbYesNo + vbQuestion, _
              "Exit") = vbNo Then
        Cancel = 1
        Else
        End
        End If
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
If MsgBox("Are you sure that You want to delete this Record?", _
              vbYesNo + vbQuestion, _
              "Confirm") = vbNo Then
        'Do nothing
        Else
        Adodc1.Recordset.Delete
Adodc1.Recordset.MovePrevious
End If
End Sub

Private Sub Command6_Click()
If MsgBox("Are you sure That you want to add this Record Unless you will be Unable to Change it?", _
              vbYesNo + vbQuestion, _
              "Confirm") = vbNo Then
        'Do nothing
        Else
        Adodc1.Recordset.Update
       

MsgBox ("The Current Record is Successfully Saved by your Administrator")
Text5.Locked = True
Text8.Locked = True
Text7.Locked = True
Text1.Locked = True
Text2.Locked = True
Text3.Locked = True
Text5.Locked = True
Text4.Locked = True
Text6.Locked = True
Text9.Locked = True
Label11.Visible = False
Text8.Visible = False
End If

End Sub


Private Sub Timer2_Timer()
Ghanta.Caption = FormatDateTime(Date, 1) & vbCrLf & Time()
End Sub

Private Sub Command7_Click()
Text8.Locked = False
Text7.Locked = False
Text1.Locked = False
Text2.Locked = False
Text3.Locked = False
Text5.Locked = False
Text4.Locked = False
Text6.Locked = False
Text9.Locked = False
Label11.Visible = True
Text8.Visible = True
Adodc1.Recordset.AddNew
End Sub


Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label11.Visible = True
Text8.Visible = True
End Sub



Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label11.Visible = False
Text8.Visible = False
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub
