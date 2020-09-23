VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form10 
   BorderStyle     =   0  'None
   ClientHeight    =   8325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10020
   LinkTopic       =   "Form5"
   Picture         =   "Fares (frm).frx":0000
   ScaleHeight     =   8325
   ScaleWidth      =   10020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
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
      Height          =   345
      Left            =   5040
      Locked          =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   32
      ToolTipText     =   "Enter The Capacity of Passengers : "
      Top             =   2040
      Visible         =   0   'False
      Width           =   2295
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
      Picture         =   "Fares (frm).frx":1D15E2
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   6600
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
      Picture         =   "Fares (frm).frx":1D42A4
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Click Here To Delete Record"
      Top             =   6600
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
      Picture         =   "Fares (frm).frx":1D70B1
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Go Through The Main Form"
      Top             =   6600
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
      Picture         =   "Fares (frm).frx":1D9F1C
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Exit The Application"
      Top             =   6600
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
      Picture         =   "Fares (frm).frx":1DCB1E
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Click Here To Save Record"
      Top             =   6600
      Width           =   1095
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
      Picture         =   "Fares (frm).frx":1DFA79
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Click Here To Add New Record"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      DataField       =   "Pcap"
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
      Height          =   285
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   3240
      Width           =   2295
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      DataField       =   "Crno"
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
      Height          =   285
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   23
      ToolTipText     =   "Enter The Crew No of Plane"
      Top             =   3720
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      DataField       =   "Rno"
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
      Height          =   285
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   21
      ToolTipText     =   "Enter The Route No of Plane"
      Top             =   4200
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      DataField       =   "Aid"
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
      Height          =   285
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   19
      ToolTipText     =   "Please Enter The Aircrafy ID"
      Top             =   2760
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   11
      ToolTipText     =   "Enter The Flight No"
      Top             =   2760
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "Departure"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "Fares (frm).frx":1E284F
      Left            =   2880
      List            =   "Fares (frm).frx":1E2877
      Locked          =   -1  'True
      TabIndex        =   10
      ToolTipText     =   "Enter The City From Which The Plane Had Departed"
      Top             =   3240
      Width           =   2295
   End
   Begin VB.ComboBox Combo3 
      DataField       =   "Arrival"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "Fares (frm).frx":1E28E5
      Left            =   2880
      List            =   "Fares (frm).frx":1E290D
      Locked          =   -1  'True
      TabIndex        =   9
      ToolTipText     =   "Enter The City From Which The Plane Had Arrived"
      Top             =   3720
      Width           =   2295
   End
   Begin VB.Timer Timer2 
      Left            =   480
      Top             =   6600
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   9480
      Picture         =   "Fares (frm).frx":1E297B
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Command12 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Height          =   255
      Left            =   9000
      Picture         =   "Fares (frm).frx":1E30BD
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   0
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Left            =   480
      Top             =   7560
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      DataField       =   "Deptime"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2880
      TabIndex        =   13
      ToolTipText     =   "Enter The Time on Which The Plane Had Departed"
      Top             =   5160
      Width           =   2295
      _ExtentX        =   4048
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
      Format          =   20709378
      CurrentDate     =   39836
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      DataField       =   "Arrtime"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2880
      TabIndex        =   14
      ToolTipText     =   "Enter The City on Which The Plane Had Arrived"
      Top             =   5640
      Width           =   2295
      _ExtentX        =   4048
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
      Format          =   20709378
      CurrentDate     =   39836
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      DataField       =   "Depdate"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2880
      TabIndex        =   16
      ToolTipText     =   "Enter The Date on Which The Plane Had Departed"
      Top             =   4200
      Width           =   2295
      _ExtentX        =   4048
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
   Begin MSComCtl2.DTPicker DTPicker4 
      DataField       =   "Arrdate"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2880
      TabIndex        =   17
      ToolTipText     =   "Enter The Date on Which The Plane Had Arrived"
      Top             =   4680
      Width           =   2295
      _ExtentX        =   4048
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
      ToolTipText     =   "Click Arrrows To Move Records"
      Top             =   7320
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
      Connect         =   $"Fares (frm).frx":1E3597
      OLEDBString     =   $"Fares (frm).frx":1E3637
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Flight"
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
   Begin VB.Label Label13 
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
      Left            =   2040
      TabIndex        =   33
      Top             =   2040
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Pilot Caption : "
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
      Left            =   5400
      TabIndex        =   24
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Crew No : "
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
      Left            =   5400
      TabIndex        =   22
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Route No : "
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
      Left            =   5400
      TabIndex        =   20
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Aircraft ID : "
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
      Left            =   5400
      TabIndex        =   18
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Arrival Time : "
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
      TabIndex        =   15
      Top             =   5640
      Width           =   2415
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Flight"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   4440
      TabIndex        =   12
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Arrival Date : "
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
      Top             =   4680
      Width           =   2415
   End
   Begin VB.Label Label5 
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
      Left            =   960
      TabIndex        =   7
      Top             =   1200
      Width           =   6975
   End
   Begin VB.Image Image1 
      Height          =   1350
      Left            =   7920
      Picture         =   "Fares (frm).frx":1E36D7
      Top             =   1080
      Width           =   1350
   End
   Begin VB.Image Image5 
      Height          =   705
      Left            =   0
      Picture         =   "Fares (frm).frx":1E3FBB
      Top             =   0
      Width           =   10020
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Departure City : "
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
      TabIndex        =   4
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Arrival City : "
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
      TabIndex        =   3
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label Label4 
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
      TabIndex        =   2
      Top             =   4200
      Width           =   2655
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Departure Time : "
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
      TabIndex        =   1
      Top             =   5160
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Flight No : "
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
      TabIndex        =   0
      Top             =   2760
      Width           =   1695
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim B As Integer

Private Sub Command11_Click()
Unload Me
End Sub


Private Sub Command12_Click()
Form10.WindowState = 1
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



Private Sub Image2_Click()

End Sub

Private Sub Command6_Click()
If MsgBox("Are you sure That you want to add this Record Unless you will be Unable to Change it?", _
              vbYesNo + vbQuestion, _
              "Confirm") = vbNo Then
        'Do nothing
        Else
        Adodc1.Recordset.Update
MsgBox ("The Current Record is Successfully Saved by your Administrator")
Text4.Locked = True
Text1.Locked = True
Text2.Locked = True
Text6.Locked = True
Text5.Locked = True
Text3.Locked = True
Combo1.Locked = True

Combo3.Locked = True
Label13.Visible = False
Text4.Visible = False
End If
End Sub

Private Sub Command7_Click()
Text4.Locked = False
Text1.Locked = False
Text2.Locked = False
Text6.Locked = False
Text5.Locked = False
Text3.Locked = False
Combo1.Locked = False

Combo3.Locked = False
Label13.Visible = True
Text4.Visible = True
Adodc1.Recordset.AddNew
End Sub



Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label13.Visible = True
Text4.Visible = True
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label13.Visible = False
Text4.Visible = False
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

Private Sub Text5_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

