VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form6 
   BackColor       =   &H80000014&
   BorderStyle     =   0  'None
   Caption         =   "4"
   ClientHeight    =   7755
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9975
   FillColor       =   &H00800080&
   LinkTopic       =   "Form6"
   Picture         =   "Air Craft (frm).frx":0000
   ScaleHeight     =   7755
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text5 
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
      Left            =   4920
      Locked          =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   21
      ToolTipText     =   "Enter The Capacity of Passengers : "
      Top             =   2160
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataField       =   "atype"
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
      Left            =   4920
      Locked          =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      ToolTipText     =   "Enter The Capacity of Passengers : "
      Top             =   3960
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
      Left            =   3720
      MaskColor       =   &H00C0FFFF&
      Picture         =   "Air Craft (frm).frx":1D15E2
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6000
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
      Left            =   4800
      MaskColor       =   &H00C0FFFF&
      Picture         =   "Air Craft (frm).frx":1D42A4
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Click Here To Delete Records"
      Top             =   6000
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
      Left            =   5880
      MaskColor       =   &H00C0FFFF&
      Picture         =   "Air Craft (frm).frx":1D70B1
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Go Through The Main Form"
      Top             =   6000
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
      Left            =   6960
      MaskColor       =   &H00C0FFFF&
      Picture         =   "Air Craft (frm).frx":1D9F1C
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "End Application"
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
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
      Left            =   2640
      MaskColor       =   &H00C0FFFF&
      Picture         =   "Air Craft (frm).frx":1DCB1E
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Click Here To Save Records"
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
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
      Left            =   1560
      MaskColor       =   &H00C0FFFF&
      Picture         =   "Air Craft (frm).frx":1DFA79
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Click Here To Add Records"
      Top             =   6000
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataField       =   "Fuelrequirement"
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
      Left            =   4920
      Locked          =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      ToolTipText     =   "Enter The Amount of Fuel Which You Required For Flight : "
      Top             =   5160
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataField       =   "Capacity"
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
      Left            =   4920
      Locked          =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      ToolTipText     =   "Enter The Capacity of Passengers : "
      Top             =   3360
      Width           =   2295
   End
   Begin VB.CommandButton Command7 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Height          =   255
      Left            =   9000
      Picture         =   "Air Craft (frm).frx":1E284F
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Minimize"
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   9480
      Picture         =   "Air Craft (frm).frx":1E2D29
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Exit"
      Top             =   0
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   360
      Left            =   4920
      Locked          =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      ToolTipText     =   "Enter The Given Aircraft ID : "
      Top             =   2760
      Width           =   2295
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      DataField       =   "Lastservice"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4920
      TabIndex        =   6
      ToolTipText     =   "Enter The date of Last Service :  "
      Top             =   4560
      Width           =   2280
      _ExtentX        =   4022
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
      DateIsNull      =   -1  'True
      Format          =   20709377
      CurrentDate     =   39820
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   450
      Left            =   1560
      ToolTipText     =   "Click The Arrows To Move Records"
      Top             =   6720
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
      Connect         =   $"Air Craft (frm).frx":1E346B
      OLEDBString     =   $"Air Craft (frm).frx":1E350B
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Aircraft"
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
   Begin VB.Label Label6 
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
      Left            =   1920
      TabIndex        =   20
      Top             =   2160
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Aircraft Form"
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
      Left            =   3840
      TabIndex        =   12
      Top             =   120
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   1350
      Left            =   7680
      Picture         =   "Air Craft (frm).frx":1E35AB
      Top             =   840
      Width           =   1350
   End
   Begin VB.Image Image5 
      Height          =   705
      Left            =   0
      Picture         =   "Air Craft (frm).frx":1E3E8F
      Top             =   0
      Width           =   10020
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
      Left            =   480
      TabIndex        =   7
      Top             =   1200
      Width           =   6735
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Air Type : "
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
      Left            =   1920
      TabIndex        =   4
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Service : "
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
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fuel Requirement : "
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
      Left            =   1800
      TabIndex        =   2
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Label Label7 
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
      Left            =   1920
      TabIndex        =   1
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Capacity : "
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
      Left            =   1920
      TabIndex        =   0
      Top             =   3360
      Width           =   2055
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim B As Integer

Private Sub Command1_Click()
Text5.Locked = False
Text1.Locked = False
Text2.Locked = False
Text3.Locked = False
Text4.Locked = False
Label6.Visible = True
Text5.Visible = True
Adodc1.Recordset.AddNew
End Sub


Private Sub Command2_Click()
If MsgBox("Are you sure That you want to add this Record Unless you will be Unable to Change it?", _
              vbYesNo + vbQuestion, _
              "Confirm") = vbNo Then
        'Do nothing
        Else
        Adodc1.Recordset.Update
       

Label6.Visible = False
Text5.Visible = False
MsgBox ("The Current Record is Successfully Saved by your Administrator")
Text5.Locked = True
Text1.Locked = True
Text2.Locked = True
Text3.Locked = True
Text4.Locked = True
End If
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
Unload Me
End Sub

Private Sub Command7_Click()
Form6.WindowState = 1
End Sub


Private Sub Command9_Click()

Command9.Visible = False
Label6.Visible = False
Text5.Visible = False
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Text5.Visible = True
Label6.Visible = True
End Sub



Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Label6.Visible = False
Text5.Visible = False
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

