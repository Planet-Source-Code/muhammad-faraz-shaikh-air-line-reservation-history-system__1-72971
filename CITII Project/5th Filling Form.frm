VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form8 
   BorderStyle     =   0  'None
   Caption         =   "Form8"
   ClientHeight    =   7140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   LinkTopic       =   "Form8"
   Picture         =   "5th Filling Form.frx":0000
   ScaleHeight     =   7140
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
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
      Left            =   4800
      Locked          =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   18
      ToolTipText     =   "Enter The Capacity of Passengers : "
      Top             =   2040
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.ComboBox Combo2 
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
      Height          =   330
      ItemData        =   "5th Filling Form.frx":341DEA
      Left            =   4800
      List            =   "5th Filling Form.frx":341E2A
      Locked          =   -1  'True
      TabIndex        =   17
      ToolTipText     =   "Enter The Place of The Departure of Flight : "
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
      Height          =   330
      ItemData        =   "5th Filling Form.frx":341EDF
      Left            =   4800
      List            =   "5th Filling Form.frx":341F1F
      Locked          =   -1  'True
      TabIndex        =   16
      ToolTipText     =   "Enter The Place of The Departure of Flight : Enter The Place of The Departure of Flight : "
      Top             =   3840
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
      Left            =   1440
      MaskColor       =   &H00C0FFFF&
      Picture         =   "5th Filling Form.frx":341FD4
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Click Here To Add Records"
      Top             =   5280
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
      Left            =   2520
      MaskColor       =   &H00C0FFFF&
      Picture         =   "5th Filling Form.frx":344DAA
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Click Here To Save Records"
      Top             =   5280
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
      Left            =   6840
      MaskColor       =   &H00C0FFFF&
      Picture         =   "5th Filling Form.frx":347D05
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "End"
      Top             =   5280
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
      Left            =   5760
      MaskColor       =   &H00C0FFFF&
      Picture         =   "5th Filling Form.frx":34A907
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Go Through The Main Form"
      Top             =   5280
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
      Left            =   4680
      MaskColor       =   &H00C0FFFF&
      Picture         =   "5th Filling Form.frx":34D772
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Click Here To Delete Records"
      Top             =   5280
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
      Left            =   3600
      MaskColor       =   &H00C0FFFF&
      Picture         =   "5th Filling Form.frx":35057F
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5280
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   360
      Left            =   4800
      Locked          =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      ToolTipText     =   "Enter Roue No of Flight : "
      Top             =   2640
      Width           =   2295
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataField       =   "Distance"
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
      Left            =   4800
      Locked          =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      ToolTipText     =   "Just Enter The Route Distance of Flight : "
      Top             =   4440
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Height          =   255
      Left            =   8520
      Picture         =   "5th Filling Form.frx":353241
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   9000
      Picture         =   "5th Filling Form.frx":35371B
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   495
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   450
      Left            =   1440
      Top             =   6000
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
      Connect         =   $"5th Filling Form.frx":353E5D
      OLEDBString     =   $"5th Filling Form.frx":353EFD
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Route"
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
      Left            =   1800
      TabIndex        =   19
      Top             =   2040
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Route"
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
      Left            =   4080
      TabIndex        =   15
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label13 
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
      Left            =   2400
      TabIndex        =   7
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Route Distance : "
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
      Left            =   2400
      TabIndex        =   6
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Route No :"
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
      Left            =   2400
      TabIndex        =   5
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label1 
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
      Left            =   2400
      TabIndex        =   4
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Image Image5 
      Height          =   705
      Left            =   0
      Picture         =   "5th Filling Form.frx":353F9D
      Top             =   0
      Width           =   10020
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "INTERNATIONAL AIR TRANSPORT ASSOCIATION"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   1080
      Width           =   7695
   End
   Begin VB.Image Image1 
      Height          =   1350
      Left            =   7800
      Picture         =   "5th Filling Form.frx":36AFCB
      Top             =   1080
      Width           =   1350
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Me.WindowState = 1
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
Text1.Locked = True
Text2.Locked = True
Text5.Locked = True
Combo2.Locked = True
Combo3.Locked = True
Label6.Visible = False
Text2.Visible = False
End If
End Sub

Private Sub Timer1_Timer()
Ghanta.Caption = FormatDateTime(Date, 1) & vbCrLf & Time()
End Sub

Private Sub Command7_Click()
Text1.Locked = False
Text2.Locked = False
Text5.Locked = False
Combo2.Locked = False
Combo3.Locked = False
Label6.Visible = True
Text2.Visible = True
Adodc1.Recordset.AddNew

End Sub



Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.Visible = True
Text2.Visible = True
End Sub



Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.Visible = False
Text2.Visible = False
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub
