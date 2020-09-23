VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form9 
   BorderStyle     =   0  'None
   Caption         =   "Please Enter The Address of Passenger"
   ClientHeight    =   9150
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   10005
   LinkTopic       =   "Form4"
   Picture         =   "Passengers (frm).frx":0000
   ScaleHeight     =   9150
   ScaleWidth      =   10005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text9 
      DataField       =   "Phone1"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   6840
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   45
      ToolTipText     =   "Enter another Phone No of Passenger If U Have"
      Top             =   5880
      Width           =   2535
   End
   Begin VB.TextBox Text7 
      DataField       =   "Phone"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   6840
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   44
      ToolTipText     =   "Enter The Phone No of Passenger"
      Top             =   3000
      Width           =   2535
   End
   Begin VB.TextBox Text6 
      DataField       =   "Address"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   2040
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   43
      ToolTipText     =   "Please Enter The Address of Passenger"
      Top             =   4440
      Width           =   2295
   End
   Begin VB.TextBox Text11 
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
      Left            =   5160
      Locked          =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   41
      ToolTipText     =   "Enter The Capacity of Passengers : "
      Top             =   1920
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      DataField       =   "Class"
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
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   40
      ToolTipText     =   "Enter The Type of Class Which You are Giving To Passenger in Plane"
      Top             =   6360
      Width           =   2535
   End
   Begin VB.TextBox Text10 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      DataField       =   "Fare"
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
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   38
      ToolTipText     =   "Enter The PNR of Passenger"
      Top             =   6720
      Width           =   2295
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      DataField       =   "Tno"
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
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   2520
      Width           =   2535
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
      Picture         =   "Passengers (frm).frx":1F118A
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   7440
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
      Picture         =   "Passengers (frm).frx":1F3E4C
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   7440
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
      Picture         =   "Passengers (frm).frx":1F6C59
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   7440
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
      Picture         =   "Passengers (frm).frx":1F9AC4
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   7440
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
      Left            =   2640
      MaskColor       =   &H00C0FFFF&
      Picture         =   "Passengers (frm).frx":1FC6C6
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   7440
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
      Left            =   1560
      MaskColor       =   &H00C0FFFF&
      Picture         =   "Passengers (frm).frx":1FF621
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   7440
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      DataField       =   "PNR"
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
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   29
      ToolTipText     =   "Enter The PNR of Passenger"
      Top             =   2520
      Width           =   2295
   End
   Begin VB.ComboBox Combo3 
      DataField       =   "Gender"
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
      ItemData        =   "Passengers (frm).frx":2023F7
      Left            =   2040
      List            =   "Passengers (frm).frx":202401
      Locked          =   -1  'True
      TabIndex        =   27
      ToolTipText     =   "Enter Gender of Passenger"
      Top             =   6120
      Width           =   2295
   End
   Begin VB.ComboBox Combo2 
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
      ItemData        =   "Passengers (frm).frx":202413
      Left            =   2040
      List            =   "Passengers (frm).frx":202453
      Locked          =   -1  'True
      TabIndex        =   26
      ToolTipText     =   "Enter The Name of City From Which The Had Arrival"
      Top             =   5520
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
      ItemData        =   "Passengers (frm).frx":202508
      Left            =   2040
      List            =   "Passengers (frm).frx":202548
      Locked          =   -1  'True
      TabIndex        =   25
      ToolTipText     =   "Enter The Name of City From Which The Had Departed"
      Top             =   4920
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      DataField       =   "Age"
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
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   19
      ToolTipText     =   "Enter The Age of Passenger"
      Top             =   3960
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      DataField       =   "Name"
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
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   18
      ToolTipText     =   "Enter The Name of Passenger"
      Top             =   3480
      Width           =   2295
   End
   Begin VB.TextBox Text2 
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
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   17
      ToolTipText     =   "Enter The Flight No"
      Top             =   3000
      Width           =   2295
   End
   Begin VB.CommandButton Command12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Height          =   255
      Left            =   9000
      Picture         =   "Passengers (frm).frx":2025FD
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   9480
      Picture         =   "Passengers (frm).frx":202AD7
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Left            =   360
      Top             =   9480
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      DataField       =   "Depdate"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   6840
      TabIndex        =   21
      ToolTipText     =   "Enter The Departed Date of Passenger"
      Top             =   3480
      Width           =   2520
      _ExtentX        =   4445
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
   Begin MSComCtl2.DTPicker DTPicker1 
      DataField       =   "Deptime"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   6840
      TabIndex        =   22
      ToolTipText     =   "Enter The Departed Time of Passenger"
      Top             =   4080
      Width           =   2520
      _ExtentX        =   4445
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
      Format          =   20709378
      CurrentDate     =   39820
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      DataField       =   "Arrdate"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   6840
      TabIndex        =   23
      ToolTipText     =   "Enter The Arrived Date of Passenger"
      Top             =   4680
      Width           =   2520
      _ExtentX        =   4445
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
   Begin MSComCtl2.DTPicker DTPicker4 
      DataField       =   "Arrtime"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   6840
      TabIndex        =   24
      ToolTipText     =   "Enter The Arrived Time of Passenger"
      Top             =   5280
      Width           =   2520
      _ExtentX        =   4445
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
      Format          =   20709378
      CurrentDate     =   39820
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   450
      Left            =   1560
      Top             =   8160
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
      Connect         =   $"Passengers (frm).frx":203219
      OLEDBString     =   $"Passengers (frm).frx":2032B9
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Passenger"
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
   Begin VB.Label Label18 
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
      TabIndex        =   42
      Top             =   1920
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fare : "
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
      Index           =   2
      Left            =   360
      TabIndex        =   39
      Top             =   6720
      Width           =   855
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "T-No : "
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
      Left            =   4920
      TabIndex        =   37
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Passengers"
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
      Left            =   3960
      TabIndex        =   28
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label11 
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
      Left            =   840
      TabIndex        =   20
      Top             =   1080
      Width           =   6735
   End
   Begin VB.Image Image2 
      Height          =   1350
      Left            =   8040
      Picture         =   "Passengers (frm).frx":203359
      Top             =   720
      Width           =   1350
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Depatrue City : "
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
      Left            =   360
      TabIndex        =   16
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Class : "
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
      Left            =   4920
      TabIndex        =   15
      Top             =   6360
      Width           =   735
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Gender : "
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
      Left            =   360
      TabIndex        =   14
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Label Label14 
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
      Left            =   360
      TabIndex        =   13
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label Label7 
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
      Left            =   4920
      TabIndex        =   12
      Top             =   4080
      Width           =   1935
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
      Left            =   4920
      TabIndex        =   11
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Arrical Date : "
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
      Left            =   4920
      TabIndex        =   10
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Address : "
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
      Left            =   360
      TabIndex        =   9
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone 1 : "
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
      Left            =   4920
      TabIndex        =   8
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label Label2 
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
      Left            =   360
      TabIndex        =   7
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Name : "
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
      Left            =   360
      TabIndex        =   6
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone : "
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
      Left            =   4920
      TabIndex        =   5
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Age : "
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
      Left            =   360
      TabIndex        =   4
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label6 
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
      Left            =   4920
      TabIndex        =   3
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PNR : "
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
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   2520
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   705
      Left            =   0
      Picture         =   "Passengers (frm).frx":203C3D
      Top             =   0
      Width           =   10020
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim B As Integer

Private Sub Command11_Click()
Unload Me
End Sub

Private Sub Command12_Click()
Me.WindowState = 1
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
Label18.Visible = False
Text11.Visible = False

Text1.Locked = True
Text2.Locked = True
Text3.Locked = True
Text4.Locked = True
Text5.Locked = True
Text6(7).Locked = True
Text7(0).Locked = True
Text8.Locked = True
Text9(1).Locked = True
Text10.Locked = True
Text11.Locked = True
Combo1.Locked = True
Combo2.Locked = True
Combo3.Locked = True
End If
End Sub

Private Sub Command7_Click()
Text1.Locked = False
Text2.Locked = False
Text3.Locked = False
Text4.Locked = False
Text5.Locked = False
Text6(7).Locked = False
Text7(0).Locked = False
Text8.Locked = False
Text9(1).Locked = False
Text10.Locked = False
Text11.Locked = False
Combo1.Locked = False
Combo2.Locked = False
Combo3.Locked = False
Label18.Visible = True
Text11.Visible = True
Adodc1.Recordset.AddNew
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label18.Visible = True
Text11.Visible = True
End Sub

Private Sub Label18_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label18.Visible = False
Text11.Visible = False
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub
