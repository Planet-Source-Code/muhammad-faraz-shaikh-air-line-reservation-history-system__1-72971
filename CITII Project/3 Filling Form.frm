Type=Exe
Form=Form1.frm
Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\WINNT\system32\stdole2.tlb#OLE Automation
Form=Log.frm
Form=MainForm.frm
Object={831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0; mscomctl.ocx
Object={86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0; MSCOMCT2.OCX
Form=3 Filling Form.frm
Form=4 Fillling Form.frm
Form=5th Filling Form.frm
Reference=*\G{00025E01-0000-0000-C000-000000000046}#4.0#0#C:\Program Files\Common Files\Microsoft Shared\DAO\DAO350.DLL#Microsoft DAO 3.51 Object Library
Object={67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0; MSADODC.OCX
Reference=*\G{00000200-0000-0010-8000-00AA006D2EA4}#2.0#0#C:\Program Files\Common Files\System\ADO\msado20.tlb#Microsoft ActiveX Data Objects 2.0 Library
Form=Passengers (frm).frm
Form=Fares (frm).frm
Reference=*\G{56BF9020-7A2F-11D0-9482-00A0C91110ED}#1.0#0#C:\WINNT\system32\MSBIND.DLL#Microsoft Data Binding Collection
IconForm="Form1"
Startup="Form1"
HelpFile=""
Title="Project1"
ExeName32="Project1.exe"
Path32=".."
Command32=""
Name="Project1"
HelpContextID="0"
CompatibleMode="0"
MajorVer=1
MinorVer=0
RevisionVer=0
AutoIncrementVer=0
ServerSupportFiles=0
VersionCompanyName="Home"
CompilationType=0
OptimizationType=0
FavorPentiumPro(tm)=0
CodeViewDebugInfo=0
NoAliasing=0
BoundsCheck=0
OverflowCheck=0
FlPointCheck=0
FDIVCheck=0
UnroundedFP=0
StartMode=0
Unattended=0
Retained=0
ThreadPerObject=0
MaxNumberOfThreads=1
DebugStartupOption=0

[MS Transaction Server]
AutoRefresh=1
                  ItemData        =   "5th Filling Form.frx":341DEA
      Left            =   4800
      List            =   "5th Filling Form.frx":341DF4
      TabIndex        =   16
      Top             =   3240
      Width           =   2295
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      MaskColor       =   &H00C0FFFF&
      Picture         =   "5th Filling Form.frx":341E06
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      MaskColor       =   &H00C0FFFF&
      Picture         =   "5th Filling Form.frx":344BDC
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6840
      MaskColor       =   &H00C0FFFF&
      Picture         =   "5th Filling Form.frx":347B37
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      MaskColor       =   &H00C0FFFF&
      Picture         =   "5th Filling Form.frx":34A739
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      MaskColor       =   &H00C0FFFF&
      Picture         =   "5th Filling Form.frx":34D5A4
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      MaskColor       =   &H00C0FFFF&
      Picture         =   "5th Filling Form.frx":3503B1
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4440
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataField       =   "Rno"
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
      Height          =   360
      Left            =   4800
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataField       =   "Distance"
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
      Height          =   360
      Left            =   4800
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   3840
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Height          =   255
      Left            =   8520
      Picture         =   "5th Filling Form.frx":353073
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   9000
      Picture         =   "5th Filling Form.frx":35354D
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   495
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   450
      Left            =   1440
      Top             =   5160
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
      Connect         =   $"5th Filling Form.frx":353C8F
      OLEDBString     =   $"5th Filling Form.frx":353D34
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
      Caption         =   "Arrival : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Route Distance : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Route No :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Departure : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Image Image5 
      Height          =   705
      Left            =   0
      Picture         =   "5th Filling Form.frx":353DD9
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
      Picture         =   "5th Filling Form.frx":36AE07
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

Private Sub Command6_Click()
Adodc1.Recordset.Update
End Sub

Private Sub Timer1_Timer()
Ghanta.Caption = FormatDateTime(Date, 1) & vbCrLf & Time()
End Sub

Private Sub Command7_Click()
Adodc1.Recordset.AddNew
End Sub

Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub
                                                                                                                                                                                                                                             �9V�A��L.pe@PFj�t��w,��5�3@�p(g����:�P�%�EL,���Pt�*"D�]�ìCD!��?ҷ�:B��V�5��{��v �
=���G��ݥ8h�����5��PGx�)�>E���L��S5��'L &u�t)%lS9~t^-����K��0-�h��<���B�W�X�� "�����X�)�n O��*X�*Pj�.����@�^�������-�R���GV��R��#X��  �v��N�����f���^�`�6�W A����M�C�R2H%�PM��C��p)��Z0��t

�+sD]���g���f`�p�����z�}/�J�����P��ړVtw��"�����|��N@��6��#�+�TN�o�������;m<
7
�r@�\=#ɾ��k@��l��j �D��d�6\p��E�K������D�sA<�20N�փo��Ȕ���a�0�3@PK.[*�$�֨qz�ՕBĄ�7����Ky	������6�q� ��Pካ(��@�9�$��O����IO����[#�|��K/-��F�L@����큄@ t20t$@C��D��A_Ay��Z�O����+�(��޺�%��#��^om�W�/t(���8�(�jtJ=9)����.JA��x��H.�T#�r@b�!�&���D��G�ǲ�:-%8���B: (4�Che5�[������_���;�%���gzW@	��l�,^'�!����m��U�Ox����ԛ�����V&�������T�����0���[Ro�;UB1$$u��+!�L�J�E	��\�	��e�H�Du,�(LI� �"Lh!�+E��@q���8u[K��$�Z��<�� �)=BG7d�vA��� �}v��d���X쉟��u�W�>[�T�lF�M�ظ��T�4��F�Mڀ�F�0*�r	8IaG4Aq]�U�ւZk�@�� _X��u950��L���fH�BdE�>4V�X,�6S����$�1�8v�8=|�P- S^���Y1� 2�����������1����J=N�@�uNV�#��J�3�(O��3�(J8m	L&���fj"�5�6v���h�D���h���>Y��VtOk��(ɭF���杩i��N	o�4�D��a����ԑ�p+":���ZH<��&�Q������-۝�Dd0"�`);l�v%!��]� ���N�"�S����W|v"�V����h�o��S1��}���V�V<� �H;}=��(�W_Sk|9����3\�	)j`�pژ�H�Yj5Dnuqn�^$y�����gM�(`H5Zg+��^��)m�x��4%Ge��jSs���������t�����#��D5�0܊�*٬�E ��m�ˀ�,��=��t�[t|h���2�!x����J��+��Hh�g{G��h�Sub Command3_Click()
End
End Sub

Private Sub Command4_Click()
Form3.Show
Form6.Hide
End Sub

Private Sub Command5_Click()
Adodc1.Recordset.Delete
End Sub

Private Sub Timer1_Timer()
Ghanta.Caption = FormatDateTime(Date, 1) & vbCrLf & Time()
End Sub

Private Sub Image2_Click()

End Sub

Private Sub Command6_Click()
Unload Me

End Sub

Private Sub Image5_Click()
Me.WindowState = 1
End Sub
