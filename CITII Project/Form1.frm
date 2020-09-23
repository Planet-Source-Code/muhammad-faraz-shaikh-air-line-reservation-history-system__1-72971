VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   488
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7320
      Left            =   0
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   488
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   550
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   8250
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       SAKI_Intitute of Science and Teachnology Sukkur Project For CIT Final         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hdcDest As Long, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Type SHITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

Public Sub AlwaysOnTop(myfrm As Form, SetOnTop As Boolean)
    If SetOnTop Then
        lFlag = HWND_TOPMOST
    Else
        lFlag = HWND_NOTOPMOST
    End If
    'SetWindowPos myfrm.hwnd, lFlag, _
    myfrm.Left / Screen.TwipsPerPixelX, _
    myfrm.Top / Screen.TwipsPerPixelY, _
    myfrm.Width / Screen.TwipsPerPixelX, _
    myfrm.Height / Screen.TwipsPerPixelY, _
    SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub



Public Function GetDesktop(FRM As Form)
    Dim HW As Long
    Dim HA As Long
    Dim iLeft As Integer
    Dim iTop As Integer
    Dim iWidth As Integer
    Dim iHeight As Integer
    FRM.AutoRedraw = True
    FRM.Show
    FRM.Hide
    
    DoEvents
    HA = GetDC(GetDesktopWindow())
    iLeft = FRM.Left / Screen.TwipsPerPixelX
    iTop = FRM.Top / Screen.TwipsPerPixelY
    iWidth = FRM.ScaleWidth
    iHeight = FRM.ScaleHeight
    Call BitBlt(FRM.hDC, 0, 0, iWidth, iHeight, HA, iLeft, iTop, vbSrcCopy)
    FRM.Picture = FRM.Image
    
    FRM.Show

End Function

Public Function Gerade(Number) As Boolean 'Function to see if number is dividable by 2
If Round(Number / 2, 0) = Number / 2 Then
    Gerade = True
Else
    Gerade = False
End If
End Function

Public Sub Pause(Delay)
Dim StartTime
    StartTime = GetTickCount
    Do
    Loop Until StartTime + Delay < GetTickCount
End Sub

Private Sub Form_Load()
Dim H, W, col As Long
    
    Me.Width = Picture1.Width * 15 'Size the Splashscreenform to
    Me.Height = Picture1.Height * 15 ' fit the Picture

    Call GetDesktop(Me) 'Make Screenshot of Screen behind frmSplash
                        'and copy it to the form
    Me.Show

    For W = 0 To Picture1.ScaleWidth 'This is the effect itself
        For H = 0 To Picture1.ScaleHeight
            If Gerade(W) = Gerade(H) Then
            ' ^ Makes sure only every second Pixel is shown
                col = GetPixel(Picture1.hDC, W, H)
                SetPixel Me.hDC, destX + W, destY + H, col
            Else
                col = GetPixel(Picture1.hDC, Picture1.ScaleWidth - W, H)
                SetPixel Me.hDC, destX + Picture1.ScaleWidth - W, destY + H, col
            End If
        Next H
    
        Me.Refresh
        DoEvents
        Pause 1 'Here you can change the speed of the effect
    Next W
    
    Sleep 3000
    
    Form2.Show vbModeless
    Unload Me
End Sub



