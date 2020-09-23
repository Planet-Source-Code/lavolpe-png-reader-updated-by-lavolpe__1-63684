VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPNGmenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PNG Load and Options"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3135
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   3135
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdWrite 
      Caption         =   "PNG Writer"
      Height          =   315
      Left            =   45
      TabIndex        =   42
      ToolTipText     =   "Select PNG File"
      Top             =   5340
      Width           =   3060
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4710
      Left            =   30
      TabIndex        =   1
      Top             =   390
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   8308
      _Version        =   393216
      TabOrientation  =   2
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Canvas"
      TabPicture(0)   =   "frmPNGmenu.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtAlphaBlend"
      Tab(0).Control(1)=   "chkAutoErase"
      Tab(0).Control(2)=   "chkAlphaBlend"
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(4)=   "lblBkg(3)"
      Tab(0).Control(5)=   "lblBkg(2)"
      Tab(0).Control(6)=   "lblBkg(1)"
      Tab(0).Control(7)=   "lblBkg(0)"
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "&Size && Pos"
      TabPicture(1)   =   "frmPNGmenu.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(0)"
      Tab(1).Control(1)=   "Label1(1)"
      Tab(1).Control(2)=   "Label3(3)"
      Tab(1).Control(3)=   "txtImgXY(1)"
      Tab(1).Control(4)=   "txtImgXY(0)"
      Tab(1).Control(5)=   "txtScale(1)"
      Tab(1).Control(6)=   "txtScale(0)"
      Tab(1).Control(7)=   "chkScale"
      Tab(1).Control(8)=   "chkRatio"
      Tab(1).Control(9)=   "cmdApply"
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Prg &Display"
      TabPicture(2)   =   "frmPNGmenu.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1(1)"
      Tab(2).Control(1)=   "Frame1(0)"
      Tab(2).Control(2)=   "Label3(2)"
      Tab(2).Control(3)=   "Label3(1)"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "PNG &Info"
      TabPicture(3)   =   "frmPNGmenu.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "lblInfo(0)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "lblInfo(1)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "lblInfo(2)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "lblInfo(3)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Label3(0)"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "lblInfo(4)"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Text1"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "chkShowInfo"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).ControlCount=   8
      Begin VB.TextBox txtAlphaBlend 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -72885
         MaxLength       =   3
         TabIndex        =   4
         Text            =   "255"
         Top             =   4140
         Width           =   840
      End
      Begin VB.CheckBox chkShowInfo 
         Caption         =   "Show this tab after PNG shown"
         Height          =   240
         Left            =   405
         TabIndex        =   45
         Top             =   4380
         Value           =   1  'Checked
         Width           =   2625
      End
      Begin VB.Frame Frame1 
         Height          =   1680
         Index           =   1
         Left            =   -74520
         TabIndex        =   31
         Top             =   2970
         Width           =   2400
         Begin VB.CheckBox chkIL 
            Caption         =   "Non-Interlaced PNGs Only Always use scanner effect"
            Height          =   420
            Left            =   75
            TabIndex        =   35
            ToolTipText     =   "Automode will apply effect if image is > 1/3 screen size"
            Top             =   1155
            Width           =   2235
         End
         Begin VB.OptionButton optIL 
            Caption         =   "Fade-In (non Alpha Only)"
            Height          =   210
            Index           =   2
            Left            =   75
            TabIndex        =   33
            ToolTipText     =   "Note: If Alpha, then Fade In is also the Default"
            Top             =   525
            Width           =   2175
         End
         Begin VB.OptionButton optIL 
            Caption         =   "Don't Progressive Display"
            Height          =   210
            Index           =   1
            Left            =   75
            TabIndex        =   34
            Top             =   855
            Width           =   2175
         End
         Begin VB.OptionButton optIL 
            Caption         =   "Auto Mode"
            Height          =   210
            Index           =   0
            Left            =   75
            TabIndex        =   32
            ToolTipText     =   "Alpha=FadeIn,Non-Alpha=Pixelated,Non-Interlaced=Scanner if >1/3 screen size"
            Top             =   210
            Value           =   -1  'True
            Width           =   2175
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1980
         Index           =   0
         Left            =   -74550
         TabIndex        =   25
         Top             =   540
         Width           =   2520
         Begin VB.OptionButton optTrans 
            Caption         =   "Use the Transparent Color"
            Height          =   210
            Index           =   4
            Left            =   45
            TabIndex        =   28
            Top             =   1095
            Width           =   2400
         End
         Begin VB.OptionButton optTrans 
            Caption         =   "Don't allow any transparency"
            Height          =   210
            Index           =   3
            Left            =   45
            TabIndex        =   30
            Top             =   1710
            Width           =   2400
         End
         Begin VB.OptionButton optTrans 
            Caption         =   "Use another Color"
            Height          =   210
            Index           =   2
            Left            =   45
            TabIndex        =   29
            ToolTipText     =   "Click ellipse to select a color"
            Top             =   1395
            Width           =   1635
         End
         Begin VB.OptionButton optTrans 
            Caption         =   "Use Suggested Window Bkg"
            Height          =   210
            Index           =   1
            Left            =   45
            TabIndex        =   27
            ToolTipText     =   "If no BKG color provided, white will be default"
            Top             =   810
            Width           =   2400
         End
         Begin VB.OptionButton optTrans 
            Caption         =   "Always transparent"
            Height          =   210
            Index           =   0
            Left            =   45
            TabIndex        =   26
            Top             =   210
            Value           =   -1  'True
            Width           =   2235
         End
         Begin VB.Label Label1 
            Caption         =   "Following destroys Alpha info..."
            ForeColor       =   &H000000C0&
            Height          =   225
            Index           =   2
            Left            =   90
            TabIndex        =   46
            Top             =   525
            Width           =   2325
         End
         Begin VB.Label lblShowColorDlg 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            Caption         =   "..."
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1725
            TabIndex        =   36
            ToolTipText     =   "Click to select a color"
            Top             =   1410
            Width           =   420
         End
         Begin VB.Label lblBkgColor 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   1710
            TabIndex        =   41
            ToolTipText     =   "Click to select a color"
            Top             =   1395
            Width           =   690
         End
      End
      Begin VB.TextBox Text1 
         Height          =   2460
         Left            =   450
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   23
         Top             =   1890
         Width           =   2535
      End
      Begin VB.CommandButton cmdApply 
         Caption         =   "Apply Above Settings"
         Height          =   420
         Left            =   -74310
         TabIndex        =   11
         ToolTipText     =   "Select PNG File"
         Top             =   3375
         Width           =   2070
      End
      Begin VB.CheckBox chkRatio 
         Caption         =   "Lock Ratio"
         Height          =   345
         Left            =   -74025
         TabIndex        =   10
         Top             =   2745
         Value           =   1  'Checked
         Width           =   1350
      End
      Begin VB.CheckBox chkScale 
         Caption         =   "Reset to 0,0 coordinates && 100% scale for each newly loaded PNG"
         Height          =   810
         Left            =   -74070
         TabIndex        =   7
         Top             =   1140
         Value           =   1  'Checked
         Width           =   1740
      End
      Begin VB.TextBox txtScale 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   0
         Left            =   -74025
         TabIndex        =   8
         Text            =   "100"
         Top             =   2385
         Width           =   645
      End
      Begin VB.TextBox txtScale 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   1
         Left            =   -73290
         TabIndex        =   9
         Text            =   "100"
         Top             =   2385
         Width           =   645
      End
      Begin VB.TextBox txtImgXY 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   0
         Left            =   -73965
         MaxLength       =   4
         TabIndex        =   5
         Text            =   "0"
         Top             =   675
         Width           =   645
      End
      Begin VB.TextBox txtImgXY 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   1
         Left            =   -73230
         MaxLength       =   4
         TabIndex        =   6
         Text            =   "0"
         Top             =   660
         Width           =   645
      End
      Begin VB.CheckBox chkAutoErase 
         Caption         =   "Always Erase Canvas before Loading new PNG File"
         Height          =   495
         Left            =   -74490
         TabIndex        =   2
         Top             =   810
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CheckBox chkAlphaBlend 
         Caption         =   "Alpha blend PNG over canvas using opacity of"
         Height          =   495
         Left            =   -74490
         TabIndex        =   3
         Top             =   3885
         Width           =   2475
      End
      Begin VB.Label Label3 
         Caption         =   "Tip:  Double click settings to reset to defaults"
         Height          =   420
         Index           =   3
         Left            =   -74325
         TabIndex        =   40
         Top             =   4170
         Width           =   2115
      End
      Begin VB.Label lblInfo 
         Caption         =   "Type: "
         Height          =   240
         Index           =   4
         Left            =   555
         TabIndex        =   39
         Tag             =   "Type: "
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "How is Interlacing handled. Progressive display will be"
         Height          =   435
         Index           =   2
         Left            =   -74490
         TabIndex        =   38
         Top             =   2565
         Width           =   2160
      End
      Begin VB.Label Label3 
         Caption         =   "How are background colors applied for images with transparencies"
         Height          =   465
         Index           =   1
         Left            =   -74565
         TabIndex        =   37
         Top             =   120
         Width           =   2580
      End
      Begin VB.Label Label3 
         Caption         =   "Other Information within PNG"
         Height          =   195
         Index           =   0
         Left            =   495
         TabIndex        =   24
         Top             =   1665
         Width           =   2460
      End
      Begin VB.Label lblInfo 
         Caption         =   "Interlaced: "
         Height          =   240
         Index           =   3
         Left            =   555
         TabIndex        =   22
         Tag             =   "Interlaced: "
         Top             =   1035
         Width           =   2295
      End
      Begin VB.Label lblInfo 
         Caption         =   "Bit Depth:"
         Height          =   240
         Index           =   2
         Left            =   555
         TabIndex        =   21
         Tag             =   "Bit Depth: "
         Top             =   750
         Width           =   2295
      End
      Begin VB.Label lblInfo 
         Caption         =   "Last Modified: "
         Height          =   240
         Index           =   1
         Left            =   555
         TabIndex        =   20
         Tag             =   "Last Modified: "
         Top             =   465
         Width           =   2295
      End
      Begin VB.Label lblInfo 
         Caption         =   "Size: "
         Height          =   240
         Index           =   0
         Left            =   555
         TabIndex        =   19
         Tag             =   "Size: "
         Top             =   195
         Width           =   2295
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "CANVAS CLEARING ACTIONS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -74520
         TabIndex        =   18
         Top             =   405
         Width           =   2475
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Scale W / H %"
         Height          =   195
         Index           =   1
         Left            =   -73890
         TabIndex        =   17
         Top             =   2055
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Image Left / Top"
         Height          =   195
         Index           =   0
         Left            =   -73935
         TabIndex        =   16
         Top             =   330
         Width           =   1200
      End
      Begin VB.Label lblBkg 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Click to Clear the Canvas"
         ForeColor       =   &H000000C0&
         Height          =   300
         Index           =   3
         Left            =   -74565
         TabIndex        =   15
         Top             =   3285
         Width           =   2565
      End
      Begin VB.Label lblBkg 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Click to Remove Bkg Image"
         ForeColor       =   &H000000C0&
         Height          =   300
         Index           =   2
         Left            =   -74565
         TabIndex        =   14
         Top             =   2655
         Width           =   2565
      End
      Begin VB.Label lblBkg 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Click to Change Bkg Image"
         ForeColor       =   &H00C00000&
         Height          =   300
         Index           =   1
         Left            =   -74565
         TabIndex        =   13
         Top             =   2055
         Width           =   2565
      End
      Begin VB.Label lblBkg 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Click to Change Background Color"
         ForeColor       =   &H00C00000&
         Height          =   300
         Index           =   0
         Left            =   -74565
         TabIndex        =   12
         Top             =   1500
         Width           =   2565
      End
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Load PNG Image"
      Height          =   315
      Left            =   60
      TabIndex        =   0
      ToolTipText     =   "Select PNG File"
      Top             =   45
      Width           =   3060
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1155
      Top             =   4470
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblPBar 
      Appearance      =   0  'Flat
      BackColor       =   &H0024CACE&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   60
      TabIndex        =   43
      Tag             =   "3045"
      Top             =   5115
      Visible         =   0   'False
      Width           =   3045
   End
   Begin VB.Label lblInfo 
      Caption         =   "Processing Time:"
      Height          =   225
      Index           =   5
      Left            =   330
      TabIndex        =   44
      Tag             =   "Processing Time: "
      Top             =   5115
      Width           =   2760
   End
End
Attribute VB_Name = "frmPNGmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' For those that know my stuff, I generally add a ton of comments.
' This is for 3 main reasons:
'   1) If this stuff is new to you, hopefully it doubles as a teaching aid
'   2) If you understand it better, then you might be apt to recommend improvements
'   3) So I know what the heck I was thinking when I revisit this in a year or so :)

' Used to close Owner form when owned form is closed
    Private Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Const WM_CLOSE As Long = &H10

' used to create offscreen bitmap
    Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
    Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
    Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
    Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
    Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
    Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
    Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hDC As Long) As Long
    Private Declare Function GetGDIObject Lib "gdi32.dll" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
    Private Type BITMAP
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
    End Type
    Private hOldBmp As Long
    Private hBmp As Long
    Private tDC As Long

' used for simple timing
    Private Declare Function GetTickCount Lib "kernel32.dll" () As Long
    Private compileTipShown As Boolean

' The routines, when compiled are fast, however, when using progressive display
' and also using a custom progress meter, the difference can be >100 ms depending
' on what code you are processing in the class's Progress event. So we'll
' turn off custom progress meter when progressive display is being used.
Private bUseCustomProgressMeter As Boolean

' the class has an option to CRC check each block/chunk of data received from the
' PNG file. Though this is optional, it is recommended. To make the class relatively
' safe from crashing on corrupted PNG data, On Error statements & hard-coded CRC
' checks on critical chunks are used. The following option is for use on
' non-critical chunks that do not prevent the image from being displayed.
' For example, the "last modified timestamp" chunk
Private Const ValidateNonCriticalData As Boolean = True

' the stdPNG class has optional feedback/events
' If feedback and progressive display are not wanted then,
' do not declare using the WithEvents keyword
' i.e.:  Private myPNG As StdPNG

' Following are the 2 Events provided with the class
Private WithEvents myPNG  As StdPNG
Attribute myPNG.VB_VarHelpID = -1

Private Sub myPNG_ProgessiveDisplay(ByVal Width As Long, ByVal Height As Long, _
                    destinationDC As Long, destinationHwnd As Long, _
                    ByVal WinBkgColor As Long, ByVal AlphaPng As Boolean, _
                    X As Long, Y As Long, ByVal IsInterlaced As Boolean, _
                    ByVal useProgression As Boolean)
    ' ^^ Width & Height are the true PNG sizes in pixels
    ' ^^ destinationDC, destinationHwnd must be set to use progressive display
    ' ^^ WinBkgColor is the suggested DC backcolor for the PNG
    ' ^^ AlphaPng is True if the PNG uses any transparencies
    ' ^^ X,Y are set to where progressively displayed PNG should be drawn at
    ' ^^ IsInterlaced is True only if the image is encoded as progressively displayed
    ' ^^ useProgression is a suggestion only. It will be True under the following cases
    '       isInterlaced = True
    '       isInterlaced = False and Width*Height> (ScreenWidth*ScreenHeight)\3
    
    
    ' if you want to use progressive display, you must reply to this event.
    ' Even if the PNG is not encoded with interlacing, simply passing
    ' a valid hDC and hWnd value back to this event, the non-interlaced
    ' PNG will be displayed using a scanner-type progressive display.
    
    ' Progressively displayed images cannot be performed if you want to
    ' simultaneously stretch the image; therefore, there are no options to
    ' specify a blt width/height.
    
    ' This is also the one area where you should check and modify any PNG
    ' properties you want to possibly change (i.e., palette, progressive display mode)
    
    ' NOTE: This event will not be fired if you opted to
    ' set class's ProgressiveDisplay property to pngNeverProgressive
    
    bUseCustomProgressMeter = True
    lblPBar.Visible = True
    
    ' progressive display is not available when stretching image
    ' so we will use our own progress meter
    If Val(txtScale(0)) <> 100 Then Exit Sub
    If Val(txtScale(1)) <> 100 Then Exit Sub
    
    ' basically, don't show a progress meter if we are going to use
    ' progressive display as that, in itself, acts as a progress meter
    If (IsInterlaced = False And chkIL = 1) Or useProgression = True Then
        ' user wants to always display non-interlaced images w/scanner effect
        ' or it is suggested to use progression. Note that if the option to
        ' never use progressive display was set, we wouldn't get this event
        bUseCustomProgressMeter = False
        destinationDC = Form1.hDC
        destinationHwnd = Form1.hwnd
        X = Val(txtImgXY(0))
        Y = Val(txtImgXY(1))
        
        lblPBar.Visible = False
    End If
End Sub
Private Sub myPNG_Progress(ByVal Percentage As Long)

    ' Whether or not you want to display a progress bar is up to you
    ' However, displaying a progress bar while also allowing progressive
    ' display and the image is interlaced, will slow down the routines
    ' just a bit while this event is being called.

    
    ' Percentage is in whole numbers; Percentage/100 for decimals
    If bUseCustomProgressMeter Then
        If Percentage = 100 Then
            lblPBar.Visible = False
            lblPBar.Width = 0
        Else
            lblPBar.Width = Val(lblPBar.Tag) * (Percentage / 100)
        End If
        lblPBar.Refresh
    End If

End Sub


Private Sub cmdApply_Click()

    ' modify last drawn PNG, per user request
    If myPNG Is Nothing Then Exit Sub
    Dim X As Long, Y As Long, cx As Long, cy As Long
    
    X = Val(txtImgXY(0)): Y = Val(txtImgXY(1))
    cx = (Val(txtScale(0)) / 100) * myPNG.Width
    cy = (Val(txtScale(1)) / 100) * myPNG.Height
    
    RefreshCanvas
    myPNG.Paint Form1.hDC, X, Y, cx, cy, , , , myPNG.Width, myPNG.Height
    Form1.Refresh
    
End Sub

Private Sub cmdOpen_Click()

    With CommonDialog1  ' show file select dialog window
        .Flags = cdlOFNFileMustExist Or cdlOFNExplorer
        .CancelError = True
        .Filter = "PNG Files|*.png"
        .FilterIndex = 0
        .DialogTitle = "Select PNG File to Display"
    End With
    On Error GoTo UserAbort
    
    CommonDialog1.ShowOpen
    Me.Refresh
    DoEvents ' refresh our form
    LoadPNGfile CommonDialog1.Filename, CommonDialog1.FileTitle
    If myPNG.Handle = 0 Then
        ' whatever you want to do. Png file did not load
    End If
    
UserAbort:
'If Err Then MsgBox Err.Description
'Resume
End Sub

Private Sub LoadPNGfile(sFileName As String, displayName As String)

    Dim lRtn As Long, dRtn As Double, dDate As Date
    Dim dChromo() As Double, lArray() As Long
    Dim myTimer As Long
    
    SetUserOptions displayName
    
    If myPNG Is Nothing Then Set myPNG = New StdPNG
    
    ' ok, let's start the process
    myTimer = GetTickCount
    If chkAlphaBlend Then lRtn = Val(txtAlphaBlend) Else lRtn = 255
    If myPNG.LoadFile(sFileName, ValidateNonCriticalData, , CByte(lRtn)) Then
    '^^ Parameter info:
    ' 1st :: the full path & filename of the PNG to load/display
    ' 2nd :: True to validate all PNG data, else only critical data is validated
    ' 3rd :: True will force PNG to 32bpp BMP. Needed for alphablending if desired
    ' 4th :: from 0-255 and will blend the final result into a DC when
    '        progressively displaying PNG. Any value < 255 will force a 32bpp BMP
    
        ' Display PNG properties
        myTimer = GetTickCount() - myTimer
        If myTimer < 50 Then '
            ' GetTickcount isn't very accurate at small intervals
            lblInfo(5).Caption = lblInfo(5).Tag & "less than 50 ms"
        Else
            lblInfo(5).Caption = lblInfo(5).Tag & myTimer & " ms"
        End If
    
        lblInfo(0).Caption = lblInfo(0).Tag & myPNG.Width & " x " & myPNG.Height
        
        dRtn = myPNG.LastModified
        
        If dRtn = 0 Then
            lblInfo(1).Caption = lblInfo(1).Tag & "Not provided"
        Else
            lblInfo(1).Caption = lblInfo(1).Tag & Format(dRtn, "Short Date")
        End If
            
        lblInfo(2).Caption = lblInfo(2).Tag & myPNG.BitCount_PNG & ", Converted: " & myPNG.BitCount_BMP
        
        lblInfo(3).Caption = lblInfo(3).Tag & Format(myPNG.IsInterlaced, "Yes/No")
        
        Select Case myPNG.ColorType
        Case 0: ' grayscale
            lblInfo(4).Caption = lblInfo(4).Tag & "Gray Scaled"
        Case 2: ' true color
            lblInfo(4).Caption = lblInfo(4).Tag & "True Color"
        Case 3: ' paletted
            lblInfo(4).Caption = lblInfo(4).Tag & "Paletted"
        Case 4: ' grayscale with transparency
            lblInfo(4).Caption = lblInfo(4).Tag & "Gray Scale w/Alpha"
        Case 6: ' true color with transparency
            lblInfo(4).Caption = lblInfo(4).Tag & "True Color w/Alpha"
        End Select
        
        Text1.Text = myPNG.Comments("No embedded comments")
        dRtn = myPNG.GammaCorrection
        If dRtn Then
            lRtn = Int(1 / dRtn)
            dRtn = Int(1 / dRtn * 10) / 10
            Text1.Text = Text1.Text & vbCrLf & "Gamma Correction: " & FormatNumber(dRtn, 1)
        End If
        
        dChromo() = myPNG.Chromaticity()
        If UBound(dChromo) > -1 Then
            Text1.Text = Text1.Text & vbCrLf & "Chromaticity:" & vbCrLf & "     WhiteX " & FormatNumber(dChromo(0), 2) & _
                vbCrLf & "     WhiteY " & FormatNumber(dChromo(1), 2) & _
                vbCrLf & "     RedX " & FormatNumber(dChromo(2), 2) & _
                vbCrLf & "     RedY " & FormatNumber(dChromo(3), 2) & _
                vbCrLf & "     GreenX " & FormatNumber(dChromo(4), 2) & _
                vbCrLf & "     GreenY " & FormatNumber(dChromo(5), 2) & _
                vbCrLf & "     BlueX " & FormatNumber(dChromo(6), 2) & _
                vbCrLf & "     BlueY " & FormatNumber(dChromo(7), 2)
        End If
        lArray() = myPNG.AspectRatio()
        If UBound(lArray) > -1 Then
            '.0254 is the offical png meter to inch conversion ratio
            If lArray(2) = 1 Then ' pixel size is meters to inches
                Text1.Text = Text1.Text & vbCrLf & "Aspect Ratio:" & _
                    vbCrLf & "  Logical X = " & Int(0.0254 * lArray(0)) & _
                    vbCrLf & "  Logical Y = " & Int(0.0254 * lArray(1))
            Else
                Text1.Text = Text1.Text & vbCrLf & "Aspect Ratio:" & _
                    vbCrLf & "  Aspect Ratio: " & lArray(0) & ":" & lArray(1)
            End If
        End If
        lArray() = myPNG.Offsets()
        If UBound(lArray) > -1 Then
            Text1.Text = Text1.Text & vbCrLf & "Page Offsets:" & _
                vbCrLf & "  X = " & lArray(0) & " " & Choose(lArray(2) + 1, "Pixels", "Microns") & _
                vbCrLf & "  Y = " & lArray(1) & " " & Choose(lArray(2) + 1, "Pixels", "Microns")
        End If
        If myPNG.StdRGB Then
            Text1.Text = Text1.Text & vbCrLf & "Standard RGB color space:" & vbCrLf & _
                "  :: " & Choose(myPNG.StdRGB, "Perceptual", "Relative Colorimetric", "Saturation", "Absolute Colorimetric")
        End If
        
        If bUseCustomProgressMeter = True Then
            ' we were not progressively displaying, so paint the PNG now
            If chkAlphaBlend Then
                myPNG.Paint Form1.hDC, Val(txtImgXY(0)), Val(txtImgXY(1)), _
                    myPNG.Width * (Val(txtScale(0)) / 100), myPNG.Height * (Val(txtScale(1)) / 100), , , , , , Val(txtAlphaBlend)
            Else
                myPNG.Paint Form1.hDC, Val(txtImgXY(0)), Val(txtImgXY(1)), _
                    myPNG.Width * (Val(txtScale(0)) / 100), myPNG.Height * (Val(txtScale(1)) / 100)
            End If
            Form1.Refresh
        End If
        
        If Not compileTipShown Then
            compileTipShown = True
            On Error Resume Next
            Debug.Print 1 / 0
            If Err Then
                MsgBox "Compiling a copy of this application will result in dramatic speed increases", vbInformation + vbOKOnly
            End If
        End If
        
    Else
        ' error occurred... Reset form
        For lRtn = 0 To lblInfo.UBound
            lblInfo(lRtn).Caption = lblInfo(lRtn).Tag
        Next
        Text1 = ""
        
    End If
    
    If chkShowInfo Then SSTab1.Tab = 3 ' show the PNG Info tab
    
End Sub

Private Sub cmdWrite_Click()
    MsgBox "Not yet implemented", vbInformation + vbOKOnly
End Sub

Private Sub Form_Load()
    Me.Move Form1.Left - Me.Width, (Screen.Height - Me.Height) \ 2
    lblPBar.Tag = lblPBar.Width
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' this is a owned form which stays on top of its owner (Form1).
    ' So if I want the owner to close too,
    ' post it a message, otherwise, we go into an infinite loop
    If UnloadMode = 0 Then PostMessage Form1.hwnd, WM_CLOSE, 0, 0
    ' When owner is closing this form closes automatically too. The UnloadMode will = 5
End Sub

Private Sub Form_Unload(Cancel As Integer)

    EraseOffscreenBitmap
    If Not myPNG Is Nothing Then Set myPNG = Nothing

End Sub

Private Sub SetTextBoxFocus(txtObj As TextBox)
    With txtObj
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub lblBkg_Click(Index As Integer)

    EraseOffscreenBitmap
    Select Case Index
    Case 0 ' back color
        With CommonDialog1
            .Flags = cdlCCFullOpen
            .CancelError = True
        End With
        On Error GoTo UserAbort
        CommonDialog1.ShowColor
        Form1.BackColor = CommonDialog1.Color
        
    Case 1 ' back image
        With CommonDialog1
            .Flags = cdlOFNFileMustExist
            .Filter = "Bitmaps|*.bmp|JPeg|*.jpg;*.jpeg|GIFs|*.gif"
            .FilterIndex = 0
            .Filename = ""
            .DialogTitle = "Select Canvas Background Image"
            .CancelError = True
        End With
        On Error GoTo UserAbort
        CommonDialog1.ShowOpen
        Set Form1.Picture = LoadPicture(CommonDialog1.Filename)
        
    Case 2 ' remove back image
        Set Form1.Picture = LoadPicture("")
        
    Case 3 ' clear
        Form1.Cls
    End Select

    If Not myPNG Is Nothing Then
        If Index < 3 Then Call cmdApply_Click
    End If

UserAbort:
End Sub

Private Sub lblBkgColor_Click()
    Call lblShowColorDlg_Click
End Sub

Private Sub lblShowColorDlg_Click()
    With CommonDialog1
        .Flags = cdlCCFullOpen
        .CancelError = True
    End With
    On Error GoTo UserAbort
    CommonDialog1.ShowColor
    lblBkgColor.BackColor = CommonDialog1.Color
    
UserAbort:
End Sub

Private Sub optIL_Click(Index As Integer)
    If myPNG Is Nothing Then Set myPNG = New StdPNG
    Select Case Index
    Case 0: ' Default
        myPNG.ProgressiveDisplay = pngAuto
    Case 1: ' Never use it
        myPNG.ProgressiveDisplay = pngNeverProgressive
        ' can't have scanner-effect for non-interlaced either then
        chkIL = 0
    Case 2: ' Fade-In for non-transparent interlaced images
        myPNG.ProgressiveDisplay = pngFadeIn
    End Select
    If chkIL = 1 Then myPNG.ProgressiveDisplay = myPNG.ProgressiveDisplay Or pngScanner
End Sub

Private Sub optTrans_Click(Index As Integer)
    
' Things to think about when replacing transparency with solid/opaque colors:
' 1. The converted DIB will be smaller
' 2. All transparency information is permanently lost within the DIB; the PNG
'    is untouched. The loss is due to a DIB bit depth < 32bpp. If you wish to
'    keep transparency but want a different bkg color, load the PNG using
'    full transparency, but don't display it progressively. Next fill your
'    DC with the bkg color and simply stdPNG.Paint the DIB over that DC.
'    The .Paint method accepts an alphablend value so you can blend it to a bkg too
' 3. Any 32bpp DIB cannot be saved 100% lossless to a PNG format. This is for 2
'    main reasons. First: 32bpp images have premultiplied RGB values and calculating
'    the non-premultiplied values can be off by 1 due to rounding errors. And the
'    2nd reason: Pre-multiplied full transparent pixels are RGB(0,0,0) regardless
'    of its original color; therefore, getting the original transparent color back
'    is impossible. This would generally not be too big of an issue, but
'    if you want that color back, the DIB would have to be processed pixel by pixel
'    and the color would have to manually added.
    
    If myPNG Is Nothing Then Set myPNG = New StdPNG
    
    Select Case Index
        Case 0: myPNG.TransparentStyle = alphaTransparent
                '^^ alpha is alpha
        Case 1: myPNG.TransparentStyle = alphaPNGwindowBkg  ' use default bkg
                '^^ if not provided by png, it will be white
                '   Maybe best option for color types 4,6
        Case 2: myPNG.TransparentStyle = lblBkgColor.BackColor
                '^^ supply own bkg
        Case 3: myPNG.TransparentStyle = alphaNoBkgNoAlpha
                '^^ worse possible choice, shows alpha as non-alpha & those colors
                '   can be any color in the world. The image may really look bad IMO
        Case 4: myPNG.TransparentStyle = alphaTransColorBkg
        '^^ this is usually the best choice for converting alpha to non-alpha
        '   When the PNG color type is 0,2,3. Color types 4,6 have embeeded
        '   alpha values and those generally look better by selecting the
        '   suggest window bkg color or providing your own if no suggested color provided
    End Select
End Sub

Private Sub txtAlphaBlend_GotFocus()
    SetTextBoxFocus txtAlphaBlend
End Sub

Private Sub txtAlphaBlend_Validate(Cancel As Boolean)
    Select Case Val(txtAlphaBlend)
        Case Is < 0: txtAlphaBlend = 0
        Case Is > 255: txtAlphaBlend = 255
        Case Else: txtAlphaBlend = Int(txtAlphaBlend)
    End Select
End Sub

Private Sub txtImgXY_DblClick(Index As Integer)
    ' just reset the X,Y coordinates to 0,0 when double clicked
    txtImgXY(Index) = 0
    If chkRatio Then txtImgXY(Abs(Index - 1)) = 0
End Sub

Private Sub txtImgXY_GotFocus(Index As Integer)
    SetTextBoxFocus txtImgXY(Index)
End Sub

Private Sub txtImgXY_Validate(Index As Integer, Cancel As Boolean)

    Dim maxCX As Long, maxCY As Long
    maxCX = Screen.Width \ Screen.TwipsPerPixelX
    maxCY = Screen.Height \ Screen.TwipsPerPixelY
    
    If Index = 0 Then
        If Val(txtImgXY(Index)) > maxCX Then
            txtImgXY(Index) = maxCX
        ElseIf Val(txtImgXY(Index)) < -maxCX Then
            txtImgXY(Index) = -maxCX
        End If
    Else
        If Val(txtImgXY(Index)) > maxCY Then
            txtImgXY(Index) = maxCY
        ElseIf Val(txtImgXY(Index)) < -maxCY Then
            txtImgXY(Index) = -maxCY
        End If
    End If
End Sub

Private Sub txtScale_Change(Index As Integer)
    If chkRatio Then txtScale(Abs(Index - 1)).Text = txtScale(Index)
End Sub

Private Sub txtScale_DblClick(Index As Integer)
    ' reset the scaling to 100% when double clicked
    txtScale(Index) = 100
    If chkRatio Then txtScale(Abs(Index - 1)) = 100

End Sub

Private Sub txtScale_GotFocus(Index As Integer)
    SetTextBoxFocus txtScale(Index)
End Sub

Private Sub txtScale_Validate(Index As Integer, Cancel As Boolean)
    If Index = 0 Then
        If Val(txtScale(Index)) > 500 Then txtScale(Index) = 500
    Else
        If Val(txtScale(Index)) > 500 Then txtScale(Index) = 500
    End If
    If Val(txtScale(Index)) < 10 Then txtScale(Index) = 10
End Sub

Private Sub chkIL_Click()
    If chkIL = 1 Then
        ' if wanting non-interlaced to have the interlaced scanner effect,
        ' then ensure the "Prevent Interlacing" option is not active
        If optIL(1) = True Then optIL(0) = True
    End If
End Sub

Private Sub RefreshCanvas()
    Form1.Cls
    If tDC Then
        Dim bmpInfo As BITMAP
        GetGDIObject hBmp, Len(bmpInfo), bmpInfo
        BitBlt Form1.hDC, 0, 0, bmpInfo.bmWidth, bmpInfo.bmHeight, tDC, 0, 0, vbSrcCopy
    End If
    Form1.Refresh
End Sub

Private Sub CreateScreenShot()

    Dim Wd As Long, Ht As Long
    
    If tDC = 0 Then
        tDC = CreateCompatibleDC(Me.hDC)
    Else
        DeleteObject SelectObject(tDC, hOldBmp)
    End If
    
    Wd = Form1.ScaleWidth \ Screen.TwipsPerPixelX
    Ht = Form1.ScaleHeight \ Screen.TwipsPerPixelY
    
    hBmp = CreateCompatibleBitmap(Me.hDC, Wd, Ht)
    ReleaseDC Me.hwnd, Me.hDC
    
    hOldBmp = SelectObject(tDC, hBmp)
    BitBlt tDC, 0, 0, Wd, Ht, Form1.hDC, 0, 0, vbSrcCopy
    
End Sub

Private Sub EraseOffscreenBitmap()

    If tDC Then
        DeleteObject SelectObject(tDC, hOldBmp)
        DeleteDC tDC
    End If
    tDC = 0

End Sub

Private Sub SetUserOptions(displayName As String)

    ' following is to set up the form & options based on the different
    ' option buttons/checkboxes, etc, that you may have been clicking on
    Form1.Caption = displayName

    If chkAutoErase = 0 Then
        CreateScreenShot
    Else
        EraseOffscreenBitmap
        Form1.Cls
        DoEvents
    End If
    If chkScale = 1 Then
        txtImgXY(0) = 0
        txtImgXY(1) = 0
        txtScale(0) = 100
        txtScale(1) = 100
    End If
    
    ' ensure any settings while PNG not active are accounted for
    Dim I As Integer
    For I = 0 To optIL.UBound
        If optIL(I) = True Then
            Call optIL_Click(I)
            Exit For
        End If
    Next
    For I = 0 To optTrans.UBound
        If optTrans(I) = True Then
            Call optTrans_Click(I)
            Exit For
        End If
    Next
    If optIL(1) = True Then ' prevent progressive display, use custom progress meter
        bUseCustomProgressMeter = True
        lblPBar.Visible = True
    End If

End Sub
