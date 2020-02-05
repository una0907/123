VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  '單線固定
   Caption         =   "DirectShow Webcam Minimal Example"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   15270
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   15270
   StartUpPosition =   2  '螢幕中央
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   375
      Left            =   7320
      TabIndex        =   43
      Top             =   9120
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   9360
      TabIndex        =   42
      Top             =   9000
      Width           =   1335
   End
   Begin VB.CommandButton CHECK 
      Caption         =   "START"
      Height          =   495
      Left            =   10560
      TabIndex        =   41
      Top             =   6360
      Width           =   1245
   End
   Begin VB.CommandButton ALIGN 
      Caption         =   "Standardize"
      Height          =   495
      Left            =   11880
      TabIndex        =   37
      Top             =   6360
      Width           =   1455
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13320
      TabIndex        =   36
      Top             =   5640
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9120
      TabIndex        =   35
      Text            =   "127"
      Top             =   7440
      Width           =   1095
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   34
      Text            =   "51"
      Top             =   7440
      Width           =   975
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9120
      TabIndex        =   33
      Text            =   "750"
      Top             =   6720
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   32
      Text            =   "0"
      Top             =   6720
      Width           =   975
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   4920
      Top             =   7320
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13320
      TabIndex        =   26
      Top             =   4800
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13320
      TabIndex        =   25
      Top             =   4200
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      HideSelection   =   0   'False
      Left            =   13320
      TabIndex        =   24
      Top             =   7920
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13320
      TabIndex        =   23
      Top             =   7200
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "START"
      Height          =   495
      Left            =   10560
      TabIndex        =   22
      Top             =   7320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "255"
      Height          =   495
      Index           =   10
      Left            =   8760
      TabIndex        =   21
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "0"
      Height          =   495
      Index           =   9
      Left            =   11640
      TabIndex        =   20
      Top             =   5400
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "100"
      Height          =   495
      Index           =   8
      Left            =   8760
      TabIndex        =   19
      Top             =   5400
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "75"
      Height          =   495
      Index           =   7
      Left            =   9720
      TabIndex        =   18
      Top             =   5400
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "50"
      Height          =   495
      Index           =   6
      Left            =   10680
      TabIndex        =   17
      Top             =   5400
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "125"
      Height          =   495
      Index           =   5
      Left            =   10680
      TabIndex        =   16
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "150"
      Height          =   495
      Index           =   4
      Left            =   9720
      TabIndex        =   15
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "175"
      Height          =   495
      Index           =   3
      Left            =   8760
      TabIndex        =   14
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "200"
      Height          =   495
      Index           =   2
      Left            =   11640
      TabIndex        =   13
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "225"
      Height          =   495
      Index           =   1
      Left            =   10680
      TabIndex        =   12
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "250"
      Height          =   495
      Index           =   0
      Left            =   9720
      TabIndex        =   11
      Top             =   4200
      Width           =   855
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Left            =   6120
      Max             =   255
      TabIndex        =   8
      Top             =   4680
      Value           =   255
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   7560
      TabIndex        =   7
      Text            =   "150"
      Top             =   4200
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '平面
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3600
      Left            =   9960
      ScaleHeight     =   238
      ScaleMode       =   3  '像素
      ScaleWidth      =   320
      TabIndex        =   6
      Top             =   360
      Width           =   4830
      Begin VB.Label Label11 
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   1200
         TabIndex        =   44
         Top             =   3120
         Width           =   2655
      End
      Begin VB.Label ReadValue 
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   3840
         TabIndex        =   29
         Top             =   3120
         Width           =   855
      End
   End
   Begin VB.ListBox lstFilters 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2580
      ItemData        =   "Form1.frx":1B1A
      Left            =   120
      List            =   "Form1.frx":1B1C
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   4440
      Width           =   4755
   End
   Begin VB.CommandButton cmdSnap 
      Caption         =   "Binarized image"
      Height          =   495
      Left            =   6120
      TabIndex        =   2
      Top             =   5160
      Width           =   2295
   End
   Begin VB.PictureBox picSnapshot 
      Appearance      =   0  '平面
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3600
      Left            =   5040
      ScaleHeight     =   238
      ScaleMode       =   3  '像素
      ScaleWidth      =   320
      TabIndex        =   3
      Top             =   360
      Width           =   4830
   End
   Begin VB.Label Label10 
      Caption         =   "平均"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12720
      TabIndex        =   40
      Top             =   5760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label LABEL9 
      Alignment       =   2  '置中對齊
      Caption         =   "Y"
      Height          =   495
      Index           =   1
      Left            =   12720
      TabIndex        =   39
      Top             =   7920
      Width           =   495
   End
   Begin VB.Label LABEL9 
      Alignment       =   2  '置中對齊
      Caption         =   "X"
      Height          =   495
      Index           =   0
      Left            =   12720
      TabIndex        =   38
      Top             =   7200
      Width           =   495
   End
   Begin VB.Label Label8 
      Alignment       =   1  '靠右對齊
      Caption         =   "Scale        min<--->max"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      TabIndex        =   31
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "   Range"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   30
      Top             =   7440
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "讀值"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   12720
      TabIndex        =   28
      Top             =   4920
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "對應角度"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   12720
      TabIndex        =   27
      Top             =   4200
      Width           =   495
   End
   Begin VB.Image imgPlaceHolder 
      Appearance      =   0  '平面
      BorderStyle     =   1  '單線固定
      Height          =   3600
      Left            =   120
      Stretch         =   -1  'True
      Top             =   360
      Width           =   4830
   End
   Begin VB.Label Label5 
      Caption         =   "Time"
      Height          =   495
      Left            =   6120
      TabIndex        =   10
      Top             =   6000
      Width           =   4695
   End
   Begin VB.Label Label4 
      Caption         =   "Grayscale threshold"
      Height          =   375
      Left            =   5400
      TabIndex        =   9
      Top             =   4200
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   8040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  '置中對齊
      Caption         =   "Snapshot"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   1
      Top             =   60
      Width           =   4830
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      Caption         =   "Preview"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4830
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Pixel
Dim Pixel2

Dim Rred
Dim Ggreen
Dim Bblue

Dim RR1
Dim GG1
Dim BB1

Dim RR2
Dim GG2
Dim BB2

Dim RR3
Dim GG3
Dim BB3

Dim Q As String
Dim Q2 As String

Dim Temp As Integer
Dim Temp2 As Integer

Dim XXX As Integer
Dim YYY As Integer

Dim XX As Integer   '變數 控制讀值範圍的原點
Dim YY As Integer   '左右 控制讀值範圍的原點


Dim II As Integer
'Dim JJ As Long
Dim JJ As Single


Dim RR As Integer
Dim RG As Integer
Dim RB As Integer
Dim TH As Integer

Dim coordrry(500, 500) As Integer  '矩陣

Dim CurX
Dim CurY

Dim JB As Byte

' For XP manifest
Private Declare Sub InitCommonControls Lib "comctl32" ()
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" ( _
    ByVal lpLibFileName As String) As Long
    
Private Declare Sub Sleep Lib "kernel32" Alias "sleep" (ByVal dwMilliseconds As Long)

Private Declare Function FreeLibrary Lib "kernel32" ( _
   ByVal hLibModule As Long) As Long
Private m_hMod As Long
''

Private FWidORG As Long          ' Form starting & minimum W & H
Private FHitORG As Long
Private RightMargin As Long, BottomMargin As Long  ' Hard twip values

Private PathSpec$
Private CurrentPath$, FileSpec$
Private SavePath$, SaveSpec$

Private aScroll As Boolean       ' Booleans to restrict scroll bar actions
Private aZoom As Boolean
Private aRGB As Boolean
Private aMouseDown As Boolean

Private aEffAction As Boolean    ' T Continuous, F Stepped

Private aSelect As Boolean



' 以上新增



Public Oked As Boolean
Public CameraName As String
Public UI As Integer
'Requires a reference to:
'
'   ActiveMovie control type library (quartz.dll).
'

Private Const WS_BORDER = &H800000
Private Const WS_DLGFRAME = &H400000
Private Const WS_SYSMENU = &H80000
Private Const WS_THICKFRAME = &H40000
Private Const MASKBORDERLESS = Not (WS_BORDER Or WS_DLGFRAME Or WS_SYSMENU Or WS_THICKFRAME)
Private Const MASKBORDERMIN = Not (WS_DLGFRAME Or WS_SYSMENU Or WS_THICKFRAME)

'FILTER_STATE values, should have been defined in Quartz.dll,
'but another item Microsoft left out.
Private Enum FILTER_STATE
    State_Stopped = 0
    State_Paused = 1
    State_Running = 2
End Enum

Private Const E_FAIL As Long = &H80004005


'These are "scripts" followed by BuildGraph() below to create a
'DirectShow FilterGraph for webcam viewing.
'
'FILTERLIST is incomplete, and must be prepended with the name
'of your webcam's Video Capture Source filter.  Since there may
'be multiples, FILTERLIST begins with "~Capture" which is used
'when BuildGraph() interprets this script to select one having
'a pin named "Capture".
Private Const FILTERLIST As String = _
        "~Capture|" _
      & "AVI Decompressor|" _
      & "Color Space Converter|" _
      & "Video Renderer"
Private Const CONNECTIONLIST As String = _
        "Capture~XForm In|" _
      & "XForm Out~Input|" _
      & "XForm Out~VMR Input0"

Private fgmVidCap As QuartzTypeLib.FilgraphManager 'Not "Is Nothing" means camera is previewing.
Private bv2VidCap As QuartzTypeLib.IBasicVideo2
Private vwVidCap As QuartzTypeLib.IVideoWindow
Private SelectedCamera As Integer '-1 means none selected.
Private InsideWidth As Double
Private AspectRatio As Double
'Dim data(7) As String                      'XML
'Dim dataname(7) As String                   'XML
Dim data(5) As String                      'XML   上傳10個
Dim dataname(5) As String                   'XML

Private Function BuildGraph( _
    ByVal FGM As QuartzTypeLib.FilgraphManager, _
    ByVal Filters As String, _
    ByVal Connections As String) As Integer
    'Returns -1 on success, or FilterIndex when not found, or
    'ConnIndex + 100 when a pin of the connection not found.
    '
    'Filters:
    '
    '   A string with Filter Name values separated by "|" delimiters
    '   and optionally each of these can be followed by one required
    '   Pin Name value separated by a "~" delimiter for use as a tie
    '   breaker when there might be multiple filters with the same
    '   Name value.
    '
    'Connections:
    '
    '   A string with a list of output pins to be connected to
    '   input pins.  Each pin-pair is separated by "|" delimiters
    '   and each pair has out and in pins separated by a "~"
    '   delimiter.  The pin-pairs should be one less than the number
    '   of filters.
    Dim FilterNames() As String
    Dim FilterIndex As Integer
    Dim FilterParts() As String
    Dim FoundFilter As Boolean
    Dim rfiEach As QuartzTypeLib.IRegFilterInfo
    Dim fiFilters() As QuartzTypeLib.IFilterInfo
    Dim Conns() As String
    Dim ConnIndex As Integer
    Dim ConnParts() As String
    Dim piEach As QuartzTypeLib.IPinInfo
    Dim piOut As QuartzTypeLib.IPinInfo
    Dim piIn As QuartzTypeLib.IPinInfo
    
    'Setup for filter script processing.
    FilterNames = Split(UCase$(Filters), "|")
    ReDim fiFilters(UBound(FilterNames))
    
    'Find and add filters.
    For FilterIndex = 0 To UBound(FilterNames)
        FilterParts = Split(FilterNames(FilterIndex), "~")
        For Each rfiEach In FGM.RegFilterCollection
            If UCase$(rfiEach.Name) = FilterParts(0) Then
                rfiEach.Filter fiFilters(FilterIndex)
                If UBound(FilterParts) > 0 Then
                    For Each piEach In fiFilters(FilterIndex).Pins
                        If UCase$(piEach.Name) = FilterParts(1) Then
                            FoundFilter = True
                            Exit For
                        End If
                    Next
                Else
                    FoundFilter = True
                    Exit For
                End If
            End If
        Next
        If FoundFilter Then
            FoundFilter = False
        Else
            BuildGraph = FilterIndex
            Exit Function 'Error result will be 0, 1, etc.
        End If
    Next
    BuildGraph = -1
    
    'Setup for connection script processing.
    Conns = Split(UCase$(Connections), "|")
    FilterIndex = 0
    
    'Find and connect pins.
    For ConnIndex = 0 To UBound(Conns)
        ConnParts = Split(Conns(ConnIndex), "~")
        For Each piEach In fiFilters(FilterIndex).Pins
            If UCase$(piEach.Name) = ConnParts(0) Then
                Set piOut = piEach
                Exit For
            End If
        Next
        For Each piEach In fiFilters(FilterIndex + 1).Pins
            If UCase$(piEach.Name) = ConnParts(1) Then
                Set piIn = piEach
                Exit For
            End If
        Next
        If piOut Is Nothing Or piIn Is Nothing Then
            'Error, missing a pin.
            BuildGraph = ConnIndex + 100 'Error result will be 100, 101, etc.
            Exit Function
        End If
        piOut.ConnectDirect piIn
        FilterIndex = FilterIndex + 1
    Next
End Function



Private Function StartCamera(ByVal CamName As String) As Integer
    'Returns -1 on success, or BuildGraph() error on failures.
    
    Set fgmVidCap = New QuartzTypeLib.FilgraphManager
    'Tack camera name onto FILTERLIST and try to start it.
    StartCamera = BuildGraph(fgmVidCap, CamName & FILTERLIST, CONNECTIONLIST)
    If StartCamera >= 0 Then Exit Function
    
    Set bv2VidCap = fgmVidCap
    With bv2VidCap
        AspectRatio = CDbl(.VideoHeight) / CDbl(.VideoWidth)
    End With
    
    Set vwVidCap = fgmVidCap
    With vwVidCap
        .FullScreenMode = False
        .Left = ScaleX(imgPlaceHolder.Left, ScaleMode, vbPixels)
        .Top = ScaleY(imgPlaceHolder.Top, ScaleMode, vbPixels)
        .Width = ScaleX(InsideWidth, ScaleMode, vbPixels) + 2
        .Height = ScaleY(InsideWidth * AspectRatio, ScaleMode, vbPixels) + 2
        picSnapshot.Height = InsideWidth * AspectRatio + ScaleY(2, vbPixels, ScaleMode)
        imgPlaceHolder.Visible = False
        .WindowStyle = .WindowStyle And MASKBORDERMIN
        .Owner = hWnd
        .Visible = True
    End With
    
    StartCamera = -1
    cmdSnap.Enabled = True
    fgmVidCap.Run
End Function



Private Function StopCamera(ByVal CamName As String) As Integer
'Private Sub StopCamera(ByVal CamName As String) As Integer
'Private Sub StopCamera()
    Const StopWaitMs As Long = 40
    Dim State As FILTER_STATE
    
    If Not fgmVidCap Is Nothing Then
        With fgmVidCap
            .Stop
            Do
                .GetState StopWaitMs, State
            Loop Until State = State_Stopped Or Err.Number = E_FAIL
        End With
        With vwVidCap
            .Visible = False
            .Owner = 0
        End With
        Set vwVidCap = Nothing
        Set bv2VidCap = Nothing
        Set fgmVidCap = Nothing
    End If
    imgPlaceHolder.Visible = True
    cmdSnap.Enabled = False
End Function

Private Sub ALIGN_Click()  '校正用
Form1.Timer3.Enabled = False
Call StopCamera(lstFilters.Text)
End Sub

Private Sub cmdSnap_Click()
Dim ST, ET As Date

ST = Now()

    Const PauseWaitMs As Long = 16
    Const biSize = 40 'BITMAPINFOHEADER and not BITMAPV4HEADER, etc. but we don't get those.
    Dim State As FILTER_STATE
    Dim PI As Long
    Dim p As Single
    Dim e As Single
    Dim Size As Long
    Dim DIB() As Long
    Dim hBitmap As Long
    Dim PIC As StdPicture
    Dim M, D, H, MI, s As String
    Dim CYCLE As Integer
    Dim CYCLEARRAY(10) As Integer
    Dim GG As Integer
    GG = 0
    Dim START_PIC As Integer
    Dim STOP_PIC As Integer
    
    'For CYCLE = 0 To 2

    With fgmVidCap
        .Pause
        Do
            .GetState PauseWaitMs, State
        Loop Until State = State_Paused Or Err.Number = E_FAIL
        If Err.Number = E_FAIL Then
            MsgBox "Failed to pause webcam preview for snapshot!", _
                   vbOKOnly Or vbExclamation
            Exit Sub
        End If
    
        With bv2VidCap
            'Estimate size.  Correct for 32-bit RGB and generous
            'for anything with fewer bits per pixel, compressed,
            'or palette-ized (we hope).
            Size = biSize + .VideoWidth * .VideoHeight
            ReDim DIB(Size - 1)
            Size = Size * 4 'To bytes.
            .GetCurrentImage Size, DIB(0)
        End With
        
        .Run
    End With
    
    hBitmap = LongDIB2HBitmap(DIB)
    If hBitmap <> 0 Then
        Set PIC = HBitmap2Picture(hBitmap, 0)
        If Not PIC Is Nothing Then
            With picSnapshot
                .AutoRedraw = True
                .PaintPicture PIC, 0, 0, .ScaleWidth, .ScaleHeight
                .AutoRedraw = False
            End With
        End If
        DeleteObject hBitmap
    End If
    
On Error Resume Next




M = Right("0" & Month(Now()), 2)
D = Right("0" & Day(Now()), 2)
H = Right("0" & Hour(Now()), 2)
MI = Right("0" & Minute(Now()), 2)
s = Right("0" & Second(Now()), 2)

'日點檢

START_PIC = 25
STOP_PIC = 55

If Hour(Now()) = 11 And (START_PIC <= CInt(MI) And CInt(MI) < STOP_PIC) Then

'SavePicture picSnapshot.Picture, "C:\PIC\" & lstFilters.Text & "_" & M & D & "_" & H & MI & s & ".bmp"

End If

'日點檢


'二值化開始

picSnapshot.Picture = picSnapshot.Image
Picture1.Height = picSnapshot.Height
Picture1.Picture = picSnapshot.Picture

 
   
    Text1.Text = GetPrivateProfileInt(lstFilters.Text, "TH", 0, App.Path & "\config.ini ")
    Text2.Text = GetPrivateProfileInt(lstFilters.Text, "X", 0, App.Path & "\config.ini ")
    Text4.Text = GetPrivateProfileInt(lstFilters.Text, "Y", 0, App.Path & "\config.ini ")
    Text6.Text = GetPrivateProfileInt(lstFilters.Text, "V1", 0, App.Path & "\config.ini ")
    Text7.Text = GetPrivateProfileInt(lstFilters.Text, "V2", 0, App.Path & "\config.ini ")
    Text8.Text = GetPrivateProfileInt(lstFilters.Text, "A1", 0, App.Path & "\config.ini ")
    Text9.Text = GetPrivateProfileInt(lstFilters.Text, "A2", 0, App.Path & "\config.ini ")


TH = Text1.Text
   
On Error Resume Next

For YYY = 0 To Picture1.ScaleHeight - 1
For XXX = 0 To Picture1.ScaleWidth - 1

Pixel = GetPixel(Picture1.HDC, XXX, YYY)  'API介紹

'ReDim Preserve Coordrry(XXX, YYY)

GetRGB Pixel

Temp = (Rred + Ggreen + Bblue)
Temp = (Temp / 3)

If Val(Temp) >= TH Then

Pixel = vbWhite
coordrry(XXX, YYY) = 0   '矩陣

Else

Pixel = vbBlack
coordrry(XXX, YYY) = 1   '矩陣

End If




SetPixelV Picture1.HDC, XXX, YYY, Pixel
Next
Picture1.Refresh
Next
Picture1.Refresh

XX = CInt(Text2.Text)
YY = CInt(Text4.Text)

CYCLEARRAY(CYCLE) = move_image(XX, YY)

GG = GG + CYCLEARRAY(CYCLE) / 5

Text10.Text = GG

Debug.Print CYCLEARRAY(CYCLE)



ET = Now()

Label5.Caption = "處理時間： " & Format((ET - ST) * 86400, "0.00000000000") & " 秒"

'日點檢

START_PIC = 0   '起始分鐘
STOP_PIC = 5    '終止分鐘

'If Hour(Now()) = 11 And (START_PIC <= CInt(MI) And CInt(MI) < STOP_PIC) Then

'SavePicture Picture1.Image, "F:\CF_INT\PIC\" & lstFilters.Text & "_" & M & D & "_" & H & MI & s & "_VALUE_" & Text5.Text & "_TH_" & TH & ".bmp"
 SavePicture Picture1.Image, "C:\PIC\" & lstFilters.Text & "_" & M & D & "_" & H & MI & s & "_VALUE_" & Text5.Text & "_TH_" & TH & ".bmp"

'End If

'日點檢
   
Call StopCamera(lstFilters.Text)


Delay 2
'Next CYCLE

End Sub


Private Function move_image(XX, YY) As Integer
'畫線
    Dim PI As Single
    Dim p As Single
    Dim e As Single
    
    Dim Read_sec As Integer
    Dim ck As Integer
    Dim i As Long
    Dim A1 As Integer   '變數 控制讀值範圍的最小值
    Dim A2 As Integer   '控制讀值範圍的最大值
    Dim V1 As Integer   '變數 控制刻度範圍的最小值
    Dim V2 As Integer   '控制刻度範圍的最大值

PI = 4 * Atn(1)
'p = 51 * pi / 180   'start
'e = 131 * pi / 180  'stop
V1 = Int(Trim(Text6.Text))     'V1由UI輸入
V2 = Int(Trim(Text7.Text))     'V2由UI輸入
A1 = Int(Trim(Text8.Text))     'A1由UI輸入
A2 = Int(Trim(Text9.Text))     'A2由UI輸入


Dim X1, Y1 As Double
Dim X2, Y2 As Double
Dim redss As Long
Dim redss2 As Long
Dim CC As Integer
Dim MF, C1 As Integer

MF = ""
C1 = 0
     
    ' For JJ = 51 To 127 Step 0.1
      For JJ = A1 To A2 Step 0.1
      
CC = 0  '黑點顆數
    For II = 160 To 170 Step 1
    


    X1 = Round(Cos(JJ / 360 * 2 * PI) * II, 0)
    
    Y1 = Round(Sin(JJ / 360 * 2 * PI) * II, 0)
    'Debug.Print X1, Y1
    
    If (coordrry(153 - X1 + XX, 208 - Y1 + YY) = 1) Then CC = CC + 1 '???   '當取得為黑色時, 自動+1   原點
    
    Picture1.Circle (153 - X1 + XX, 208 - Y1 + YY), 1, vbRed '自動偵測
 
    Next II


Debug.Print JJ, CC

If (CC > C1) Then
    C1 = CC                                   'C1是會佔存不受上續C1=0影響
    
    
    
    MF = JJ    '對應角度
    
   
    
    Dim ff As Integer
    
    ff = Int((MF - A1) * V2 / (A2 - A1 + 1))
 
    
    Select Case ff          '補offset
        
       ' Case 300 To 445
            
          '  ff = ff - 4
'
'        Case 455 To 579
'
'            ff = ff + 4
'
'         Case 580 To 679
'
'           ff = ff + 8
           
         Case 680 To 750
        
           ff = ff + 18
           
           If ff > 750 Then ff = 750     '補offset
      
    End Select
    
    Label11.Caption = lstFilters.Text
    
    ReadValue.Caption = ff
    
    move_image = ff
    
    Text3.Text = MF
    
    Text5.Text = ff
    
    Select Case lstFilters.ListIndex                                                       'XML
 
 
    Case 0
    data(1) = ff
    dataname(1) = "CNVR06_PRCN_P1"
 
    Case 1
    data(2) = ff
    dataname(2) = "CNVR06_PRCN_P2"
    
'    Case 2
'    data(3) = ff
'    dataname(3) = "CNVR06_PRCN_P3"
'
'    Case 3
'    data(4) = ff
'    dataname(4) = "CNVR06_PRCN_A1"
'
'    Case 4
'    data(5) = ff
'    dataname(5) = "CNVR06_PRCN_A2"
'
'    Case 5
'    data(6) = ff
'    dataname(6) = "CNVR06_PRCN_D1"
'
'    Case 6
'    data(7) = ff
'    dataname(7) = "CNVR06_PRCN_D2"
'
'    Case 7
'    data(8) = ff
'    dataname(8) = "CNVR06_PRCN_D3"
'
'    Case 8
'    data(9) = ff
'    dataname(9) = "CNVR06_PRCN_D4"
'
'    Case 9
'    data(10) = ff
'    dataname(10) = "CNVR06_PRCN_D5"
'
'
    
    
    Call Write_xml
    
    End Select
   
End If

    Next JJ
    
If C1 = 0 Then
    ReadValue.Caption = "NA"
   
'    MsgBox "指針超出範圍", vbOKOnly
    
    
End If

Picture1.Circle (65 + XX, 127 + YY), 7, vbBlue '藍色小點定位1
Picture1.Circle (237 + XX, 127 + YY), 7, vbBlue '藍色小點定位2

Picture1.Circle (153 + XX, 208 + YY), 7, vbBlue '藍色小點定位2


'Picture1.Circle (151 + XX, 124 + YY), 115, vbRed '外框定位1
'Picture1.Circle (151 + XX, 124 + YY), 130, vbRed '外框定位2
'Picture1.Line (0, 160)-(340, 160), vbRed, BF '下部基準線


End Function

Private Sub GetRGB(ByVal Col As String)
On Error Resume Next
    Bblue = Col \ (256 ^ 2)
    Ggreen = (Col - Bblue * 256 ^ 2) \ 256
    Rred = (Col - Bblue * 256 ^ 2 - Ggreen * 256) '\ 256
End Sub


Private Sub CHECK_Click()

Dim StartResult As Integer
Dim rfiEach As QuartzTypeLib.IRegFilterInfo
Dim i, j As Integer


    InsideWidth = picSnapshot.Width - ScaleX(2, vbPixels, ScaleMode)
    
 
    lstFilters.Clear
    With New QuartzTypeLib.FilgraphManager
        For Each rfiEach In .RegFilterCollection
        
            '只列舉有需要的名稱
            If (InStr(UCase(rfiEach.Name), "CAMERA")) Then
            
            '刪除重複出現的名稱
            j = 0
            For i = 0 To lstFilters.ListCount - 1
            If (lstFilters.List(i) = rfiEach.Name) Then j = 1
            Next
            
            
            If (j = 0) Then
            
            lstFilters.AddItem rfiEach.Name
            'Debug.Print rfiEach.Name, lstFilters.Text
            End If
            
            End If
            
        Next
    End With


UI = 0
Call Timer1_Timer
Form1.Timer3.Enabled = True
End Sub

Private Sub Command1_Click(Index As Integer)
HScroll1.Value = Command1(Index).Caption
Call HScroll1_Change
Call cmdSnap_Click
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 
 Select Case KeyCode
 
    Case 37 'left vbkeyleft
    
        XX = XX - 1
        
        Picture1.Cls
        picSnapshot.Cls
        Call move_image(XX, YY)
        Text2.Text = CStr(XX)              'config
        Text4.Text = CStr(YY)              'config
        'Call cmdSnap_Click(XXX, YYY)
        
    Case 38 'up
    
        YY = YY - 1
        
        Picture1.Cls
        picSnapshot.Cls
        Call move_image(XX, YY)
        Text2.Text = CStr(XX)                'config
        Text4.Text = CStr(YY)                'config
       ' Call cmdSnap_Click(XXX, YYY)
    Case 39 'right
    
        XX = XX + 1
        
        Picture1.Cls
        picSnapshot.Cls
        Call move_image(XX, YY)
        Text2.Text = CStr(XX)                'config
        Text4.Text = CStr(YY)                'config
        'Call cmdSnap_Click(XXX, YYY)
        
    Case 40 'down
    
        YY = YY + 1
        
        Picture1.Cls
        picSnapshot.Cls
        Call move_image(XX, YY)
        Text2.Text = CStr(XX)                'config
        Text4.Text = CStr(YY)                'config
       ' Call cmdSnap_Click(XXX, YYY)
End Select
    

End Sub


Private Sub Form_Unload(Cancel As Integer)
Call StopCamera(lstFilters.Text)

End Sub

Private Sub HScroll1_Change()
Text1.Text = HScroll1.Value

End Sub

Private Sub HScroll1_Scroll()
Text1.Text = HScroll1.Value
End Sub

Sub Timer1_Timer()

On Error Resume Next

Call cmdSnap_Click

Label3.Caption = UI

lstFilters.ListIndex = UI Mod lstFilters.ListCount

Call StartCamera(lstFilters.Text)

UI = (UI + 1) Mod lstFilters.ListCount


End Sub

Private Sub Command2_Click()

 Form1.Timer3.Enabled = True
 Form1.Timer1_Timer
End Sub



Private Sub Timer3_Timer()                               '切換輪播


On Error Resume Next


'If lstFilters.ListCount = 10 Then



Call cmdSnap_Click

Label3.Caption = UI

lstFilters.ListIndex = UI Mod lstFilters.ListCount

Call StopCamera(lstFilters.Text)
                                                              '切換輪播 CAMERA STOP

'Delay 1

Call StartCamera(lstFilters.Text)

UI = (UI + 1) Mod lstFilters.ListCount

If UI = 1 Then

'Delay 1

'Dim StartResult As Integer
'Dim rfiEach As QuartzTypeLib.IRegFilterInfo
'Dim i, j As Integer


  '  InsideWidth = picSnapshot.Width - ScaleX(2, vbPixels, ScaleMode)
    
 
   ' lstFilters.Clear
   ' With New QuartzTypeLib.FilgraphManager
    '    For Each rfiEach In .RegFilterCollection
        
     '       '只列舉有需要的名稱
      '      If (InStr(UCase(rfiEach.Name), "BT CAMERA")) Then
            
            '刪除重複出現的名稱
       '     j = 0
        '    For i = 0 To lstFilters.ListCount - 1
         '   If (lstFilters.List(i) = rfiEach.Name) Then j = 1
          '  Next
            
            
           ' If (j = 0) Then
            
            'lstFilters.AddItem rfiEach.Name
           ' Debug.Print rfiEach.Name, lstFilters.Text
           ' End If
            
           ' End If
            
      '  Next
   ' End With

'End If

'Else

'Call CHECK_Click()
'Delay 10        '15m

'Call CHECK_Click

End If







End Sub

Private Sub Command5_Click()
  Form1.Timer3.Enabled = False
 
End Sub

Public Sub Delay(D_Long As Date)                        ' delay
Dim DelayTime
DelayTime = DateAdd("s", D_Long, Now)
While DelayTime > Now
DoEvents
Wend
End Sub

Public Function Write_xml()                             'XML
On Error GoTo ErrorHandler
Dim Err_string As String
Err_string = "Sub Write_xml() Fail"


    Dim docXML     As DOMDocument40
    Dim domRoot    As IXMLDOMElement
    Dim idx        As Integer
    Dim idy        As Integer
    Dim iEqpCnt    As Integer
    Dim iParamCnt  As Integer

    iParamCnt = UBound(data)
    Set docXML = CreateXmlDoc(domRoot, "RecipeBodyReport")
    
    '<glass_id>F2VSTKT0001</glass_id>
    '<group_id>001</group_id>
    '<lot_id>F2STK001N00</lot_id>
    '<product_id>TEQP_SPC</product_id>
    '<pfcd>PMCX</pfcd>
    '<route_no>QEQP001M</route_no>
    '<owner>PROD</owner>
    '<recipe_id>001</recipe_id>
    '<operation>1099</operation>
    '<mes_link_key>0001</mes_link_key>
    '<operator>PMCX0100</operator>

    
    
    AppendXmlItem domRoot, "glass_id", "F2VSTKT0001"
    AppendXmlItem domRoot, "group_id", "001"
    AppendXmlItem domRoot, "lot_id", "F2STK001N00"
    AppendXmlItem domRoot, "product_id", "TEQP_SPC"
    AppendXmlItem domRoot, "pfcd", "PMCX"
    AppendXmlItem domRoot, "eqp_id", "CNDC0601"
    AppendXmlItem domRoot, "ec_code", ""
    AppendXmlItem domRoot, "route_no", "QEQP001M"
    AppendXmlItem domRoot, "route_version", ""
    AppendXmlItem domRoot, "owner", "PROD"
    AppendXmlItem domRoot, "recipe_id", "001"
    AppendXmlItem domRoot, "operation", "1099"
    AppendXmlItem domRoot, "rtc_flag", ""
    AppendXmlItem domRoot, "pnp", ""
    AppendXmlItem domRoot, "chamber", "NNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNN"
    AppendXmlItem domRoot, "cassette_id", ""
    AppendXmlItem domRoot, "line_batch_id", ""
    AppendXmlItem domRoot, "split_id", ""
    AppendXmlItem domRoot, "cldate", Format(Now, "yyyy-mm-dd")
    AppendXmlItem domRoot, "cltime", Format(Now, "hh:nn:ss")
    AppendXmlItem domRoot, "mes_link_key", "0001"
    AppendXmlItem domRoot, "rework_count", ""
    AppendXmlItem domRoot, "operator", "PMCX0100"
    AppendXmlItem domRoot, "reserve_field_1", ""
    AppendXmlItem domRoot, "reserve_field_2", ""
    
    AppendXmlList domRoot, "datas"
    
    For idy = 1 To iParamCnt
        AppendXmlList domRoot, "iary"
        AppendXmlItem domRoot, "item_name", dataname(idy)
        AppendXmlItem domRoot, "item_type", "X"
        AppendXmlItem domRoot, "item_value", data(idy)
        Set domRoot = domRoot.parentNode
    Next idy
    
    '20180613101013_ECXREPR06_F186A129NF.xml
   If lstFilters.ListIndex < 1 Then
    
    docXML.save "C:\File\CF\EDC\CNDC0600\CNDC0601\" & Format(Now, "yyyymmddhhnnss") & "_" & "CNDC0601" & "_" & "F2VSTKT0001" & ".XML"
    
    Else
    
    docXML.save "C:\File\CF\EDC\CNDC0600\CNDC0602\" & Format(Now, "yyyymmddhhnnss") & "_" & "CNDC0602" & "_" & "F2VSTKT0001" & ".XML"
    
    End If
    Set docXML = Nothing
    Set domRoot = Nothing
    
Exit Function

ErrorHandler:
    MsgBox "Write xml fail"
End Function

Public Function CreateXmlDoc(ByRef oRoot As IXMLDOMElement, sFuncName As String) As DOMDocument40
    Dim TempDoc     As New DOMDocument40
    Dim TempElement As IXMLDOMElement
    Dim PI          As IXMLDOMProcessingInstruction
    
    On Error Resume Next
         
    TempDoc.preserveWhiteSpace = True
    
    Set PI = TempDoc.createProcessingInstruction("xml", "version=""1.0""")
    TempDoc.appendChild PI

    Set TempElement = TempDoc.createElement("EDC")
    Set oRoot = TempDoc.appendChild(TempElement)
'    oRoot.setAttribute "trxid", "00000003"    '測試
'    oRoot.setAttribute "function", sFuncName
    
    Set CreateXmlDoc = TempDoc
    
    Set TempElement = Nothing
    Set TempDoc = Nothing
End Function

Public Function AppendXmlList(ByRef oRoot As IXMLDOMElement, sTagName As String, Optional ListCount As Variant) As IXMLDOMElement
       Dim TempDoc     As New DOMDocument40
       Dim TempElement As IXMLDOMElement

       On Error Resume Next
    
       Set TempElement = TempDoc.createElement(sTagName)
    
       If Not IsMissing(ListCount) Then TempElement.setAttribute "count", ListCount
    
      Set AppendXmlList = TempElement
      Set oRoot = oRoot.appendChild(TempElement)
   
      Set TempElement = Nothing
      Set TempDoc = Nothing
End Function

Public Sub AppendXmlItem(ByRef oRoot As IXMLDOMElement, sTagName As String, sValue As String)
       Dim TempDoc     As New DOMDocument40
       Dim TempElement As IXMLDOMElement
    
       On Error Resume Next
 
       Set TempElement = TempDoc.createElement(sTagName)
    
       TempElement.Text = sValue
    
      oRoot.appendChild TempElement
    
      Set TempElement = Nothing
      Set TempDoc = Nothing
End Sub
