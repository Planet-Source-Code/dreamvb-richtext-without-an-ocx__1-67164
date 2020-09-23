VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Demo"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   6780
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   2610
      TabIndex        =   2
      Top             =   135
      Width           =   1215
   End
   Begin VB.CommandButton cmdget 
      Caption         =   "Get Text"
      Height          =   495
      Left            =   1335
      TabIndex        =   1
      Top             =   150
      Width           =   1215
   End
   Begin VB.CommandButton cmdset 
      Caption         =   "Set Text"
      Height          =   495
      Left            =   45
      TabIndex        =   0
      Top             =   150
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function SetWindowText Lib "user32.dll" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef Destination As Any, ByVal Length As Long)
Private Declare Function DestroyWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

Private Const WS_CHILDWINDOW As Long = &H40000000
Private Const WS_BORDER As Long = &H800000
Private Const WS_VISIBLE As Long = &H10000000
Private Const ES_MULTILINE As Long = &H4&
Private Const WS_VSCROLL As Long = &H200000

Private Type MyWindowType
    wParentHwnd As Long
    wRTFHwnd As Long
    wLeft As Long
    wTop As Long
    wWidth As Long
    wHeight As Long
End Type

Private mRTFWinType As MyWindowType

Private Function OpenFile(lFile As String) As String
Dim fp As Long
Dim sData As String
    fp = FreeFile
    
    Open lFile For Binary As #fp
        sData = Space(LOF(fp))
        Get #fp, , sData
    Close #fp
    
    OpenFile = sData
    sData = ""
End Function

Private Sub DestroyRTFWindow()
    'Destroy the window
    DestroyWindow mRTFWinType.wRTFHwnd
    ZeroMemory mRTFWinType, Len(mRTFWinType)
End Sub

Private Sub CreateRTFWindow(mWinType As MyWindowType)
Dim wStyle As Long
    'Window Style
    wStyle = WS_CHILDWINDOW Or WS_BORDER Or WS_VISIBLE Or ES_MULTILINE Or WS_VSCROLL
    'Create the RichText Window
    With mRTFWinType

        .wRTFHwnd = CreateWindowEx(&H200&, "RichEdit20A", "" _
        , wStyle, .wLeft, .wTop, .wWidth, .wHeight, .wParentHwnd, 0, App.hInstance, ByVal 0&)
    End With
End Sub

Private Sub SetRTFText(sText As String)
    'Sets text to the RTF Window
    Call SetWindowText(mRTFWinType.wRTFHwnd, sText)
End Sub

Private Function GetRTFText() As String
Dim tLen As Long
Dim sBuff As String
    'Return Plain Text of the RTF Window
    tLen = GetWindowTextLength(mRTFWinType.wRTFHwnd) + 1
    'Create buffer to hold the text
    sBuff = Space(tLen)
    'Get the windows text
    Call GetWindowText(mRTFWinType.wRTFHwnd, sBuff, tLen)
    'Return the text
    GetRTFText = Left(sBuff, InStr(1, sBuff, Chr(0)) - 1)
    sBuff = ""
    tLen = 0
End Function

Private Sub cmdexit_Click()
    Call DestroyRTFWindow
    Unload Form1
End Sub

Private Sub cmdget_Click()
    MsgBox GetRTFText, vbInformation, "Get Text"
End Sub

Private Sub cmdset_Click()
    'Display some simple text
    Call SetRTFText(OpenFile(App.Path & "\example.rtf"))
End Sub

Private Sub Form_Load()
    With mRTFWinType
        .wParentHwnd = Me.hwnd
        .wLeft = 0
        .wTop = 60
        .wHeight = (Me.ScaleHeight \ Screen.TwipsPerPixelY) - .wTop
        .wWidth = (Me.ScaleWidth \ Screen.TwipsPerPixelX)
        'We first must load the libary
        If LoadLibrary("riched20.dll") = 0 Then
            MsgBox "Faild to Load Library" & vbCrLf & "riched20.dll", vbCritical, "Class Not Created"
            Exit Sub
        End If
        
        'Create the Window
        Call CreateRTFWindow(mRTFWinType)
        
        'Test that the window is created we should get a none zero if all went well
        If (.wRTFHwnd = 0) Then
            MsgBox "Faild to create RichEdit Class.", vbCritical, "Class Not Created"
            Exit Sub
        End If
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Form1 = Nothing
End Sub
