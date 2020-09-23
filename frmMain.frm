VERSION 5.00
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   6900
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9060
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pctSplitterPanel 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   6900
      Left            =   0
      ScaleHeight     =   6900
      ScaleWidth      =   2415
      TabIndex        =   0
      Top             =   0
      Width           =   2415
      Begin VB.TextBox Text1 
         Height          =   2175
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   1
         Text            =   "frmMain.frx":0000
         Top             =   240
         Width           =   1815
      End
      Begin VB.Image imgSplitter 
         Height          =   5775
         Left            =   2160
         MouseIcon       =   "frmMain.frx":00A5
         MousePointer    =   99  'Custom
         Top             =   0
         Width           =   75
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'api declares for stopflicker
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long

Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim startx As Integer
If Button = vbLeftButton Then
    'get the current position of the right edge of the image control
    startx = imgSplitter.Left + imgSplitter.Width
    ' get the position to where the mouse has moved
    startx = startx + X
    
    'Sets the splitter panel to the width between the left edge of
    'the panel and startx which is where the mouse currently is
    
    pctSplitterPanel.Width = startx
    ' this section sets the splitter width limits
    'This makes the splitter only able to move to half the
    'width of the main form.
    'You can try changing these values to see what happens
    If pctSplitterPanel.Width > frmMain.Width / 2 Then
        
        'really messy if this function is not here
        'I found this on PCS several years ago and it comes in
        'very handy for keeping things looking neet
        StopFlicker Me.hWnd
        
        pctSplitterPanel.Width = frmMain.Width / 2 - 1
   
        Exit Sub
    'This sets the minimum width that the splitter can be
    ElseIf pctSplitterPanel.Width < 2000 Then
        StopFlicker Me.hWnd
        
        pctSplitterPanel.Width = 2000
    Else
        'releases the Stopflicker
        ReleaseFlicker
    End If
    
End If

End Sub
Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseFlicker
End Sub

Private Sub MDIForm_Load()

Form1.Show

End Sub

Private Sub pctSplitterPanel_Resize()
On Error Resume Next
'this sets the controls in the splitter to the desired size
Text1.Width = pctSplitterPanel.ScaleWidth - imgSplitter.Width
Text1.Height = pctSplitterPanel.ScaleHeight
Text1.Left = pctSplitterPanel.ScaleLeft
Text1.Top = pctSplitterPanel.ScaleTop
imgSplitter.Height = pctSplitterPanel.ScaleHeight
imgSplitter.Left = Text1.Left + Text1.Width

End Sub

Private Sub StopFlicker(ByVal lHWnd As Long)
    Dim lRet As Long
    'Object will not flicker - just be blank
    '
    lRet = LockWindowUpdate(lHWnd)
End Sub
Private Sub ReleaseFlicker()
    Dim lRet As Long
    lRet = LockWindowUpdate(0)
End Sub
