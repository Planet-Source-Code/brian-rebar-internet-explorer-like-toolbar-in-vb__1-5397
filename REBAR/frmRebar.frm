VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmRebar 
   Caption         =   "Rebar Sample Project"
   ClientHeight    =   2730
   ClientLeft      =   3660
   ClientTop       =   2520
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   ScaleHeight     =   2730
   ScaleWidth      =   7740
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   1035
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   3195
      TabIndex        =   2
      Top             =   1380
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   960
      Width           =   6735
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   840
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   600
      Width           =   6735
   End
   Begin ComctlLib.ImageList ImageList3 
      Left            =   1560
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   11
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmRebar.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmRebar.frx":005E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmRebar.frx":00BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmRebar.frx":011A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmRebar.frx":0178
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmRebar.frx":01D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmRebar.frx":0234
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmRebar.frx":0292
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmRebar.frx":02F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmRebar.frx":034E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmRebar.frx":03AC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   840
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmRebar.frx":040A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   8
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmRebar.frx":0468
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmRebar.frx":04C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmRebar.frx":0524
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmRebar.frx":0582
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmRebar.frx":05E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmRebar.frx":063E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmRebar.frx":069C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmRebar.frx":06FA
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmRebar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Rebar As New CRebar
Dim RebarHeight As Integer
 
 
Public Sub RefreshToolbars()
On Error Resume Next

Dim xCount As Integer
For xCount = 0 To Toolbar.Count - 1
             Toolbar(xCount).Refresh
Next

End Sub

 


 
 
 

Private Sub Form_Load()
'Written by Ramon Guerrero
'ZoneCorp@dallas.net
'ZoneCorp@Aol.com
'ZoneCOrp@Compuserve.com

'Toolbars our place in Picture control to
'hide the toolbar borders

'move toolbar over to left to hide left border
'make toolbar's width wide enought to hide right border


'make Pictbcontainer(2).height smaller because we are
'making this toolbar's style list.
'It makes the pictures be on the left and the text on the right

'Create The Rebar
With Rebar
Set .Parent = Me
    .Create
End With
 
'Add the bands with the child
Rebar.AddBands "Address", 1, Text1.hwnd, 0
Rebar.AddBands "Address 2", 2, Text2.hwnd, AddToEnd
 
'Activate the forms Subclass Routine
SubClass Me.hwnd
 
Me.Show
CalculateRebarHeight
End Sub


Public Sub ProcMsg(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, Result As Long)
  
On Error Resume Next
Dim hdr As NMHDR
Dim RebarHdr As NMREBAR
Dim BandInfo As RBHITTESTINFO
Dim ptApi As POINTAPI
Dim xReturn As Long
Dim BandId As Integer

Select Case uMsg


Case WM_NOTIFY 'Needed to let us know when mouse has anything to do with Rebar
'Copy hdr info so we can determine if uMsg is coming from Rebar
'CopyMemory hdr, ByVal lParam, Len(hdr)
CopyMemory RebarHdr, ByVal lParam, Len(RebarHdr)
 
'Check hwndFrom (handle of window sending message)
'If hdr.hwndFrom = Rebar.GetRebarWindow Then
If RebarHdr.NMHDR = Rebar.GetRebarWindow Then
Call GetCursorPos(ptApi)
Call ScreenToClient(Me.hwnd, ptApi)
BandInfo.ptApi = ptApi
BandInfo.flags = RBHT_CAPTION Or RBHT_GRABBER Or RBHT_CLIENT
Call SendMessage(Rebar.GetRebarWindow, RB_HITTEST, 0, BandInfo)
'Yes it's ours
'8386744 = Being Sized
'8387324 = ClickUp anywhere on rebar or gripper
'If you don't do this when using the toolbar control, then
'whenever you touch the Rebar or size the bands then
'toolbars will dissappear.
'Alot of Flicker
If lParam = 8386584 Then CalculateRebarHeight
If lParam = 8387336 Then CalculateRebarHeight
End If

'Case User changes colors while this is running
Case WM_SYSCOLORCHANGE
Rebar.SetBandColors


End Select

 
End Sub


Public Sub SubClass(hwnd As Long)
NextProcs = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub
Public Sub UnSubClass()
Dim hWndCur As Long
    hWndCur = Me.hwnd
    If NextProcs Then
        SetWindowLong hWndCur, GWL_WNDPROC, NextProcs
        NextProcs = 0
    End If
End Sub

Private Sub Form_Resize()
If Me.Height < 3000 Then Me.Height = 3000
If Me.Width < 3000 Then Me.Width = 3000
Rebar.Resize Me
Picture1.Width = Me.ScaleWidth
Picture1.Height = Me.ScaleHeight - Picture1.Top - 10
End Sub


Private Sub Form_Unload(Cancel As Integer)
 UnSubClass
 Rebar.DestroyRebar
End Sub

 Sub CalculateRebarHeight()
RebarHeight = Rebar.GetHeight()
Picture1.Top = RebarHeight + 5
Picture1.Height = Me.ScaleHeight - RebarHeight - 10
 End Sub
 

