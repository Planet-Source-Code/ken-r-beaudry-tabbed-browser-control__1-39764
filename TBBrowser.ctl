VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.UserControl TBBrowser 
   ClientHeight    =   5760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6630
   ScaleHeight     =   5760
   ScaleWidth      =   6630
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   2
      Top             =   5490
      Visible         =   0   'False
      Width           =   6630
      _ExtentX        =   11695
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9075
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   7
            Alignment       =   2
            Enabled         =   0   'False
            TextSave        =   "KANA"
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser browser 
      Height          =   1695
      Index           =   0
      Left            =   2.00000e5
      TabIndex        =   0
      Top             =   1170
      Width           =   4065
      ExtentX         =   7170
      ExtentY         =   2990
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin MSComctlLib.TabStrip btab 
      Height          =   2055
      Left            =   600
      TabIndex        =   1
      Top             =   840
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   3625
      TabWidthStyle   =   2
      ShowTips        =   0   'False
      TabFixedWidth   =   2117
      TabMinWidth     =   998
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Browser"
            Object.Tag             =   "1"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar Progress1 
      Height          =   255
      Left            =   4920
      TabIndex        =   3
      Top             =   3600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Menu mnuTabs 
      Caption         =   "tabs"
      Visible         =   0   'False
      Begin VB.Menu mnuTabDelete 
         Caption         =   "&Delete Active Tab"
      End
   End
End
Attribute VB_Name = "TBBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As _
Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, lParam _
As Any) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Const WM_USER As Long = &H400
Private Const SB_GETRECT As Long = (WM_USER + 10)
Public Enum PopupState
    No_Popups = 0
    Ask = 1
    Allow = 2
End Enum
Private m_shomepage As String
Private IsUnloading As Boolean
Private m_sCustomError As String
Private m_bStatusVisable As Boolean
Private m_sCurrentAddress As String
Private SelectedTab As Integer
Private m_inumtabs As Integer
Private m_iCurrentBrowser As Integer
Private m_ePopups As PopupState
Private w, h As Long

Private Sub browser_BeforeNavigate2(index As Integer, ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
If index = 0 Then Exit Sub
If btab.Tabs(CurrentBrowser).Tag = index Then
    browser(index).ZOrder
    If CurrentBrowser > 1 Then
        ResizeNew
    End If
'    Call browser(index).SetFocus
End If
If CurrentBrowser > 1 Then
    ResizeNew
End If
End Sub

Private Sub browser_CommandStateChange(index As Integer, ByVal Command As Long, ByVal Enable As Boolean)
Dim i As Integer
Select Case Command
    Case -1
        For i = 1 To btab.Tabs.Count
            If btab.Tabs(i).Tag = index Then
                If Len(browser(index).LocationName) < 16 Then
                    btab.Tabs(i).Caption = browser(index).LocationName
                Else
                    btab.Tabs(i).Caption = Left$(browser(index).LocationName, 15)
                End If
            End If
        Next
    End Select
End Sub

Private Sub browser_DocumentComplete(index As Integer, ByVal pDisp As Object, URL As Variant)
Dim q As Integer
If index = 0 Then Exit Sub
If CurrentBrowser > 1 Then ResizeNew
For i = 1 To btab.Tabs.Count
    q = btab.Tabs(i).Tag
    If q = index Then
        If Len(browser(index).LocationName) < 16 Then
        btab.Tabs(i).Caption = browser(index).LocationName
      Else
        btab.Tabs(i).Caption = Left$(browser(index).LocationName, 15)
      End If
    If CurrentBrowser > 1 Then ResizeNew
    'set the CurrentAddress Property to the current browser address
    CurrentAddress = browser(CurrentBrowser).LocationURL
    End If
Next
End Sub

Private Sub browser_DownloadBegin(index As Integer)
Dim i As Integer
Dim q As Integer
If index = 0 Then Exit Sub
For i = 1 To btab.Tabs.Count
    If btab.Tabs(i).Tag = index Then
        btab.Tabs(i).Caption = "Loading..."
        If CurrentBrowser > 1 Then ResizeNew
    End If
Next
End Sub

Private Sub browser_NavigateComplete2(index As Integer, ByVal pDisp As Object, URL As Variant)
Dim i As Integer
If index = 0 Then Exit Sub
If CurrentBrowser > 1 Then ResizeNew
For i = 1 To btab.Tabs.Count
    If btab.Tabs(i).Tag = index Then
        btab.Tabs(index).Caption = Left$(browser(index).LocationName, 15)
    End If
    If browser(index).LocationURL = "about:blank" Then
        btab.Tabs(index).Caption = "Blank"
    End If
Next
End Sub

Private Sub browser_NavigateError(index As Integer, ByVal pDisp As Object, URL As Variant, Frame As Variant, StatusCode As Variant, Cancel As Boolean)
If CustomError <> "" Then
    browser(index).Name CustomError
    Set ppDisp = browser(index).object
End If
End Sub

Private Sub browser_NewWindow2(index As Integer, ppDisp As Object, Cancel As Boolean)
Dim strLocationUrl, URL As String
If index = 0 Then Exit Sub
  If IsUnloading = True Then
    Cancel = True
    DoEvents
    Exit Sub
  End If
strLocationUrl = browser(btab.Tabs(CurrentBrowser).Tag).LocationURL
Select Case Popups

    Case No_Popups
        Cancel = True
        DoEvents
    Case Ask
        Select Case MsgBox("The following URL: " & strLocationUrl _
                           & vbCrLf & "Is attempting to open a new browser tab." _
                           & vbCrLf & "Allow popup?" _
                           , vbYesNo + vbQuestion + vbDefaultButton1, "Allow Popup")
        
            Case vbNo
                Cancel = True
                DoEvents
                Exit Sub
            Case vbYes
                NewTab URL
                If CurrentBrowser > 1 Then ResizeNew
                SelectTab (CurrentBrowser)
                Set ppDisp = browser(btab.Tabs(CurrentBrowser).Tag).object
                If CurrentBrowser > 1 Then ResizeNew
            
        End Select
        
    Case Allow
        NewTab URL
        If CurrentBrowser > 1 Then ResizeNew
        SelectTab (CurrentBrowser)
        Set ppDisp = browser(btab.Tabs(CurrentBrowser).Tag).object
        If CurrentBrowser > 1 Then ResizeNew
End Select

End Sub

Private Sub browser_ProgressChange(index As Integer, ByVal Progress As Long, ByVal ProgressMax As Long)
If index = 0 Then Exit Sub
If index = CurrentBrowser Then
    If browser(index).Busy And StatusVisable = True Then
        If Progress = -1 Then Progress1.Value = 0
        If Progress > 0 And ProgressMax > 0 Then
        If ProgressMax > Progress Then
            With Progress1
                .Max = ProgressMax
                .Value = Progress
            End With
        End If
        End If
    End If
End If
DoEvents
End Sub

Private Sub browser_StatusTextChange(index As Integer, ByVal Text As String)
If index = 0 Then Exit Sub
If index = CurrentBrowser Then
    StatusBar1.Panels(1).Text = Text
    If CurrentBrowser > 1 Then ResizeNew
End If

End Sub

Private Sub btab_GotFocus()
Dim i As Integer
For i = 1 To btab.Tabs.Count
    If btab.Tabs(i).Selected = True Then
        SelectedTab = i
    Else
        browser(btab.Tabs(i).Tag).Visible = False
    End If
Next
With browser(btab.Tabs(SelectedTab).Tag)
.Visible = True
.ZOrder
.SetFocus
CurrentAddress = .LocationURL
End With
CurrentBrowser = SelectedTab
DoEvents
End Sub

Private Sub btab_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    PopupMenu mnuTabs
End If
End Sub

Private Sub mnuTabDelete_Click()
DeleteTab
End Sub

Private Sub UserControl_InitProperties()
homepage = "www.msn.com"
StatusVisable = True
Popups = Allow
End Sub

Private Sub UserControl_Paint()
btab.Top = 0
btab.Left = 0
btab.Width = UserControl.Width
If StatusVisable = True Then
    StatusBar1.Visible = True
    btab.Height = UserControl.Height - StatusBar1.Height
Else
    StatusBar1.Visible = False
    btab.Height = UserControl.Height
End If
With browser(0)
.Top = 335
'.Left = 0
.Width = btab.Width - 30
.Height = btab.Height - 360
w = .Width: h = .Height
End With
If StatusBar1.Visible = True Then
    Call ShowProgressInStatusBar(True)
End If

End Sub

Public Sub SelectTab(index As Integer)
If index > numtabs Then
    Call MsgBox("The tab that you selected" & vbCrLf & "is outof range" _
                , vbCritical, "Error Selecting Tab")
    Exit Sub
End If
btab.Tabs(index).Selected = True
browser(btab.Tabs(CurrentBrowser).Tag).Visible = False
browser(btab.Tabs(index).Tag).Visible = True
browser(btab.Tabs(index).Tag).ZOrder
browser(btab.Tabs(index).Tag).SetFocus
CurrentBrowser = index
ResizeNew
CurrentAddress = browser(btab.Tabs(CurrentBrowser).Tag).LocationURL
End Sub

Private Sub ResizeNew()
With browser(btab.Tabs(CurrentBrowser).Tag)
.Top = 355
.Left = 0
.Width = browser(0).Width
.Height = browser(0).Height

End With

End Sub

Public Sub NewTab(URL As String, Optional Options As Integer)
'Adds a new tab and browser to the control
Dim tmp As Integer
numtabs = btab.Tabs.Count
If Tabs <> 50 Then          ' 50 tabs max
    btab.Tabs.Add
    numtabs = numtabs + 1
    CurrentBrowser = numtabs
    With btab.Tabs(CurrentBrowser)
    .Caption = "New Page"
    .Selected = False
    .Tag = FirstIndex
    tmp = .Tag
End With
' load a new browser control with the first available index
Load browser(btab.Tabs(CurrentBrowser).Tag)
browser(tmp).Visible = True
browser(tmp).ZOrder
ResizeNew
SelectTab CurrentBrowser
DoEvents                       'let other events process for a second
If URL = "" Then
    btab.Tabs(CurrentBrowser).Caption = "Blank"
    browser(tmp).Navigate "about:blank"
    Exit Sub
End If
' you can add specific options here to tailor this code to your needs
' for example:
'
' select case Options
'       case 1          'Navigate to homepage
'           browser(tmp).Navigate "www.mypage.com"
'       case 2          'Navigate to GOOGLE
'           browser(tmp).Navigate "www.google.com"
' end select
End If
End Sub

Private Function FirstIndex() As Integer
'Finds first unused or released index of the browser control array
Dim found As Boolean
Dim i As Integer
For Each object In browser()
    If object.index <> i Then
        found = True
        Exit For
    End If
    i = i + 1
Next
If found = True Then
    FirstIndex = i
Else
FirstIndex = numtabs
End If
End Function

Public Property Get numtabs() As Integer

    numtabs = m_inumtabs

End Property

Public Property Let numtabs(ByVal inumtabs As Integer)

    m_inumtabs = inumtabs

    Call UserControl.PropertyChanged("numtabs")
    UserControl_Paint
End Property

Public Property Get CurrentBrowser() As Integer

    CurrentBrowser = m_iCurrentBrowser

End Property

Public Property Let CurrentBrowser(ByVal iCurrentBrowser As Integer)

    m_iCurrentBrowser = iCurrentBrowser

    Call UserControl.PropertyChanged("CurrentBrowser")
    UserControl_Paint
End Property

Public Property Get Popups() As PopupState

    Popups = m_ePopups

End Property

Public Property Let Popups(ByVal ePopups As PopupState)

    m_ePopups = ePopups

    Call UserControl.PropertyChanged("Popups")
    UserControl_Paint
End Property

Public Property Get CurrentAddress() As String

    CurrentAddress = m_sCurrentAddress

End Property

Public Property Let CurrentAddress(ByVal sCurrentAddress As String)

    m_sCurrentAddress = sCurrentAddress

    Call UserControl.PropertyChanged("CurrentAddress")
    UserControl_Paint
End Property

Public Property Get StatusVisable() As Boolean

    StatusVisable = m_bStatusVisable
    StatusBar1.Visible = m_bStatusVisable
End Property

Public Property Let StatusVisable(ByVal bStatusVisable As Boolean)

    m_bStatusVisable = bStatusVisable

    Call UserControl.PropertyChanged("StatusVisable")
    StatusBar1.Visible = StatusVisable
    UserControl_Paint                   ' repaint the user control
End Property

Public Property Get CustomError() As String

    CustomError = m_sCustomError

End Property

Public Property Let CustomError(ByVal sCustomError As String)

    m_sCustomError = sCustomError

    Call UserControl.PropertyChanged("CustomError")
    UserControl_Paint
End Property

Public Property Get homepage() As String

    homepage = m_shomepage

End Property

Public Property Let homepage(ByVal shomepage As String)

    m_shomepage = shomepage

    Call UserControl.PropertyChanged("homepage")

End Property

Private Sub ShowProgressInStatusBar(ByVal bShowProgressBar As Boolean)

    Dim tRC As RECT

    If bShowProgressBar Then
        SendMessageAny StatusBar1.hwnd, SB_GETRECT, 1, tRC
        With tRC
            .Top = (.Top * Screen.TwipsPerPixelY)
            .Left = (.Left * Screen.TwipsPerPixelX)
            .Bottom = (.Bottom * Screen.TwipsPerPixelY) - .Top
            .Right = (.Right * Screen.TwipsPerPixelX) - .Left
        End With
        With Progress1
            SetParent .hwnd, StatusBar1.hwnd
            .Move tRC.Left, tRC.Top, tRC.Right, tRC.Bottom
            .Visible = True
            .Value = 0
        End With

    Else
        '
        ' Reparent the progress bar back to the form and hide it
        '
        SetParent Progress1.hwnd, UserControl.hwnd
        Progress1.Visible = False
    End If

End Sub

Public Sub InitControl(StartUrl As String)
numtabs = 1
CurrentBrowser = 1
Load browser(1)
SelectedTab = 1
btab.Tabs(1).Tag = 1
ResizeNew
browser(1).Visible = True
browser(1).Navigate StartUrl
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
    m_sCustomError = .ReadProperty("CustomError", "")
    m_bStatusVisable = .ReadProperty("StatusVisable", False)
    m_sCurrentAddress = .ReadProperty("CurrentAddress", "")
    m_inumtabs = .ReadProperty("numtabs", 0)
    m_iCurrentBrowser = .ReadProperty("CurrentBrowser", 0)
    m_ePopups = .ReadProperty("Popups", 0)
    m_shomepage = .ReadProperty("homepage", "www.msn.com")
End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
    .WriteProperty "CustomError", m_sCustomError
    .WriteProperty "StatusVisable", m_bStatusVisable
    .WriteProperty "CurrentAddress", m_sCurrentAddress
    .WriteProperty "numtabs", m_inumtabs
    .WriteProperty "CurrentBrowser", m_iCurrentBrowser
    .WriteProperty "Popups", m_ePopups
    .WriteProperty "homepage", m_shomepage
End With
End Sub

Public Sub Back()
On Error GoTo q1
browser(btab.Tabs(CurrentBrowser).Tag).GoBack
q1:
End Sub

Public Sub Forward()
On Error GoTo q1
browser(btab.Tabs(CurrentBrowser).Tag).GoForward
q1:
End Sub

Public Sub Home()
On Error GoTo q1
browser(btab.Tabs(CurrentBrowser).Tag).Navigate homepage
q1:
End Sub
Public Sub Refresh()
browser(btab.Tabs(CurrentBrowser).Tag).Refresh
End Sub
Public Sub Stop1()
browser(btab.Tabs(CurrentBrowser).Tag).Stop
End Sub

Public Sub DeleteTab()
Dim i As Integer
browser(btab.Tabs(CurrentBrowser).Tag).Stop
If browser.Count > 2 Then
    browser(btab.Tabs(CurrentBrowser).Tag).Visible = False
    Unload browser(btab.Tabs(CurrentBrowser).Tag)
    btab.Tabs.Remove (CurrentBrowser)
    i = 0
    numtabs = numtabs - 1
    If btab.Tabs.Count < CurrentBrowser Then
        CurrentBrowser = CurrentBrowser - 1
    End If
    SelectTab (CurrentBrowser)
End If
End Sub
