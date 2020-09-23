VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8685
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   8685
   StartUpPosition =   3  'Windows Default
   Begin Project1.TBBrowser TBBrowser 
      Height          =   5775
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   10186
      CustomError     =   ""
      StatusVisable   =   -1  'True
      CurrentAddress  =   ""
      numtabs         =   0
      CurrentBrowser  =   0
      Popups          =   1
      homepage        =   "www.msn.com"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
TBBrowser.InitControl "www.msn.com"
End Sub

