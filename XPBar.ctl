VERSION 5.00
Begin VB.UserControl XPBar 
   ClientHeight    =   510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2670
   ScaleHeight     =   510
   ScaleWidth      =   2670
   Begin OfficeXPStyleRollovers.XPButton XPButton 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin OfficeXPStyleRollovers.XPButton XPButton 
      Height          =   375
      Index           =   1
      Left            =   540
      TabIndex        =   1
      Top             =   60
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin OfficeXPStyleRollovers.XPButton XPButton 
      Height          =   375
      Index           =   2
      Left            =   960
      TabIndex        =   2
      Top             =   60
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin OfficeXPStyleRollovers.XPButton XPButton 
      Height          =   375
      Index           =   3
      Left            =   1380
      TabIndex        =   3
      Top             =   60
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin OfficeXPStyleRollovers.XPButton XPButton 
      Height          =   375
      Index           =   4
      Left            =   1800
      TabIndex        =   4
      Top             =   60
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin OfficeXPStyleRollovers.XPButton XPButton 
      Height          =   375
      Index           =   5
      Left            =   2220
      TabIndex        =   5
      Top             =   60
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin VB.Shape bgBorder 
      BorderColor     =   &H80000014&
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   2655
   End
   Begin VB.Line Line1 
      X1              =   60
      X2              =   60
      Y1              =   60
      Y2              =   420
   End
End
Attribute VB_Name = "XPBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sub UserControl_Resize()
bgBorder.Width = UserControl.Width - 10
bgBorder.Height = UserControl.Height - 10
End Sub

Private Sub XPButton_mousemove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
For a = 0 To 5
 If a <> Index Then
  XPButton(a).hover_over = False
 Else
  XPButton(Index).hover_over = True
 End If
Next a
End Sub
