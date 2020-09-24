VERSION 5.00
Begin VB.UserControl XPButton 
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   375
   ScaleHeight     =   315
   ScaleWidth      =   375
   Begin VB.Image hovercontrol 
      Height          =   315
      Left            =   0
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   225
      Left            =   60
      Picture         =   "XPButton.ctx":0000
      Stretch         =   -1  'True
      Top             =   60
      Width           =   240
   End
   Begin VB.Shape hoveredge 
      BackColor       =   &H00C0C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808000&
      Height          =   315
      Left            =   0
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "XPButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim hovercolor As OLE_COLOR
'
Event mousemove(Button As Integer, Shift As Integer, x As Single, y As Single)
'
Private Sub hovercontrol_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 RaiseEvent mousemove(Button, Shift, x, y)
End Sub

Private Sub UserControl_Initialize()
 hoveredge.Visible = False
End Sub

Public Property Let hover_over(ByVal doit1 As Boolean)
 hoveredge.Visible = doit1
End Property

Private Sub UserControl_Click()
 hoveredge.FillColor = hovercolor
End Sub

Private Sub UserControl_Resize()
  hoveredge.Width = UserControl.Width - (hoveredge.Left * 2)
  hoveredge.Height = UserControl.Height - (hoveredge.Top * 2)
End Sub

