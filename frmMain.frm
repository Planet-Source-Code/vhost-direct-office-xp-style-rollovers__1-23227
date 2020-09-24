VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Office XP Style Rollovers"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   3135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   3135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   3255
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "frmMain.frx":0000
      Top             =   60
      Width           =   3015
   End
   Begin VB.PictureBox Picture1 
      Height          =   795
      Left            =   60
      ScaleHeight     =   735
      ScaleWidth      =   2955
      TabIndex        =   0
      Top             =   3720
      Width           =   3015
      Begin OfficeXPStyleRollovers.XPBar XPBar1 
         Height          =   495
         Left            =   180
         TabIndex        =   1
         Top             =   120
         Width           =   2655
         _extentx        =   4683
         _extenty        =   873
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Example:"
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   3480
      Width           =   915
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
