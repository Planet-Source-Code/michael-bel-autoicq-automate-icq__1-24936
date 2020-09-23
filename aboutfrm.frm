VERSION 5.00
Begin VB.Form aboutfrm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About AutoICQ"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5550
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK "
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "AutoICQ BETA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   240
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   240
      Picture         =   "aboutfrm.frx":0000
      Top             =   240
      Width           =   450
   End
   Begin VB.Label Label1 
      Caption         =   $"aboutfrm.frx":0B0A
      ForeColor       =   &H00000000&
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   4815
   End
End
Attribute VB_Name = "aboutfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload aboutfrm
End Sub

