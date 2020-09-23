VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4215
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Show a Message Box"
      Height          =   585
      Left            =   855
      TabIndex        =   2
      Top             =   2415
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Click on the button below to send a modal message box from this form.  Again, the test form should not loose its color."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Index           =   1
      Left            =   270
      TabIndex        =   1
      Top             =   1380
      Width           =   3750
   End
   Begin VB.Label Label1 
      Caption         =   $"Form2.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Index           =   0
      Left            =   270
      TabIndex        =   0
      Top             =   225
      Width           =   3750
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
If MsgBox("Close this window?", vbYesNo + vbQuestion, "Testing") = vbYes Then Unload Me
End Sub
