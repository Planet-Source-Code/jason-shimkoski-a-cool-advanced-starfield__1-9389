VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Help"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5190
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "Groovy !"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox txtHelp 
      Height          =   1695
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   840
      Width           =   4935
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      Caption         =   "The Advance Starfield !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4980
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    txtHelp.Text = "This is the updated version of my Starfield. In this" + _
                    " version you can change the direction of the stars movement, the" + _
                    " speed at which they travel, and the background's color." + vbCrLf + vbCrLf + _
                    "If you would like one of the planes to move in an opposite direction, simply put" + _
                    " a negative number for that planes speed." + vbCrLf + vbCrLf + _
                    "To start the starfields movement, click anywhere on the main form." + vbCrLf + vbCrLf + _
                    "The largest update for this version is the addition of the ability to change the Form's" + _
                    " background color and the fact that it doesn't require a picture box control any longer." + _
                    vbCrLf + vbCrLf + "Thanks," + vbCrLf + _
                    "Jason Shimkoski (basspler@aol.com)"

End Sub
