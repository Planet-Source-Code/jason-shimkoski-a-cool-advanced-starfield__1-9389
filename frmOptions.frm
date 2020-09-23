VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3855
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Background Color"
      Height          =   2055
      Left            =   2040
      TabIndex        =   20
      Top             =   2520
      Width           =   1695
      Begin VB.OptionButton optGreen 
         Caption         =   "Green"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton optBlack 
         Caption         =   "Black"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optBlue 
         Caption         =   "Blue"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1560
         Width           =   1215
      End
      Begin VB.OptionButton optRed 
         Caption         =   "Red"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   1215
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   105
         X2              =   1575
         Y1              =   705
         Y2              =   705
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   120
         X2              =   1605
         Y1              =   720
         Y2              =   720
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help !"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   4680
      Width           =   3615
   End
   Begin VB.Frame fraDirection 
      Caption         =   "Star Direction"
      Height          =   1575
      Left            =   120
      TabIndex        =   19
      Top             =   2520
      Width           =   1815
      Begin VB.OptionButton optVertical 
         Caption         =   "Vertical"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton optDiagonal 
         Caption         =   "Diagonal"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optHorizontal 
         Caption         =   "Horizontal"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame fraSpeed 
      Caption         =   "Plane Speeds (from back to front)"
      Height          =   2295
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   3615
      Begin VB.TextBox txtPlane2 
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Text            =   "2"
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox txtPlane4 
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Text            =   "4"
         Top             =   1800
         Width           =   2175
      End
      Begin VB.TextBox txtPlane3 
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Text            =   "3"
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox txtPlane1 
         Height          =   285
         Left            =   1320
         TabIndex        =   0
         Text            =   "1"
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Plane 4 Speed:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Plane 3 Speed:"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Plane 2 Speed:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Plane 1 Speed:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit !"
      Height          =   375
      Left            =   1080
      TabIndex        =   12
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Alright !"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   4200
      Width           =   855
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    frmHelp.Show
End Sub

Private Sub cmdOK_Click()
    
    'this sets the background color of the main form
    If optBlack.Value = True Then
        frmMain.BackColor = RGB(0, 0, 0)
        
    ElseIf optRed.Value = True Then
        frmMain.BackColor = RGB(40, 0, 0)
        
    ElseIf optGreen.Value = True Then
        frmMain.BackColor = RGB(0, 40, 0)
        
    ElseIf optBlue.Value = True Then
        frmMain.BackColor = RGB(0, 0, 40)
        
    End If
    
    'if any of the fields are left blank, don't let the form show
    If txtPlane1.Text = "" Or txtPlane2.Text = "" Or txtPlane3.Text = "" Or txtPlane4.Text = "" Then
        MsgBox "Please Leave no Blank Fields!", vbCritical, "Hey Man!"
    Else
        frmMain.Show
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
