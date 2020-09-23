VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "New Starfield"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4710
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   217
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   314
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Form_Resize
    
    'this initializes the stars
    InitStars
End Sub

Private Sub Form_Resize()
    
    'this gets rid of any errors on maximize and minimize
    On Error Resume Next
    
    'this resizes the form to fit the screen
    With frmMain
    .Top = 0
    .Left = 0
    .Width = Screen.Width
    .Height = Screen.Height
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmOptions.Show
End Sub

Private Sub Form_Click()
    'this is the main starfield loop
    Do
        DoEvents
        DrawStars
    Loop
End Sub
