Attribute VB_Name = "modGeneral"
Option Explicit

Type Star
    X As Integer
    Y As Integer
    Color As Long
End Type

'these are the speeds the user will set and the planes will move in
Dim Plane1Velocity
Dim Plane2Velocity
Dim Plane3Velocity
Dim Plane4Velocity

'this is an array used throughout the program
Public stars(1 To 4, 1 To 100) As Star

'this initializes the stars
Sub InitStars()
Dim i As Integer
Dim j As Integer

    For i = 1 To 4
        For j = 1 To 100
        
        'this randomizes the stars
        Randomize
        stars(i, j).X = Int((frmMain.ScaleWidth * Rnd) + 1)
        stars(i, j).Y = Int((frmMain.ScaleHeight * Rnd) + 1)
        
        'this sets the stars colors per plane
        Select Case i
        Case 1
        stars(i, j).Color = RGB(50, 50, 50)
        Case 2
        stars(i, j).Color = RGB(75, 75, 75)
        Case 3
        stars(i, j).Color = RGB(150, 150, 150)
        Case 4
        stars(i, j).Color = RGB(255, 255, 255)
        End Select
        
        Next j
    Next i
    
End Sub

'this draws the stars and moves them
Sub DrawStars()
Dim i As Integer
Dim j As Integer
Dim bgColor As Long
    
    bgColor = frmMain.BackColor
    
    'this sets the planes velocity to what the user typed into the options form
    Plane1Velocity = frmOptions.txtPlane1.Text
    Plane2Velocity = frmOptions.txtPlane2.Text
    Plane3Velocity = frmOptions.txtPlane3.Text
    Plane4Velocity = frmOptions.txtPlane4.Text

    On Error Resume Next
    
    For i = 1 To 4
        For j = 1 To 100
        
        'this erases the stars
        frmMain.PSet (stars(i, j).X, stars(i, j).Y), bgColor
        
        'if the horizontal direction was chosen then
        If frmOptions.optHorizontal.Value = True Then
            
            'this moves the stars to their certain planes speed setting
            Select Case i
            Case 1
            stars(i, j).X = stars(i, j).X + Plane1Velocity
            Case 2
            stars(i, j).X = stars(i, j).X + Plane2Velocity
            Case 3
            stars(i, j).X = stars(i, j).X + Plane3Velocity
            Case 4
            stars(i, j).X = stars(i, j).X + Plane4Velocity
            End Select
            
            'this wraps the stars
            If stars(i, j).X > frmMain.ScaleWidth Then
                stars(i, j).X = frmMain.ScaleLeft
                    
            ElseIf stars(i, j).X < 1 Then
                stars(i, j).X = frmMain.ScaleWidth
                
            End If
        
        'if the diagonal direction was chosen then
        ElseIf frmOptions.optDiagonal.Value = True Then
            
            'this moves the stars to their certain planes speed setting
            Select Case i
            Case 1
            stars(i, j).X = stars(i, j).X + Plane1Velocity
            Case 2
            stars(i, j).X = stars(i, j).X + Plane2Velocity
            Case 3
            stars(i, j).X = stars(i, j).X + Plane3Velocity
            Case 4
            stars(i, j).X = stars(i, j).X + Plane4Velocity
            End Select
            
            'this moves the stars to their certain planes speed setting
            Select Case i
            Case 1
            stars(i, j).Y = stars(i, j).Y + Plane1Velocity
            Case 2
            stars(i, j).Y = stars(i, j).Y + Plane2Velocity
            Case 3
            stars(i, j).Y = stars(i, j).Y + Plane3Velocity
            Case 4
            stars(i, j).Y = stars(i, j).Y + Plane4Velocity
            End Select
            
            'this wraps the stars
            If stars(i, j).X > frmMain.ScaleWidth Then
                stars(i, j).X = frmMain.ScaleLeft
                    
            ElseIf stars(i, j).X < 1 Then
                stars(i, j).X = frmMain.ScaleWidth
                    
            ElseIf stars(i, j).Y > frmMain.ScaleHeight Then
                stars(i, j).Y = frmMain.ScaleTop
                    
            ElseIf stars(i, j).Y < 1 Then
                stars(i, j).Y = frmMain.ScaleHeight
                
            End If
        
        'if the vertical direction was chosen then
        ElseIf frmOptions.optVertical.Value = True Then
            
            'this moves the stars to their certain planes speed setting
            Select Case i
            Case 1
            stars(i, j).Y = stars(i, j).Y + Plane1Velocity
            Case 2
            stars(i, j).Y = stars(i, j).Y + Plane2Velocity
            Case 3
            stars(i, j).Y = stars(i, j).Y + Plane3Velocity
            Case 4
            stars(i, j).Y = stars(i, j).Y + Plane4Velocity
            End Select
            
            'this wraps the stars
            If stars(i, j).Y > frmMain.ScaleHeight Then
                stars(i, j).Y = frmMain.ScaleTop
                    
            ElseIf stars(i, j).Y < 1 Then
                stars(i, j).Y = frmMain.ScaleHeight
                
            End If
        
        End If
        
        'this draws the stars
        frmMain.PSet (stars(i, j).X, stars(i, j).Y), stars(i, j).Color

        Next j
    Next i
    
End Sub
