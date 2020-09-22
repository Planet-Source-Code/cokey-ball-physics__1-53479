Attribute VB_Name = "Module1"
Public Const Pi = 3.141592653   'Set up a constant for Pi

Public Sub Collide(Ball1 As clsBall, Ball2 As clsBall)
Dim tmpX As Double, tmpY As Double  'Temp values for the x and y speed for ball 1
Dim tmpX2 As Double, tmpY2 As Double    'Temp values for the x and y speed for ball 1
    
    'This routine basically bounces the 2 balls of each other in the correct
    'direction with the correct velocity
    'If physics were not your thing at school, this might look nasty! :)
    tmpX = (Ball1.xSpeed - Ball2.xSpeed) * (Sin(Angle(Ball1, Ball2))) ^ 2 - (Ball1.ySpeed - Ball2.ySpeed) * Sin(Angle(Ball1, Ball2)) * Cos(Angle(Ball1, Ball2)) + Ball2.xSpeed
    tmpY = (Ball1.ySpeed - Ball2.ySpeed) * (Cos(Angle(Ball1, Ball2))) ^ 2 - (Ball1.xSpeed - Ball2.xSpeed) * Sin(Angle(Ball1, Ball2)) * Cos(Angle(Ball1, Ball2)) + Ball2.ySpeed
    tmpX2 = (Ball1.xSpeed - Ball2.xSpeed) * (Cos(Angle(Ball1, Ball2))) ^ 2 + (Ball1.ySpeed - Ball2.ySpeed) * Sin(Angle(Ball1, Ball2)) * Cos(Angle(Ball1, Ball2)) + Ball2.xSpeed
    tmpY2 = (Ball1.ySpeed - Ball2.ySpeed) * (Sin(Angle(Ball1, Ball2))) ^ 2 + (Ball1.xSpeed - Ball2.xSpeed) * Sin(Angle(Ball1, Ball2)) * Cos(Angle(Ball1, Ball2)) + Ball2.ySpeed
    
    'Now we have the x and y speeds for each ball
    Ball1.xSpeed = tmpX
    Ball1.ySpeed = tmpY
    Ball2.xSpeed = tmpX2
    Ball2.ySpeed = tmpY2
End Sub

Public Function FindDistance(Ball1 As clsBall, Ball2 As clsBall) As Double
    'This routine checks the positon of 2 balls
    'It looks at their x and y position and returns a single value
    'If this value is 0 then the balls are on top of each other,
    'ie. their x and y's are identical
    FindDistance = Sqr((Ball2.x - Ball1.x) * (Ball2.x - Ball1.x) + (Ball2.y - Ball1.y) * (Ball2.y - Ball1.y))
End Function

Private Function Angle(Ball1 As clsBall, Ball2 As clsBall) As Single
    'This function returns the angle between the 2 balls
    If Ball1.x - Ball2.x = 0 Then
        Angle = (Pi / 2)
    Else
        Angle = Atn((Ball1.y - Ball2.y) / (Ball1.x - Ball2.x))
    End If
End Function
