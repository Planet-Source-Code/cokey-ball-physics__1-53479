VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Balls"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7680
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   7680
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Colin Woor
'Email: colin.woor@gmsl.co.uk
'Description:
'This program is a demonstration of the way balls collide and bounce off
'each other, it uses classes and could be improved a lot ie. bit blitting sprites
'instead of using shapes.
'My physics sucks, BIG style, so I wrote this to try and remember
'all that stuff i forgot! :D So please excuse any terms/maths that are not correct.
'This program is not perfect and sometimes the balls stick to each other
'which does make a nice effect! :P
'Feel free to contact me for any reason regarding this program,
'I will be happy to help anyone.

'Also feel free to change, use and abuse this program, I am happy to help :D


Private Const Xmin = 500            'Min for the X
Private Const Xmax = 9000           'Max for the X
Private Const Ymin = 500            'Min for the Y
Private Const Ymax = 7000           'Min for the Y
Private Const BallNum = 5           'Number Of Balls
Private BallCol As Collection       'Collection Of Ball Class
Private Const CircleRadius = 450    'Size of circle, change this for bigger balls
Private Const BallColor = 255       'Set ball color to 255(Red)
Private Const Turbo = 150           'Add a little zip into the balls ;)
Private Const InertiaRate = 0.998   'This is the rate at which the balls slow down
                                    'ie. the effect of inertia/mass on the balls
                                    'Make this number smaller to increase mass/inertia of/on the balls
                                    'which will slow the balls down much quicker
                                    'Increase the number and the balls will slow down much slower
                                    'Play with this number to see the effect

Private Sub RunBalls()
Dim i As Long, X As Long    'Looping variables
    
    Do
        Form1.Cls 'Clear the screen
        'Loop through all the balls in the collection
        DrawBoundary    'Draw a box around the defined boundary
        For i = 1 To BallNum
        
            'Set the fill color up
            Form1.FillColor = BallCol(i).BallColor  'Fill in the balls
            
            'Draw the ball onto the screen
            Form1.Circle (BallCol(i).X, BallCol(i).Y), BallCol(i).BallRadius, 0
            
            'Draw the name on the ball
            Form1.CurrentX = Form1.CurrentX - 200
            Form1.CurrentY = Form1.CurrentY - 60
            Form1.Print BallCol(i).BallName
            
            'Move the ball by adding the xspeed/yspeed to the x & y variables
            BallCol(i).X = BallCol(i).X + BallCol(i).xSpeed
            BallCol(i).Y = BallCol(i).Y + BallCol(i).ySpeed
            
            'Slow the balls down, inflict intertia
            If BallCol(i).xSpeed <> 0 Then
                BallCol(i).xSpeed = BallCol(i).xSpeed * BallCol(i).Inertia
            End If
            If Abs(BallCol(i).xSpeed) <= 0.0001953125 Then
                'The above number is kind of the stop speed, change this for
                'slight adjustments
                BallCol(i).xSpeed = 0
            End If
            If BallCol(i).ySpeed <> 0 Then
                BallCol(i).ySpeed = BallCol(i).ySpeed * BallCol(i).Inertia
            End If
            If Abs(BallCol(i).ySpeed) <= 0.0001953125 Then
                'The above number is kind of the stop speed, change this for
                'slight adjustments
                BallCol(i).ySpeed = 0
            End If
            
            'These lines increase the inertia on the ball, this gives
            'it a more realistic slow down (deceleration)
            If BallCol(i).xSpeed = 0 And BallCol(i).ySpeed = 0 Then
                BallCol(i).Inertia = InertiaRate
            Else
                BallCol(i).Inertia = BallCol(i).Inertia * BallCol(i).Mass
            End If
            
            
            'Check the bounderies to make sure that the ball has
            'not gone outside of the defined area ie. XMin/XMax etc.
            If BallCol(i).X <= Xmin Or BallCol(i).X >= Xmax Then
                BallCol(i).xSpeed = -BallCol(i).xSpeed
                BallCol(i).X = IIf(BallCol(i).X <= Xmin, Xmin, Xmax)
            End If
            If BallCol(i).Y <= Ymin Or BallCol(i).Y >= Ymax Then
                BallCol(i).ySpeed = -BallCol(i).ySpeed
                BallCol(i).Y = IIf(BallCol(i).Y <= Ymin, Ymin, Ymax)
            End If
            
            'Check for collision's
            For X = 1 To BallNum
                'Set up another loop through the ball collection
                If Not X = i Then   'If x=i then we have the same ball a i, and we dont
                                    'need to check a collision with ourselves! :)
                    If FindDistance(BallCol(X), BallCol(i)) < (BallCol(X).BallRadius * 2) Then
                        'If the distance we get back from the above function
                        'is the same as the circumference of the ball then the balls have hit
                        
                        'The 2 balls are close enough for a collision to occur
                        'so call the collide routine to bounce them
                        Collide BallCol(X), BallCol(i)
                    End If
                End If
            Next X
        Next i
        DoEvents
    Loop
    
End Sub

Private Sub Form_Load()
    
    Me.Show
    Randomize                       'Reseed the Rnd() engine for better randomnous ;)
    Form1.Width = Xmax + 800        'Set up the form
    Form1.Height = Ymax + 1000      ' ' '
    Form1.FillStyle = 0             'Set this to 0 to create solid circles
    Form1.DrawWidth = 1             'Change the line width
    Set BallCol = New Collection    'Instantiate(create) our collection
    SetUpBalls          'Create the balls and set them up with position's and speed
    RunBalls            'Start it off
End Sub

Private Sub SetUpBalls()
Dim i As Long           'Looping Variable
Dim tmpBall As clsBall  'Our temp ball class
    
    'This routine add 'BallNum' of balls to the BallCol collection
    
    For i = 1 To BallNum
        'Create a new ball class
        Set tmpBall = New clsBall
        With tmpBall
            'To choose a random number within a defined limit
            'you use this: Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
            
            'Set the x & y coord's so it is within our defined Xmin/Xmax/Ymin/Ymax area
            'Do this by adding/subtracting the circle radius to Xmax/Ymax
            .X = Int((Xmax - ((CircleRadius + 20) * 2) - (Xmin + ((CircleRadius + 20) * 2)) + 1) * Rnd + (Xmin + ((CircleRadius + 20) * 2)))
            .Y = Int((Ymax - ((CircleRadius + 20) * 2) - (Ymin + ((CircleRadius + 20) * 2)) + 1) * Rnd + (Ymin + ((CircleRadius + 20) * 2)))
            
            'Set the xSpeed and the ySpeed values
            .xSpeed = Int((20 - (-20) + 1) * Rnd + -20)
            .xSpeed = IIf(.xSpeed < 0, xSpeed - Turbo, xSpeed + Turbo)
            .ySpeed = Int((20 - (-20) + 1) * Rnd + -20)
            .ySpeed = IIf(.ySpeed < 0, ySpeed - Turbo, ySpeed + Turbo)
            
            'Set the ball color up to be red
            'If you want random colors then use :-
            '.BallColor = RGB(Int((255 - 1 + 1) * Rnd + 1), Int((255 - 1 + 1) * Rnd + 1), Int((255 - 1 + 1) * Rnd + 1))
            .BallColor = RGB(255, 0, 0)
            
            'Set up the ball radius
            .BallRadius = CircleRadius
            
            'Set up the ball own intertia rate
            .Inertia = InertiaRate
            
            'Set up the balls mass, this affects how quickly the ball slows down
            .Mass = 0.99999     'This number is fairly good, change it and see :D
            'You could use this for random mass for
            'each of the balls :- Round((0.99999 - 0.9999) * Rnd + 0.9999, 5)
    
            'Set the ballname to be the value of i
            .BallName = i
        End With
        
        'Add the ball class to the ball collection
        BallCol.Add tmpBall
    Next i
        
End Sub

Private Sub DrawBoundary()
    'Draw a box around the defined area
    'Add the circle radius to the edges, so the ball doesnt
    'appear to go over the lines
    
    ' | Left Side
    Form1.Line ((Xmin - CircleRadius), Ymin - (CircleRadius))-((Xmin - CircleRadius), Ymax + (CircleRadius)), RGB(0, 0, 255)
    ' - Bottom
    Form1.Line (Xmin - (CircleRadius), (Ymax + CircleRadius))-(Xmax + (CircleRadius), (Ymax + CircleRadius)), RGB(0, 0, 255)
    ' | Right Side
    Form1.Line ((Xmax + CircleRadius), Ymax + (CircleRadius))-((Xmax + CircleRadius), Ymin - (CircleRadius)), RGB(0, 0, 255)
    ' - Top
    Form1.Line (Xmax + (CircleRadius), (Ymin - CircleRadius))-(Xmin - (CircleRadius), (Ymin - CircleRadius)), RGB(0, 0, 255)
End Sub
