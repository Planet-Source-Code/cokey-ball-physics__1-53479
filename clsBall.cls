VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mX As Double    'X position of the ball
Private mY As Double    'Y position of the ball
Private mXSpeed As Double   'X speed of the ball
Private mYSpeed As Double   'Y speed of the ball
Private mBallColor As Long 'Colour of the ball
Private mBallName As String
Private mBallRadius As Long
Private mInertia As Double
Private mMass As Double
'This just contains all the usual let's and get's that
'allow the ball class to store and retrieve it's values
Public Property Let BallName(ByVal InBallName As String)
    mBallName = InBallName
End Property
Public Property Get BallName() As String
    BallName = mBallName
End Property
Public Property Let X(ByVal Inx As Long)
    mX = Inx
End Property
Public Property Let Y(ByVal Iny As Long)
    mY = Iny
End Property
Public Property Get X() As Long
    X = mX
End Property
Public Property Get Y() As Long
    Y = mY
End Property
Public Property Let xSpeed(ByVal Inxspeed As Double)
    mXSpeed = Inxspeed
End Property
Public Property Let ySpeed(ByVal InySpeed As Double)
    mYSpeed = InySpeed
End Property
Public Property Get xSpeed() As Double
    xSpeed = mXSpeed
End Property
Public Property Get ySpeed() As Double
    ySpeed = mYSpeed
End Property
Public Property Let BallColor(ByVal InColor As Long)
    mBallColor = InColor
End Property
Public Property Get BallColor() As Long
    BallColor = mBallColor
End Property
Public Property Get BallRadius() As Long
    BallRadius = mBallRadius
End Property
Public Property Let BallRadius(ByVal InRadius As Long)
    mBallRadius = InRadius
End Property
Public Property Let Inertia(ByVal InInertia As Double)
    mInertia = InInertia
End Property
Public Property Get Inertia() As Double
    Inertia = mInertia
End Property
Public Property Let Mass(ByVal InMass As Double)
    mMass = InMass
End Property
Public Property Get Mass() As Double
    Mass = mMass
End Property
