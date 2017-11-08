Option Strict On
Option Explicit On
Option Infer Off
Imports System.Timers
Public Structure FastPoint
    Public Enum FastPointComparisom
        BothFalse = 0
        OneTrue = 1
        BothTrue = 2
    End Enum
    Public X, Y As Integer
    Public Shared Operator =(A As FastPoint, B As FastPoint) As Boolean
        If A.X = B.X AndAlso A.Y = B.Y Then
            Return True
        End If
        Return False
    End Operator
    Public Shared Operator <>(A As FastPoint, B As FastPoint) As Boolean
        If A.X <> B.X OrElse A.Y <> B.Y Then
            Return True
        End If
        Return False
    End Operator
    Public Shared Operator <(A As FastPoint, B As FastPoint) As FastPointComparisom
        Return CInt(A.X < B.X) + CInt(A.Y < B.Y)
        If A.X < B.X OrElse A.Y < B.Y Then
            Return True
        End If
        Return False
    End Operator
    Public Shared Operator <=(A As FastPoint, B As FastPoint) As Boolean
        If A.X <= B.X AndAlso A.Y <= B.Y Then
            Return True
        End If
        Return False
    End Operator
    Public Shared Operator >=(A As FastPoint, B As FastPoint) As Boolean
        If A.X >= B.X AndAlso A.Y >= B.Y Then
            Return True
        End If
        Return False
    End Operator
    Public Shared Operator >(A As FastPoint, B As FastPoint) As Boolean
        If A.X > B.X OrElse A.Y > B.Y Then
            Return True
        End If
        Return False
    End Operator
    Public Shared Operator Like(A As FastPoint, B As FastPoint) As Boolean
        If A.X < B.X AndAlso A.Y < B.Y Then
            Return True
        End If
        Return False
    End Operator
    Public Sub New(ByVal X As Integer, ByVal Y As Integer)
        With Me
            .X = X
            .Y = Y
        End With
    End Sub
    Public Sub New(ByVal Point As Point)
        With Point
            X = .X
            Y = .Y
        End With
    End Sub
    Public Sub Add(ByVal Value As FastPoint)
        With Value
            X += .X
            Y += .Y
        End With
    End Sub
    Public Sub Subtract(ByVal Value As FastPoint)
        With Value
            X -= .X
            Y -= .Y
        End With
    End Sub
    Public Function NewAdd(ByVal Value As FastPoint) As FastPoint
        With Value
            Return New FastPoint(X + .X, Y + .Y)
        End With
    End Function
    Public Function NewSubtract(ByVal Value As FastPoint) As FastPoint
        With Value
            Return New FastPoint(X - .X, Y - .Y)
        End With
    End Function
    Public Function ToPoint() As Point
        Return New Point(X, Y)
    End Function
End Structure


Public Class Univers
    Inherits PictureBox
    Protected Friend Shared Verdier As List(Of Hastighet)

    Private WithEvents Pi As System.Timers.Timer
    ' Private WithEvents testtimer1, testtimer2 As New System.Windows.Forms.Timer

    Private PiThreaded As New Threading.Timer(AddressOf TimerCallback, Ready, 1000, 16)

    Private SC As Threading.SynchronizationContext = Threading.SynchronizationContext.Current
    'Private Time As Integer = 0
    'Private Center As Point
    'Private M As Integer
    Private CollisionBrush As New SolidBrush(Color.Red)
    Private PointBrush As New SolidBrush(Color.Yellow)
    Private Rnd As New Random(Date.Now.Millisecond)
    Dim Ready As Boolean = True
    Private Sub TimerCallback(State As Object)
        If DirectCast(State, Boolean) Then
            SC.Post(AddressOf Tick, Nothing)
        End If
    End Sub
    Public Sub New(Parent As Control, Optional ByVal Interval As Integer = 50, Optional ByVal Width As Integer = 300, Optional ByVal Height As Integer = 300)
        'Dim NyVerdier(Mass - 1) As Hastighet
        'Verdier = NyVerdier
        SuspendLayout()
        Me.Parent = Parent
        With Parent.ClientSize
            Hastighet.Plane = New Plane(.Width, .Height)
        End With
        If Verdier Is Nothing Then
            Verdier = New List(Of Hastighet)(2000)
        End If
        Hide()
        'M = Particles - 1
        'Dim rnd As New Random(Date.Now.Millisecond)
        With Me
            .DoubleBuffered = True
            .Width = Width
            .Height = Height
            .BackColor = Color.Blue
        End With


        'testtimer1.Interval = 16
        'testtimer1.Start()
        'testtimer2.Interval = 16
        'testtimer2.Start()

        Pi = New Timers.Timer(30)
        With Pi
            .AutoReset = False
            .Start()
        End With
        'PiThreaded.Change(1000, 16)
        Show()
        ResumeLayout()
    End Sub
    Public Sub Air(Position As FastPoint, Intensity As Double, ByVal Radius As Integer)
        Dim iLast As Integer = Verdier.Count - 1
        For i As Integer = 0 To iLast
            With Verdier(i)
                Dim DiffXY As FastPoint = Position.NewSubtract(.Position)
                Dim Distance As Double = Math.Sqrt(DiffXY.X ^ 2 + DiffXY.Y ^ 2)
                If Distance < Radius Then
                    Dim VelXY As New FastPoint(0, 0)
                    Dim AngleInRadians As Double = Math.Atan2(DiffXY.X, DiffXY.Y)
                    Dim F As Double = 10 * (Radius - Distance)
                    Dim Fx As Integer = CInt(F * Math.Cos(AngleInRadians))
                    Dim Fy As Integer = CInt(F * Math.Sin(AngleInRadians))
                    VelXY.X = Fx
                    VelXY.Y = Fy

                    If (Fx < 0) <> (DiffXY.X < 0) Then
                        Fx = -Fx
                    End If
                    If (Fy < 0) <> (DiffXY.Y < 0) Then
                        Fy = -Fy
                    End If
                    If DiffXY.X > 0 AndAlso DiffXY.Y < 0 Then
                        .Velocity.Add(VelXY)
                    ElseIf DiffXY.X < 0 AndAlso DiffXY.Y < 0 Then
                        .Velocity.Subtract(VelXY)
                    End If
                End If
            End With
        Next
    End Sub
    Public Sub Clear()
        Dim ilast As Integer = Verdier.Count - 1
        For i As Integer = ilast To 0 Step -1
            With Verdier
                With .Item(i)
                    Hastighet.Plane.Subtract(.Position)
                End With
                .RemoveAt(i)
            End With
        Next
    End Sub
    'Public Function SumValues() As Point
    '    Dim SumPoint As New Point(0, 0)
    '    Dim Errors As Integer = 0
    '    With Hastighet.Plane.Velocities
    '        For i As Integer = 0 To .GetLength(0) - 1
    '            For n As Integer = 0 To .GetLength(1) - 1
    '                With Hastighet.Plane.Velocities(i, n)
    '                    If .X <> 0 OrElse .Y <> 0 Then
    '                        Errors += 1
    '                    End If
    '                End With
    '                SumPoint.Offset(Hastighet.Plane.Velocities(i, n))
    '            Next
    '        Next
    '    End With
    '    Return SumPoint
    'End Function
    Protected Overrides Sub OnMouseMove(e As MouseEventArgs)
        MyBase.OnMouseMove(e)
        With e
            Dim PlaneLimit As FastPoint = Hastighet.Plane.Limit
            If .Button = MouseButtons.Left Then
                With .Location
                    If .Y > 0 AndAlso .Y < PlaneLimit.Y AndAlso .X > 0 AndAlso .X < PlaneLimit.X Then
                        AddParticles(5, New FastPoint(e.Location))
                    End If
                End With
            End If
        End With
    End Sub
    Public Sub AddParticles(ByVal Amount As Integer, Optional ByVal Position As FastPoint = Nothing)
        Amount -= 1
        With Verdier
            For i As Integer = 0 To Amount
                'Dim RndPoint As Point
                .Add(New Hastighet(Position, New FastPoint(Rnd.Next(-360, 360), Rnd.Next(-360, 360))))
            Next
        End With
        'Refresh()
    End Sub
    'Protected Overrides Sub OnSizeChanged(e As EventArgs)
    '    MyBase.OnSizeChanged(e)
    'End Sub
    Protected Overrides Sub OnClientSizeChanged(e As EventArgs)
        MyBase.OnClientSizeChanged(e)
        With ClientSize
            'Center = New Point(.Width \ 2, .Height \ 2)
            If Hastighet.Plane IsNot Nothing Then
                Hastighet.Plane.Limit = New FastPoint(.Width, .Height)
            End If
        End With
    End Sub
    Private Sub Tick(State As Object)
        If Ready Then
            Ready = False
            With Verdier
                Dim iLast As Integer = .Count - 1
                For i As Integer = iLast To 0 Step -1
                    With Verdier(i)
                        If Not .Remove Then
                            .Time()
                        Else
                            Hastighet.Plane.Subtract(.Position)
                            Verdier.RemoveAt(i)
                        End If
                    End With
                Next
            End With
            'Time += 1
            Refresh()
            Ready = True
        End If
        Pi.Start()
    End Sub

    'Private Sub FormsTimerTick() Handles testtimer1.Tick
    '    Tick(Nothing)
    'End Sub
    'Private Sub FormsTimerTick2() Handles testtimer2.Tick
    '    Tick(Nothing)
    'End Sub


    'Private Sub OnTime(sender As Object, e As ElapsedEventArgs) Handles Pi.Elapsed
    '    SC.Post(AddressOf Tick, Nothing)
    'End Sub
    Private CollisionSize As New Size(3, 3)
    Private RegularSize As New Size(2, 2)
    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e)
        Dim iLast As Integer = Verdier.Count - 1
        For i As Integer = 0 To iLast
            With Verdier(i)
                Dim Pos As Point = .Position.ToPoint
                Dim Col As Boolean = .Collision
                With e.Graphics
                    If Col Then
                        .FillRectangle(CollisionBrush, New Rectangle(Pos, CollisionSize))
                    Else
                        .FillRectangle(PointBrush, New Rectangle(Pos, RegularSize))
                    End If
                End With
            End With
        Next
    End Sub
End Class
Public Class Plane
    'Protected Friend Velocities(,) As Point
    Protected Friend Occupying(,) As Integer
    'Protected Friend CenterX, CenterY As Integer
    Protected Friend Shared BounceParticlesOnEdge As Boolean = True
    Private LimXY As FastPoint
    'Private EqualCounter, UnequalCounter As Integer
    Public Sub Add(ByVal XY As FastPoint)
        With XY
            Occupying(.X, .Y) += 1
        End With
    End Sub
    Public Sub Subtract(ByVal XY As FastPoint)
        With XY
            Occupying(.X, .Y) -= 1
        End With
    End Sub
    Public Sub ResetPositions()
        Occupying = Nothing
        With LimXY
            Occupying = New Integer(.X, .Y) {}
        End With
        With Univers.Verdier
            Dim xLast As Integer = .Count - 1
            For x As Integer = xLast To 0 Step -1
                Dim Loc As FastPoint = .Item(x).Position
                With Univers.Verdier
                    If Loc.X <= 0 OrElse Loc.X >= LimXY.X OrElse Loc.Y <= 0 OrElse Loc.Y >= LimXY.Y Then
                        .RemoveAt(x)
                    Else
                        Add(Loc)
                    End If
                End With
            Next
        End With
    End Sub
    Public Property Limit As FastPoint
        Get
            Return LimXY
        End Get
        Set(value As FastPoint)
            LimXY = value
            ResetPositions()
        End Set
    End Property
    Public Sub New(ByVal LimitX As Integer, ByVal LimitY As Integer)
        Occupying = New Integer(LimitX, LimitY) {}
        LimXY = New FastPoint(LimitX, LimitY)
    End Sub
    Public Sub UpdatePosition(ByRef H As Hastighet)
        With H
            Dim Pos As FastPoint = .Position
            If Pos.X < LimXY.X AndAlso Pos.Y < LimXY.Y AndAlso Pos.X > 0 AndAlso Pos.Y > 0 Then
                Add(Pos)
                If Occupying(Pos.X, Pos.Y) > 1 Then
                    .Collision = True
                    If .Lifetime > 100 Then
                        .TimeToLive = .Lifetime + 5
                    End If
                Else
                    .Collision = False
                End If
            Else
                If BounceParticlesOnEdge Then
                    Dim Flip(1) As Boolean
                    With LimXY
                        Select Case Pos.X
                            Case 0, .X
                                Flip(0) = True
                        End Select
                        Select Case Pos.Y
                            Case 0, .Y
                                Flip(1) = True
                        End Select
                    End With
                    .FlipVelocity(Flip(0), Flip(1))
                Else
                    .Remove = True
                End If
            End If
        End With
    End Sub
End Class
Public Class Hastighet
    'Inherits Control
    Public Remove As Boolean
    Public Shared BounceVelocity As Double = 0.8
    Public Shared Gravity As Integer = 2
    Public TimeToLive As Integer = -1
    Public Lifetime As Integer
    'Private D As Integer
    Public Velocity As New FastPoint(0, 0)
    Private ValueXY As New FastPoint(0, 0)
    Protected Friend Shared Plane As Plane
    'Private M As Integer
    Public Collision As Boolean
    'Private XPen As New Pen(Color.Green, 3)
    'Private YPen As New Pen(Color.Blue, 2)
    'Private RBrush As New SolidBrush(Color.White)
    Public Position As New FastPoint(0, 0)
    'Private DrawRect As Rectangle
    'Public ReadOnly Property Mass As Integer
    '    Get
    '        Return M
    '    End Get
    'End Property

    Public Sub AddVelocity(ByVal Velocities As FastPoint)
        'For i As Integer = 0 To D
        With Plane
            .Subtract(Position)
            Velocity = Velocities
            .Add(Position)
        End With
        'Next
    End Sub
    Public Sub FlipVelocity(ByVal FlipX As Boolean, ByVal FlipY As Boolean)
        If FlipX Then
            With Velocity
                .X = CInt(- .X * BounceVelocity)
                With ValueXY
                    .X = CInt(- .X * BounceVelocity)
                End With
                Select Case .X
                    Case > 0
                        Position.X += 1
                    Case < 0
                        Position.X -= 1
                    Case Else
                        Remove = True
                End Select
            End With
        End If
        If FlipY Then
            With Velocity
                .Y = CInt(- .Y * BounceVelocity)
                With ValueXY
                    .Y = CInt(- .Y * BounceVelocity)
                End With
                Select Case .Y
                    Case > 0
                        Position.Y += 1
                    Case < 0
                        Position.Y -= 1
                    Case Else
                        Remove = True
                End Select
            End With
        End If
        Plane.Add(Position)
    End Sub
    Public Sub New(ByVal Location As FastPoint, Optional ByVal VelocityXY As FastPoint = Nothing)
        'If Plane Is Nothing Then
        '    Plane = New Plane(500, 500)
        'End If
        'Dimensions -= 1
        'Dim Velocities(Dimensions) As Integer
        'Dim Values(Dimensions) As Integer
        'Dim Revolutions(Dimensions) As Integer
        'Revolutions(0) = Position(0)
        'Revolutions(1) = Position(1)
        Velocity = VelocityXY
        'ValueXY = Values
        Position = Location
        'D = Dimensions
        Plane.UpdatePosition(Me)
        'DoubleBuffered = True
        'With Me
        '    .Size = New Size(64, 64)
        '    .Location = New Point((xMax - x) * 64, (yMax - y) * 64)
        'End With
        'DrawRect = New Rectangle(New Point(1, 1), New Size(62, 62))
    End Sub
    'Protected Overrides Sub OnPaint(e As PaintEventArgs)
    '    MyBase.OnPaint(e)
    '    'e.Graphics.DrawArc(XPen, DrawRect, 90, ValueXY(0))
    '    'e.Graphics.DrawString(CStr(Rotations(0)), Label.DefaultFont, RBrush, New Point(15, 25))
    '    'e.Graphics.DrawArc(YPen, DrawRect, 90, ValueXY(1))
    '    'e.Graphics.DrawString(CStr(Rotations(1)), Label.DefaultFont, RBrush, New Point(35, 25))
    'End Sub
    Public Sub Time()
        ' I singulariteten skal alle ligge over hverandre, og alle har lysfart (det må en asymmetri til for å utløse eksplosjonen).
        ' Lag en funksjon som lister alle objekter i et spesifikt punkt når en kollisjon oppdages i ticken til universet. Ta deretter å bytt om på
        ' verdiene på den måten flere objekter som kolliderer i et punkt vil få.
        With Plane
            .Subtract(Position)
            'For iDimension As Integer = 0 To D
            With Velocity
                If .Y < 200 Then
                    .Y += Gravity
                End If
                ValueXY.X += .X
                ValueXY.Y += .Y
            End With
            With ValueXY
                Select Case .X
                    Case >= 360
                        .X -= 360
                        Position.X += 1
                    Case <= -360
                        .X += 360
                        Position.X -= 1
                End Select
                Select Case .Y
                    Case >= 360
                        .Y -= 360
                        Position.Y += 1
                    Case <= -360
                        .Y += 360
                        Position.Y -= 1
                End Select
            End With
            'Next
            .UpdatePosition(Me)
        End With
        If Not Lifetime = TimeToLive Then
            Lifetime += 1
        Else
            Remove = True
        End If
    End Sub
End Class
