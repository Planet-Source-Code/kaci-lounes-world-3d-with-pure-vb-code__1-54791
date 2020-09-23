VERSION 5.00
Begin VB.Form Main 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5580
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PointsList() As Vector4D
Dim TmpPoints() As Vector4D

Dim ClipDistanceFar As Double
Dim ClipDistanceNear As Double

Dim DrawFlag As Boolean

Dim MyCamera As Camera
Dim AniCamera As Boolean

'############# Déclarations internes #############
Dim Brightness As Integer, DeltaVisible As Double
Dim Indx As Long, PixX As Double, PixY As Double, Distance As Double
Dim View As Matrix44, VecVPN As Vector4D, VecVUP As Vector4D, VecVRP As Vector4D
Dim I As Integer, J As Integer
Sub DrawPoints()

 For Indx = LBound(TmpPoints) To UBound(TmpPoints)

  PixX = TmpPoints(Indx).X
  PixY = -TmpPoints(Indx).Y
  Distance = TmpPoints(Indx).Z

  If (Distance > ClipDistanceNear) And (Distance < ClipDistanceFar) Then
   If (Abs(PixX) < (1 / MyCamera.Zoom)) And (Abs(PixY) < (1 / MyCamera.Zoom)) Then

    If DrawFlag = True Then
     Brightness = 255 - CInt((Distance / DeltaVisible) * 255)
     Me.PSet (PixX, PixY), RGB(Brightness, Brightness, Brightness)
    Else
     Me.PSet (PixX, PixY), vbWhite
    End If

   End If
  End If

 Next Indx

End Sub
Sub SetPoints()

'Formules Géométrique 3D:

'Goutte à goutte 1: Z = Sin(0.075 * (XX ^ 2 + YY ^ 2))
'Goutte à goutte 2: Z = 12 * Cos((XX ^ 2 + YY ^ 2) / 4) / (1 + XX ^ 2 + YY ^ 2) - 4 / (XX ^ 2 + YY ^ 2)
'Goutte à goutte 3: Z = 12 * Cos((XX ^ 2 + YY ^ 2) / 4) / (3 + XX ^ 2 + YY ^ 2)
'Cratère: Z = Cos(0.05 * (XX ^ 2 + YY ^ 2))
'Eruption: Z = 16.5 * Exp(-0.05 * (XX ^ 2 + YY ^ 2))
'Météorite: Z = Sin(7 * Exp(-0.02 * (XX ^ 2 + YY ^ 2)))

 ReDim PointsList(0): J = 0

 For YY = -29 To 29
  For XX = -29 To 29
   ReDim Preserve PointsList(UBound(PointsList) + 1)
   PointsList(J).X = XX
   PointsList(J).Y = YY
   PointsList(J).Z = Sin(20 * Exp(-0.002 * (XX ^ 2 + YY ^ 2))) 'Formule
   PointsList(J).W = 1
   J = J + 1
  Next XX
 Next YY

'Dessine les trois lignes pour les trois axe X,Y et Z.

 For I = -200 To 199 Step 10
  ReDim Preserve PointsList(UBound(PointsList) + 1)
  PointsList(J).X = I
  PointsList(J).W = 1
  J = J + 1
 Next I

 For I = -200 To 199 Step 10
  ReDim Preserve PointsList(UBound(PointsList) + 1)
  PointsList(J).Y = I
  PointsList(J).W = 1
  J = J + 1
 Next I

 For I = -200 To 199 Step 10
  ReDim Preserve PointsList(UBound(PointsList) + 1)
  PointsList(J).Z = I
  PointsList(J).W = 1
  J = J + 1
 Next I

 DeltaVisible = (ClipDistanceFar - ClipDistanceNear)

End Sub
Sub SetPointsPositions()

 ReDim TmpPoints(UBound(PointsList))

 VecVPN = VectorSub(MyCamera.LookAt, MyCamera.Position)

 With VecVUP
  .X = 0: .Y = 1: .Z = 0: .W = 1
 End With

 With VecVRP
  .X = MyCamera.Position.X
  .Y = MyCamera.Position.Y
  .Z = MyCamera.Position.Z
  .W = 1
 End With

 View = MatrixProjectionView(VecVPN, VecVUP, VecVRP, MyCamera.BoolLockAt)

 For Indx = LBound(PointsList) To UBound(PointsList)

  TmpPoints(Indx) = MatrixMultiplyVector(View, PointsList(Indx))

  If TmpPoints(Indx).Z <> 0 Then
   TmpPoints(Indx).X = TmpPoints(Indx).X / TmpPoints(Indx).Z
   TmpPoints(Indx).Y = TmpPoints(Indx).Y / TmpPoints(Indx).Z
  End If

 Next Indx

End Sub
Private Sub Form_Activate()

 Do

  Cls

  If AniCamera = True Then MyCamera.Position = RotateVec(MyCamera.Position, 0, (3.14159 / 180))

  SetPointsPositions
  DrawPoints

  DoEvents

 Loop

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

 Select Case KeyCode

  Case vbKeyEscape: Unload Me: End
  Case vbKeySpace: Call Form_Load

  Case vbKeyDelete: MyCamera.BoolLockAt = Not MyCamera.BoolLockAt
  Case vbKeyInsert: DrawFlag = Not DrawFlag
  Case vbKeyA: AniCamera = Not AniCamera

  Case vbKeyHome:
                 MyCamera.Zoom = MyCamera.Zoom + 0.05
                 ScaleWidth = 2 * (1 / MyCamera.Zoom)
                 ScaleLeft = -(1 / MyCamera.Zoom)
                 ScaleHeight = Me.ScaleWidth
                 ScaleTop = Me.ScaleLeft
  Case vbKeyEnd:
                 MyCamera.Zoom = MyCamera.Zoom - 0.05
                 ScaleWidth = 2 * (1 / MyCamera.Zoom)
                 ScaleLeft = -(1 / MyCamera.Zoom)
                 ScaleHeight = Me.ScaleWidth
                 ScaleTop = Me.ScaleLeft

  Case vbKeyLeft: If Shift = 0 Then MyCamera.Position.X = MyCamera.Position.X + 1 Else MyCamera.LookAt.X = MyCamera.LookAt.X + 1
  Case vbKeyRight: If Shift = 0 Then MyCamera.Position.X = MyCamera.Position.X - 1 Else MyCamera.LookAt.X = MyCamera.LookAt.X - 1

  Case vbKeyUp: If Shift = 0 Then MyCamera.Position.Y = MyCamera.Position.Y + 1 Else MyCamera.LookAt.Y = MyCamera.LookAt.Y + 1
  Case vbKeyDown: If Shift = 0 Then MyCamera.Position.Y = MyCamera.Position.Y - 1 Else MyCamera.LookAt.Y = MyCamera.LookAt.Y - 1

  Case vbKeyPageUp: If Shift = 0 Then MyCamera.Position.Z = MyCamera.Position.Z + 1 Else MyCamera.LookAt.Z = MyCamera.LookAt.Z + 1
  Case vbKeyPageDown: If Shift = 0 Then MyCamera.Position.Z = MyCamera.Position.Z - 1 Else MyCamera.LookAt.Z = MyCamera.LookAt.Z - 1

 End Select

End Sub
Private Sub Form_Load()

 Move 0, 0, (640 * 15), (480 * 15)
 AutoRedraw = True
 BackColor = vbBlack
 ScaleMode = vbPixels

 ClipDistanceFar = 200
 ClipDistanceNear = 0

 SetPoints

 With MyCamera
  .BoolLockAt = True
  .Position.X = 50
  .Position.Y = 50
  .Position.Z = 50
  .LookAt.X = 0
  .LookAt.Y = 0
  .LookAt.Z = 0
  .Zoom = 1
  .FOV = CalculateFOV(.Zoom)
 End With

 ScaleWidth = 2 * (1 / MyCamera.Zoom)
 ScaleLeft = -(1 / MyCamera.Zoom)
 ScaleHeight = Me.ScaleWidth
 ScaleTop = Me.ScaleLeft

 MsgBox "Lock at Form_KeyDown to view controls." & vbNewLine & "          Kaci Lounes  05-2004", vbInformation

End Sub
