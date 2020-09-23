Attribute VB_Name = "Bas3D"
Global Const Pi As Double = 3.14159265358979
Global Const AStep As Double = ((Pi * 2 / 360))

Public Type Matrix44
 M11 As Double: M12 As Double: M13 As Double: M14 As Double
 M21 As Double: M22 As Double: M23 As Double: M24 As Double
 M31 As Double: M32 As Double: M33 As Double: M34 As Double
 M41 As Double: M42 As Double: M43 As Double: M44 As Double
End Type

Public Type Vector4D
 X As Double
 Y As Double
 Z As Double
 W As Double
End Type

Public Type Camera
 Position As Vector4D
 LookAt As Vector4D
 BoolLockAt As Boolean
 FOV As Double
 Zoom As Double
End Type
Public Function RotateVec(V As Vector4D, Axis As Byte, Ang As Single) As Vector4D

  Select Case Axis

   Case 0:

    RotateVec.X = V.X
    RotateVec.Y = Cos(Ang) * V.Y - Sin(Ang) * V.Z
    RotateVec.Z = Sin(Ang) * V.Y + Cos(Ang) * V.Z
    RotateVec.W = 1

   Case 1:

    RotateVec.X = Cos(Ang) * V.X + Sin(Ang) * V.Z
    RotateVec.Y = V.Y
    RotateVec.Z = -Sin(Ang) * V.X + Cos(Ang) * V.Z
    RotateVec.W = 1

   Case 2:

    RotateVec.X = Cos(Ang) * V.X - Sin(Ang) * V.Y
    RotateVec.Y = Sin(Ang) * V.X + Cos(Ang) * V.Y
    RotateVec.Z = V.Z
    RotateVec.W = 1

  End Select

End Function
Public Function CalculateFOV(Zoom As Double) As Double

 CalculateFOV = ((2 * Atn(1 / Zoom)) * 57.2957795130823)

End Function
Public Function MatrixMultiplyVector(M1 As Matrix44, V1 As Vector4D) As Vector4D

 With MatrixMultiplyVector

  .X = (M1.M11 * V1.X) + (M1.M12 * V1.Y) + (M1.M13 * V1.Z) + (M1.M14 * V1.W)
  .Y = (M1.M21 * V1.X) + (M1.M22 * V1.Y) + (M1.M23 * V1.Z) + (M1.M24 * V1.W)
  .Z = (M1.M31 * V1.X) + (M1.M32 * V1.Y) + (M1.M33 * V1.Z) + (M1.M34 * V1.W)
  .W = (M1.M41 * V1.X) + (M1.M42 * V1.Y) + (M1.M43 * V1.Z) + (M1.M44 * V1.W)

 End With

End Function
Public Function MatrixProjectionView(VPN As Vector4D, VUP As Vector4D, VRP As Vector4D, BLookAt As Boolean) As Matrix44

 Dim RotateVRC As Matrix44, TranslateVRP As Matrix44
 Dim N As Vector4D, U As Vector4D, V As Vector4D

 N = VectorNormalize(VPN)

 U = CrossProduct(VUP, N)
 U = VectorNormalize(U)

 V = CrossProduct(N, U)

 RotateVRC = MatrixIdentity

 With RotateVRC
  .M11 = U.X: .M12 = U.Y: .M13 = U.Z
  .M21 = V.X: .M22 = V.Y: .M23 = V.Z
  .M31 = N.X: .M32 = N.Y: .M33 = N.Z
 End With

 TranslateVRP = MatrixTranslation(-VRP.X, -VRP.Y, -VRP.Z)

 MatrixProjectionView = MatrixIdentity()
 MatrixProjectionView = MatrixMultiply(MatrixProjectionView, TranslateVRP)
 If BLookAt = True Then MatrixProjectionView = MatrixMultiply(MatrixProjectionView, RotateVRC)

End Function
Public Function MatrixTranslation(TranslateX As Double, TranslateY As Double, TranslateZ As Double) As Matrix44

 MatrixTranslation = MatrixIdentity

 With MatrixTranslation

  .M14 = TranslateX
  .M24 = TranslateY
  .M34 = TranslateZ

 End With

End Function
Public Function MatrixMultiply(M1 As Matrix44, M2 As Matrix44) As Matrix44

 Dim M1B As Matrix44
 Dim M2B As Matrix44

 M1B = M1: M2B = M2

 MatrixMultiply = MatrixIdentity

 With MatrixMultiply

  .M11 = (M1B.M11 * M2B.M11) + (M1B.M21 * M2B.M12) + (M1B.M31 * M2B.M13) + (M1B.M41 * M2B.M14)
  .M12 = (M1B.M12 * M2B.M11) + (M1B.M22 * M2B.M12) + (M1B.M32 * M2B.M13) + (M1B.M42 * M2B.M14)
  .M13 = (M1B.M13 * M2B.M11) + (M1B.M23 * M2B.M12) + (M1B.M33 * M2B.M13) + (M1B.M43 * M2B.M14)
  .M14 = (M1B.M14 * M2B.M11) + (M1B.M24 * M2B.M12) + (M1B.M34 * M2B.M13) + (M1B.M44 * M2B.M14)

  .M21 = (M1B.M11 * M2B.M21) + (M1B.M21 * M2B.M22) + (M1B.M31 * M2B.M23) + (M1B.M41 * M2B.M24)
  .M22 = (M1B.M12 * M2B.M21) + (M1B.M22 * M2B.M22) + (M1B.M32 * M2B.M23) + (M1B.M42 * M2B.M24)
  .M23 = (M1B.M13 * M2B.M21) + (M1B.M23 * M2B.M22) + (M1B.M33 * M2B.M23) + (M1B.M43 * M2B.M24)
  .M24 = (M1B.M14 * M2B.M21) + (M1B.M24 * M2B.M22) + (M1B.M34 * M2B.M23) + (M1B.M44 * M2B.M24)

  .M31 = (M1B.M11 * M2B.M31) + (M1B.M21 * M2B.M32) + (M1B.M31 * M2B.M33) + (M1B.M41 * M2B.M34)
  .M32 = (M1B.M12 * M2B.M31) + (M1B.M22 * M2B.M32) + (M1B.M32 * M2B.M33) + (M1B.M42 * M2B.M34)
  .M33 = (M1B.M13 * M2B.M31) + (M1B.M23 * M2B.M32) + (M1B.M33 * M2B.M33) + (M1B.M43 * M2B.M34)
  .M34 = (M1B.M14 * M2B.M31) + (M1B.M24 * M2B.M32) + (M1B.M34 * M2B.M33) + (M1B.M44 * M2B.M34)

  .M41 = (M1B.M11 * M2B.M41) + (M1B.M21 * M2B.M42) + (M1B.M31 * M2B.M43) + (M1B.M41 * M2B.M44)
  .M42 = (M1B.M12 * M2B.M41) + (M1B.M22 * M2B.M42) + (M1B.M32 * M2B.M43) + (M1B.M42 * M2B.M44)
  .M43 = (M1B.M13 * M2B.M41) + (M1B.M23 * M2B.M42) + (M1B.M33 * M2B.M43) + (M1B.M43 * M2B.M44)
  .M44 = (M1B.M14 * M2B.M41) + (M1B.M24 * M2B.M42) + (M1B.M34 * M2B.M43) + (M1B.M44 * M2B.M44)

 End With

End Function
Public Function MatrixIdentity() As Matrix44

 With MatrixIdentity

  .M11 = 1: .M12 = 0: .M13 = 0: .M14 = 0
  .M21 = 0: .M22 = 1: .M23 = 0: .M24 = 0
  .M31 = 0: .M32 = 0: .M33 = 1: .M34 = 0
  .M41 = 0: .M42 = 0: .M43 = 0: .M44 = 1

 End With

End Function
Public Function CrossProduct(V1 As Vector4D, V2 As Vector4D) As Vector4D

 CrossProduct.X = (V1.Y * V2.Z) - (V1.Z * V2.Y)
 CrossProduct.Y = (V1.Z * V2.X) - (V1.X * V2.Z)
 CrossProduct.Z = (V1.X * V2.Y) - (V1.Y * V2.X)

End Function
Public Function DotProduct(V1 As Vector4D, V2 As Vector4D) As Double

 DotProduct = (V1.X * V2.X) + (V1.Y * V2.Y) + (V1.Z * V2.Z) + (V1.W * V2.W)

End Function
Public Function VectorNormalize(V As Vector4D) As Vector4D

 Dim Length As Double

 Length = Sqr((V.X ^ 2) + (V.Y ^ 2) + (V.Z ^ 2))

 If Length = 0 Then Length = 1

 VectorNormalize.X = (V.X / Length)
 VectorNormalize.Y = (V.Y / Length)
 VectorNormalize.Z = (V.Z / Length)

End Function
Public Function VectorSub(V1 As Vector4D, V2 As Vector4D) As Vector4D

 VectorSub.X = (V1.X - V2.X)
 VectorSub.Y = (V1.Y - V2.Y)
 VectorSub.Z = (V1.Z - V2.Z)

End Function
