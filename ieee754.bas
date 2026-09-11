Attribute VB_Name = "IEEE754Module"
'This module contains this program's core procedures.
Option Explicit

Private Const EXPONENT_MASK As Long = &H7F800000   'Defines the exponent's bitmask.
Private Const EXPONENT_SHIFT As Long = &H800000    'Defines the divisor used to right-shift the exponent's bits.
Private Const MANTISSA_MASK As Long = &H7FFFFF     'Defines the mantissa's bitmask.
Private Const SIGN_MASK As Long = &H80000000       'Defines the sign's bitmask.

'The Microsoft Windows API constants used by this program.
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

'This procedure is executed when this program is started.
Public Sub Main()
Dim Bytes As Long
Dim Exponent As Long
Dim Float As Single
Dim Mantissa As Long
Dim Reconstruction As Single
Dim Sign As Long
Dim TotalMantissa As Double

   Debug.Print String$(50, "=")

   Float = 0
   Debug.Print "Float: "; Float
   
   RtlMoveMemory Bytes, Float, Length:=4
   Debug.Print "Bytes: "; Hex$(Bytes); "h"

   If (Bytes And SIGN_MASK) = SIGN_MASK Then Sign = -1 Else Sign = 1
   Debug.Print "Sign: "; Sign

   Exponent = ((Bytes And EXPONENT_MASK) \ EXPONENT_SHIFT)
   Debug.Print "Exponent: "; Exponent; " (Real exponent: " & (Exponent - 127) & ")"

   Mantissa = Bytes And MANTISSA_MASK
   Debug.Print "Mantissa: "; Mantissa
   
   TotalMantissa = 1# + (Mantissa / 8388608#)
   Debug.Print "Total Mantissa: "; TotalMantissa

   Reconstruction = Sign * TotalMantissa * (2 ^ (Exponent - 127))

   Debug.Print "Reconstructed: "; Reconstruction
End Sub

