Attribute VB_Name = "mdlRndTxt"
Option Explicit


' Module to create a random text
'
' 1.0.0
' 2021-11-06
' Created RndAscii


Const VERSION = "1.0.0"


' Create a random string with the specified length
' This function can be used for creating password
' Default = 50 is a reasonable length in my opinion.
'
' -----------------------------------
'     Ascii Table
' -----------------------------------
' Decimal      Hexadecimal   Character
' 32           20
' 33           21            !
' 48           30            0
' 57           39            9
' 65           41            A
' 90           5A            Z
' 97           61            a
' 122          7A            z
' 126          7E            ~
' -----------------------------------
'
' As 0x20 is a space, it is not good as a password.
'
Public Function RndAscii( _
    Optional ByVal CharLen As Integer = 50, _
    Optional ByVal AsciiMin As Integer = &H21, _
    Optional ByVal AsciiMax As Integer = &H7E _
    ) As String
    
    Dim r As String
    Dim i As Integer
    
    ' accumulate random ascii letters
    For i = 0 To CharLen
        r = r & Chr(WorksheetFunction.RandBetween(AsciiMin, AsciiMax))
    Next i
    
    ' return
    RndAscii = r
    
End Function

' clear immediate window
Private Sub ClearImmediateWindow()
    Application.SendKeys "^g ^a {DEL}"
End Sub


' print ascii letters to immediate window
Private Sub PrintAscii()
    Dim i As Integer
    For i = &H20 To &H7E
        Debug.Print i, Hex(i), Chr(i)
    Next i
End Sub





