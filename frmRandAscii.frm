VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRandAscii 
   Caption         =   "Random Ascii"
   ClientHeight    =   4575
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9135.001
   OleObjectBlob   =   "frmRandAscii.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRandAscii"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


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
Public Function RandAscii( _
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
    RandAscii = r
    
End Function

Private Sub btnCreateRndAsc_Click()
    Dim CharLen As Integer, AsciiMin As Integer, AsciiMax As Integer
    Dim OutAscii As String
    CharLen = CInt(textboxLength)
    AsciiMin = CInt(textboxMin)
    AsciiMax = CInt(textboxMax)
    OutAscii = RandAscii(CharLen, AsciiMin, AsciiMax)
    textboxRndAsc.Value = OutAscii
End Sub


