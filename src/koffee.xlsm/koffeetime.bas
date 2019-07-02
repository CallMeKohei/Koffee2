Attribute VB_Name = "koffeetime"
''' koffeeTime.bas
''' written by callmekohei(twitter at callmekohei)
''' MIT license
Option Explicit
Option Compare Text
Option Private Module
Option Base 0

#If VBA7 Then
    Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef frequency As Double) As LongPtr
    Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef procTime As Double) As LongPtr
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
#Else
    Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef frequency As Double) As Long
    Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef procTime As Double) As Long
    Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
#End If

Private Sub investigate_Time()
    Dim startTime: startTime = Timer
    ''' do something...
    Debug.Print Format$(Timer - startTime, "0.00") & " seconds."
End Sub

''' https://docs.microsoft.com/ja-jp/previous-versions/office/developer/office-2007/aa730921(v=office.12)
Public Function MilliSecondsTimer() As Double
    MilliSecondsTimer = 0
    Dim ticks As Double:     QueryPerformanceCounter ticks
    Dim frequency As Double: QueryPerformanceFrequency frequency
    If frequency Then MilliSecondsTimer = (ticks / frequency) * 1000
End Function

Public Sub Wait(ByVal milliSeconds As Long)
    Sleep milliSeconds
End Sub
