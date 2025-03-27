VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'vbAccelerator Software License
'
'Version 1.0
'
'Copyright (c) 2002 vbAccelerator.com
'
'Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:
'
'    Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer
'    Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.
'    The end-user documentation included with the redistribution, if any, must include the following acknowledgment:
'
'    "This product includes software developed by vbAccelerator (/index.html)."
'
'    Alternately, this acknowledgment may appear in the software itself, if and wherever such third-party acknowledgments normally appear.
'    The names "vbAccelerator" and "vbAccelerator.com" must not be used to endorse or promote products derived from this software without prior written permission. For written permission, please contact vbAccelerator through steve@vbaccelerator.com.
'    Products derived from this software may not be called "vbAccelerator", nor may "vbAccelerator" appear in their name, without prior written permission of vbAccelerator.
'
'THIS SOFTWARE IS PROVIDED "AS IS" AND ANY EXPRESSED OR IMPLIED WARRANTIES, 
'INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY 
'AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL 
'VBACCELERATOR OR ITS CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, 
'INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT 
'NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, 
'DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY 
'OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING 
'NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, 
'EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

Option Explicit
Private iInterval As Long
Private id As Long

' User can attach any Variant data they want to the timer
Public Item As Variant
Attribute Item.VB_VarDescription = "Gets/sets a piece of additional data to store with the timer."

Public Event ThatTime()
Attribute ThatTime.VB_Description = "Raised when the timer fires."

' SubTimer is independent of VBCore, so it hard codes error handling

Public Enum EErrorTimer
    eeBaseTimer = 13650     ' CTimer
    eeTooManyTimers         ' No more than 10 timers allowed per class
    eeCantCreateTimer       ' Can't create system timer
End Enum

Friend Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.EXEName & ".WindowProc"
        Select Case e
        Case eeTooManyTimers
            sText = "No more than 10 timers allowed per class"
        Case eeCantCreateTimer
            sText = "Can't create system timer"
        End Select
        Err.Raise e Or vbObjectError, sSource, sText
    Else
        ' Raise standard Visual Basic error
        Err.Raise e, sSource
    End If
End Sub


Property Get Interval() As Long
Attribute Interval.VB_Description = "Gets/sets the interval at which the timer fires.  Set to zero to stop the timer."
    Interval = iInterval
End Property

' Can't just change interval--you must kill timer and start a new one
Property Let Interval(iIntervalA As Long)
    Dim f As Boolean
    If iIntervalA > 0 Then
        ' Don't mess with it if interval is the same
        If iInterval = iIntervalA Then Exit Property
        ' Must destroy any existing timer to change interval
        If iInterval Then
            f = TimerDestroy(Me)
            Debug.Assert f      ' Shouldn't fail
        End If
        ' Create new timer with new interval
        iInterval = iIntervalA
        If TimerCreate(Me) = False Then ErrRaise eeCantCreateTimer
    Else
        If (iInterval > 0) Then
            iInterval = 0
            f = TimerDestroy(Me)
            Debug.Assert f      ' Shouldn't fail
        End If
    End If
End Property

' Must be public so that Timer object can't terminate while client's ThatTime
' event is being processed--Friend wouldn't prevent this disaster
Public Sub PulseTimer()
Attribute PulseTimer.VB_MemberFlags = "40"
    RaiseEvent ThatTime
End Sub

Friend Property Get TimerID() As Long
    TimerID = id
End Property

Friend Property Let TimerID(idA As Long)
    id = idA
End Property

Private Sub Class_Terminate()
    Interval = 0
End Sub

