VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GSubclass"
Attribute VB_GlobalNameSpace = True
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

Sub AttachMessage(iwp As ISubclass, ByVal hWnd As Long, _
                  ByVal iMsg As Long)
Attribute AttachMessage.VB_Description = "Start subclassing for a particular message."
    MSubclass.AttachMessage iwp, hWnd, iMsg
End Sub

Sub DetachMessage(iwp As ISubclass, ByVal hWnd As Long, _
                  ByVal iMsg As Long)
Attribute DetachMessage.VB_Description = "Stop subclassing for a particular message."
    MSubclass.DetachMessage iwp, hWnd, iMsg
End Sub

Public Property Get CurrentMessage() As Long
Attribute CurrentMessage.VB_Description = "Returns the message currently being worked on.  Only valid in the MsgResponse and WindowProc items of your implemented interface."
   CurrentMessage = MSubclass.CurrentMessage
End Property
Public Function CallOldWindowProc( _
      ByVal hWnd As Long, _
      ByVal iMsg As Long, _
      ByVal wParam As Long, _
      ByVal lParam As Long _
   ) As Long
Attribute CallOldWindowProc.VB_Description = "Calls the window procedure for this handle which was installed before you added the subclass."
   CallOldWindowProc = MSubclass.CallOldWindowProc(hWnd, iMsg, wParam, lParam)
End Function

