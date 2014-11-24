Attribute VB_Name = "IconStuff"
' *************************************************************************
'  Copyright ©1999 Karl E. Peterson
'  All Rights Reserved, http://www.mvps.org/vb
' *************************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code, non-compiled, without prior written consent.
' *************************************************************************
Option Explicit

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, lpIconName As Any) As Long

Private Const GWL_HWNDPARENT = (-8)

Private Const WM_GETICON = &H7F
Private Const WM_SETICON = &H80

Private Const ICON_SMALL = 0
Private Const ICON_BIG = 1

Private Const IDI_WINLOGO = 32517&

Public Function SetAppIcon(obj As Object) As Boolean
        '<EhHeader>
        On Error GoTo SetAppIcon_Err
        '</EhHeader>
       Dim hIcon As Long
       Dim hWnd As Long
       Dim nRet As Long
   
100    If TypeOf obj Is Form Or TypeOf obj Is MDIForm Then
          ' Get top-level hidden window
102       nRet = GetWindowLong(obj.hWnd, GWL_HWNDPARENT)
104       Do While nRet
106          hWnd = nRet
108          nRet = GetWindowLong(hWnd, GWL_HWNDPARENT)
          Loop
      
          ' Get a handle to icon
110       hIcon = SendMessage(obj.hWnd, WM_GETICON, ICON_BIG, ByVal 0&)
112       If hIcon = 0 Then
             ' Load default waving-flag logo
114          hIcon = LoadIcon(0, ByVal IDI_WINLOGO)
          End If
      
          ' Pass form icon as new app icon
116       Call SendMessage(hWnd, WM_SETICON, ICON_BIG, ByVal hIcon)
                        
          ' See if change took
118       SetAppIcon = (hIcon = SendMessage(hWnd, WM_GETICON, ICON_BIG, ByVal 0&))
       End If
        '<EhFooter>
        Exit Function

SetAppIcon_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.IconStuff.SetAppIcon" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Function

