Attribute VB_Name = "modSystemTray"
'***************************************************************************
'modSystemTray - adds and removes the server's icon from the SytemTray
'*****************************************************************
'Respective portions copyrighted by contributors listed below.
'
'This program is free software; you can redistribute it and/or
'modify it under the terms of the GNU General Public License
'as published by the Free Software Foundation; either version 2
'of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.
'***************************************************************************

'*****************************************************************
'Contributors History
'   When releasing modifications to this source file please add your
'   date of release, name, email, and any info to the top of this list.
'   Follow this template:
'    XX/XX/200X - Your Name Here (Your Email Here)
'       - Your Description Here
'       Sub Release Contributors:
'           XX/XX/2003 - Sub Contributor Name Here (SC Email Here)
'               - SC Description Here
'*****************************************************************
'
' David Justus - 8/14/2004
'   - First Release
'*****************************************************************

Option Explicit

Private Type NOTIFYICONDATA
  cbSize As Long
  hwnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 64
End Type

Declare Function Shell_NotifyIconA Lib "SHELL32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Integer

Private Const NIM_ADD = 0
Private Const NIM_MODIFY = 1
Private Const NIM_DELETE = 2
Private Const NIF_MESSAGE = 1
Private Const NIF_ICON = 2
Private Const NIF_TIP = 4

Private Const STI_CALLBACKEVENT = &H201

Public Const STI_LBUTTONDOWN = &H201
Public Const STI_LBUTTONUP = &H202
Public Const STI_LBUTTONDBCLK = &H203
Public Const STI_RBUTTONDOWN = &H204
Public Const STI_RBUTTONUP = &H205
Public Const STI_RBUTTONDBCLK = &H206

Public Sub CreateSystemTrayIcon(ByRef parentForm As Form, ByVal Tip As String)
'**************************************************************
'Author: David Justus
'Last Modify Date: 8/14/2004
'
'**************************************************************
  Dim notIcon As NOTIFYICONDATA
  
  With notIcon
    .cbSize = Len(notIcon)
    .hwnd = parentForm.hwnd
    .uID = vbNull
    .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
    .uCallbackMessage = STI_CALLBACKEVENT
    .hIcon = parentForm.Icon
    .szTip = Tip & vbNullChar
  End With
  
  Shell_NotifyIconA NIM_ADD, notIcon
End Sub

Public Sub ModifySystemTrayIcon(ByRef parentForm As Form, ByVal Tip As String)
'**************************************************************
'Author: David Justus
'Last Modify Date: 8/14/2004
'
'**************************************************************
  Dim notIcon As NOTIFYICONDATA
  
  With notIcon
    .cbSize = Len(notIcon)
    .hwnd = parentForm.hwnd
    .uID = vbNull
    .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
    .uCallbackMessage = STI_CALLBACKEVENT
    .hIcon = parentForm.Icon
    .szTip = Tip & vbNullChar
  End With
  
  Shell_NotifyIconA NIM_MODIFY, notIcon
End Sub

Public Sub DeleteSystemTrayIcon(ByRef parentForm As Form)
'**************************************************************
'Author: David Justus
'Last Modify Date: 8/14/2004
'
'**************************************************************
  Dim notIcon As NOTIFYICONDATA
  
  With notIcon
    .cbSize = Len(notIcon)
    .hwnd = parentForm.hwnd
    .uID = vbNull
    .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
    .uCallbackMessage = vbNull
    .hIcon = vbNull
    .szTip = "" & vbNullChar
  End With
  
  Shell_NotifyIconA NIM_DELETE, notIcon
End Sub
