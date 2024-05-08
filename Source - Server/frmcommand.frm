VERSION 5.00
Begin VB.Form frmcommand 
   BackColor       =   &H00000000&
   Caption         =   "DSO Command"
   ClientHeight    =   2130
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   ScaleHeight     =   2130
   ScaleWidth      =   8595
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   900
      TabIndex        =   2
      Top             =   1770
      Width           =   7515
   End
   Begin VB.TextBox txtlog 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1650
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   45
      Width           =   8505
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "command:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   135
      TabIndex        =   1
      Top             =   1785
      Width           =   1095
   End
End
Attribute VB_Name = "frmcommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'OREServer - v0.5.0
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

Dim acc As clsAccount

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
'**************************************************************
'Author: David Justus
'Last Modify Date: 8/14/2004
'Modified by Juan Martín Sotuyo Dodero
'**************************************************************
    Dim com As String
    Dim ctype As String
    
    com = Text1.text
    
    If KeyCode = vbKeyReturn Then
        ctype = General_Field_Read(1, com, Asc(":"))
        
        'Create new Account
        If ctype = "addacc" Then
        
            Dim user As String
            Dim pass As String
            Dim email As String
            Dim first As String
            Dim last As String
            
            user = General_Field_Read(2, com, Asc(":"))
            pass = General_Field_Read(3, com, Asc(":"))
            email = General_Field_Read(4, com, Asc(":"))
            first = General_Field_Read(5, com, Asc(":"))
            last = General_Field_Read(6, com, Asc(":"))
            
            'Make sure all data is valid
            If user = "" Or pass = "" Or email = "" Or first = "" Or last = "" Then Exit Sub
            
            acc.Account_Create user, pass, first, last, email
            txtlog.text = txtlog.text & vbNewLine & "User: " & user & " added!"
            Text1.text = ""
        End If
        
        'Change Account password
        If ctype = "accpc" Then
        
            Dim pold As String
            Dim pnew As String
            
            user = General_Field_Read(2, com, Asc(":"))
            pold = General_Field_Read(3, com, Asc(":"))
            pnew = General_Field_Read(4, com, Asc(":"))
            
            If user = "" Or pold = "" Or pnew = "" Then Exit Sub
            
            acc.Account_Password_Change user, pold, pnew
            
            txtlog.text = txtlog.text & vbNewLine & "User: " & user & " Pass Changed"
            Text1.text = ""
        End If
        
        'Delete Account
        If ctype = "delacc" Then
            user = General_Field_Read(2, com, Asc(":"))
            
            If user = "" Then Exit Sub
            
            acc.Delete_Account user
        End If
        
        'Ban Account
        If ctype = "Ban" Then
            user = General_Field_Read(2, com, Asc(":"))
            
            If user = "" Then Exit Sub
            
            acc.Account_Ban True, user
        End If
        
        'UnBan Account
        If ctype = "unBan" Then
            user = General_Field_Read(2, com, Asc(":"))
            
            If user = "" Then Exit Sub
            
            acc.Account_Ban False, user
        End If
        
        'Command line help
        If ctype = "help" Then
            txtlog.text = txtlog.text & vbNewLine & "Command Line Interface Help"
            txtlog.text = txtlog.text & vbNewLine & "addacc:<user>:<pass>:<email>:<first>:<last> - adds Account"
            txtlog.text = txtlog.text & vbNewLine & "accpc:<pass old>:<pass new> - changes Account password"
            txtlog.text = txtlog.text & vbNewLine & "delacc:<user>-deletes accounts"
            txtlog.text = txtlog.text & vbNewLine & "Ban:<user>-Bans a player"
            txtlog.text = txtlog.text & vbNewLine & "unBan:<user>-unBans a user"
        End If
    End If
End Sub
