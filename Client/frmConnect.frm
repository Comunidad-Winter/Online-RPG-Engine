VERSION 5.00
Begin VB.Form frmConnect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connect to Server"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   2985
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame fmeServer 
      Caption         =   "Server"
      Height          =   4095
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   2955
      Begin VB.CheckBox chkSavePassword 
         Caption         =   "Save Password"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2760
         Width           =   2715
      End
      Begin VB.CommandButton cmdServerConnectNew 
         Caption         =   "New Player"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   3660
         Width           =   2715
      End
      Begin VB.CommandButton cmdServerConnect 
         Caption         =   "Connect"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   3180
         Width           =   2715
      End
      Begin VB.TextBox txtPlayerName 
         Height          =   315
         Left            =   120
         MaxLength       =   20
         TabIndex        =   2
         Top             =   1740
         Width           =   2715
      End
      Begin VB.TextBox txtServerPort 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   1140
         Width           =   2715
      End
      Begin VB.TextBox txtServerIP 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   540
         Width           =   2715
      End
      Begin VB.TextBox txtPlayerPassword 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   120
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   2340
         Width           =   2715
      End
      Begin VB.Label Label1 
         Caption         =   "Player Name"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1500
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Port"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   900
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Server IP"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Password"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   2100
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'frmConnect.frm - ORE Client
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
'Aaron Perkins(aaron@baronsoft.com) - 8/04/2003
'   - First Release
'*****************************************************************

Option Explicit

Private Sub Form_Load()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/9/2003
'
'**************************************************************
    Load_Configuration
End Sub

Private Sub Form_Unload(Cancel As Integer)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/9/2003
'
'**************************************************************
    If frmMain.Visible = False Then
        frmMain.main_loop_go = False
    End If
End Sub

Public Sub Load_Configuration()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/9/2003
'
'**************************************************************
    Dim file_path As String
    file_path = App.Path & "\" & "client.ini"
    
    If General_File_Exists(file_path, vbNormal) Then
        txtServerIP.text = General_Var_Get(file_path, "GENERAL", "server_ip")
        txtServerPort.text = General_Var_Get(file_path, "GENERAL", "server_port")
        txtPlayerName.text = General_Var_Get(file_path, "GENERAL", "player_name")
        txtPlayerPassword.text = General_Var_Get(file_path, "GENERAL", "player_password")
        If txtPlayerPassword.text <> "" Then
            chkSavePassword.value = 1
        End If
    End If
End Sub

Public Sub Save_Configuration()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/9/2003
'
'**************************************************************
    Dim file_path As String
    file_path = App.Path & "\" & "client.ini"
    
    If General_File_Exists(file_path, vbNormal) Then
        Kill file_path
    End If
        
    General_Var_Write file_path, "GENERAL", "server_ip", txtServerIP.text
    General_Var_Write file_path, "GENERAL", "server_port", txtServerPort.text
    General_Var_Write file_path, "GENERAL", "player_name", txtPlayerName.text
    If chkSavePassword.value Then
        General_Var_Write file_path, "GENERAL", "player_password", txtPlayerPassword.text
    End If
End Sub

Private Sub cmdServerConnect_Click()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/9/2003
'
'**************************************************************
    'Check
    If txtServerIP.text = "" Or txtServerPort.text = "" Then
        MsgBox "Blank or invalid server name or port."
        Exit Sub
    End If
    
    If txtPlayerName.text = "" Or General_String_Is_Alphanumeric(txtPlayerName.text) = False Then
        MsgBox "Blank or invalid player name."
        Exit Sub
    End If

    If txtPlayerPassword.text = "" Then
        MsgBox "Blank password."
        Exit Sub
    End If
    
    'Save
    Save_Configuration
    
    'Set connection values
    frmMain.server_ip = txtServerIP.text
    frmMain.server_port = CLng(txtServerPort.text)
    frmMain.player_name = txtPlayerName.text
    frmMain.player_password = txtPlayerPassword
    
    frmMain.connection_mode_new = False
    frmMain.player_profile_name = ""
    
    'Connect to given server
    If frmMain.dp_client.Client_Connect(txtServerIP.text, CLng(txtServerPort.text), txtPlayerName.text) Then
        'Show main form
        frmMain.Show
        'Set Status
        frmMain.lblServerStatus.Caption = "Connecting ..."
        'Get rid of this form
        Unload Me
    Else
        MsgBox "There was an error while trying to connect to the server."
    End If
End Sub

Private Sub cmdServerConnectNew_Click()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/9/2003
'
'**************************************************************
    'Check
    If txtServerIP.text = "" Or txtServerPort.text = "" Then
        MsgBox "Blank or invalid server name or port."
        Exit Sub
    End If
    
    'Save
    Save_Configuration
    
    'Set connection values
    frmMain.server_ip = txtServerIP.text
    frmMain.server_port = CLng(txtServerPort.text)

    Me.Hide
    frmNewPlayer.Show
    General_Form_On_Top_Set frmNewPlayer, True
End Sub
