VERSION 5.00
Begin VB.Form frmNewPlayer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Player"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   5460
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame frmPlayerInfo 
      Caption         =   "Player Information"
      Height          =   3915
      Left            =   60
      TabIndex        =   5
      Top             =   60
      Width           =   2955
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   2460
         Top             =   120
      End
      Begin VB.CommandButton cmdServerConnectNew 
         Caption         =   "Connect"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   3420
         Width           =   2715
      End
      Begin VB.TextBox txtPlayerPasswordConfirm 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   120
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1920
         Width           =   2715
      End
      Begin VB.TextBox txtPlayerName 
         Height          =   315
         Left            =   120
         MaxLength       =   20
         TabIndex        =   0
         Top             =   600
         Width           =   2715
      End
      Begin VB.TextBox txtPlayerPassword 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   120
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1260
         Width           =   2715
      End
      Begin VB.ComboBox cmbPlayerProfileName 
         Height          =   315
         ItemData        =   "frmNewPlayer.frx":0000
         Left            =   120
         List            =   "frmNewPlayer.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   2640
         Width           =   2715
      End
      Begin VB.Label Label2 
         Caption         =   "Password Confirm"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Player Name"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Password"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1020
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Player Character"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   2400
         Width           =   1515
      End
   End
   Begin VB.PictureBox picPlayer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   3060
      ScaleHeight     =   3825
      ScaleWidth      =   2325
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   2355
   End
End
Attribute VB_Name = "frmNewPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'frmNewPlayer.frm - ORE Client
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
End Sub

Private Sub Form_Unload(Cancel As Integer)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/9/2003
'
'**************************************************************
    If frmMain.Visible = False Then
        frmConnect.Show
    End If
End Sub

Private Sub Timer1_Timer()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/12/2003
'
'**************************************************************
    Static heading As Long
    heading = heading + 1
    If heading > 8 Then heading = 1
    If cmbPlayerProfileName.ListIndex <> -1 Then
        picPlayer.Cls
         frmMain.tile_engine.Grh_Render_To_Hdc frmMain.tile_engine.Char_Data_Grh_Index_Get(cmbPlayerProfileName.ItemData(cmbPlayerProfileName.ListIndex), 1, heading), _
                                                picPlayer.hdc, 50, 50, True
    End If
End Sub

Private Sub cmdServerConnectNew_Click()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/9/2003
'
'**************************************************************
    'Check
    If txtPlayerName.text = "" Or General_String_Is_Alphanumeric(txtPlayerName.text) = False Then
        MsgBox "Blank or invalid player name."
        Exit Sub
    End If

    If txtPlayerPassword.text = "" Or txtPlayerPassword.text <> txtPlayerPasswordConfirm.text Then
        MsgBox "Blank password or passwords do not match."
        Exit Sub
    End If

    'Set connection values
    frmMain.player_name = txtPlayerName.text
    frmMain.player_password = txtPlayerPassword
    
    frmConnect.txtPlayerName = txtPlayerName.text
    frmConnect.txtPlayerPassword = ""
    frmConnect.chkSavePassword = 0
    frmConnect.Save_Configuration
    
    frmMain.connection_mode_new = True
    If cmbPlayerProfileName.ListIndex <> -1 Then
        frmMain.player_profile_name = cmbPlayerProfileName.List(cmbPlayerProfileName.ListIndex)
    Else
        frmMain.player_profile_name = ""
    End If
    
    'Connect to given server
    If frmMain.dp_client.Client_Connect(frmMain.server_ip, frmMain.server_port, txtPlayerName.text) Then
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
