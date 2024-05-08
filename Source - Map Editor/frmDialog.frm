VERSION 5.00
Begin VB.Form frmDialog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   675
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   675
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lbldialog 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   135
      Width           =   4410
   End
End
Attribute VB_Name = "frmDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
