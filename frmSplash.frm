VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1440
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3810
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1440
   ScaleWidth      =   3810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblLoading 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loading spell check..."
      Height          =   225
      Left            =   555
      TabIndex        =   1
      Top             =   930
      Width           =   2730
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Right-Click Spell Checker"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   390
      Left            =   315
      TabIndex        =   0
      Top             =   255
      Width           =   3165
   End
   Begin VB.Shape shpBorder 
      Height          =   1410
      Left            =   30
      Top             =   15
      Width           =   3750
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

