VERSION 5.00
Begin VB.Form frmConsRent 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   9225
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8895
      Begin VB.CheckBox chkDisponible 
         Caption         =   "Mostrar solo disponibles"
         Height          =   255
         Left            =   5640
         TabIndex        =   3
         Top             =   480
         Width           =   2655
      End
      Begin VB.CheckBox chkUso 
         Caption         =   "No Mostrar para uso"
         Height          =   255
         Left            =   3120
         TabIndex        =   2
         Top             =   480
         Width           =   1815
      End
      Begin VB.ComboBox cmbTipoRent 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frmConsRent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
