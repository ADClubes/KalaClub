VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmMultiClub 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control credencial usuario MC"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   8835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEntrega 
      Caption         =   "Entrega y activa credencial"
      Height          =   615
      Left            =   7200
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdAsignaCodigo 
      Caption         =   "Crea credenciales MC"
      Height          =   615
      Left            =   5640
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbgUsuarios 
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8175
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   6
      RowHeight       =   423
      Columns.Count   =   6
      Columns(0).Width=   1799
      Columns(0).Caption=   "Inscripcion"
      Columns(0).Name =   "Inscripcion"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   4471
      Columns(1).Caption=   "Nombre"
      Columns(1).Name =   "Nombre"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2593
      Columns(2).Caption=   "Status"
      Columns(2).Name =   "Status"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   2302
      Columns(3).Caption=   "FechaPago"
      Columns(3).Name =   "FechaPago"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   1693
      Columns(4).Caption=   "CodigoMC"
      Columns(4).Name =   "CodigoMC"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   3200
      Columns(5).Caption=   "IdMember"
      Columns(5).Name =   "IdMember"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      _ExtentX        =   14420
      _ExtentY        =   2778
      _StockProps     =   79
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmMultiClub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
