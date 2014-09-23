VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmFacturacion 
   Caption         =   "Facturación"
   ClientHeight    =   8790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11280
   Icon            =   "frmFacturacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8790
   ScaleWidth      =   11280
   Begin VB.CommandButton cmdCuponesDesc 
      Caption         =   "Cupones Desc."
      Height          =   495
      Left            =   9600
      TabIndex        =   71
      Top             =   4200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdPagoEnEfectivo 
      Caption         =   "Solo efectivo"
      Height          =   495
      Left            =   11160
      TabIndex        =   70
      Top             =   6840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame frmDireccionar 
      Height          =   1095
      Left            =   240
      TabIndex        =   67
      Top             =   2400
      Width           =   2055
      Begin VB.ComboBox cmbTipoDireccionar 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   600
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CheckBox chkDireccionar 
         Caption         =   "Direccionar"
         Height          =   255
         Left            =   120
         TabIndex        =   68
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdEstacionamiento 
      Caption         =   "Estacionamiento"
      Height          =   495
      Left            =   11880
      TabIndex        =   66
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton cmdPagoRapido 
      Caption         =   "Pago rápido"
      Height          =   495
      Left            =   9840
      TabIndex        =   65
      Top             =   6840
      Width           =   1095
   End
   Begin SSDataWidgets_B.SSDBGrid ssdbgPagos 
      Height          =   1935
      Left            =   7200
      TabIndex        =   36
      Top             =   4800
      Width           =   6135
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   12
      SelectTypeRow   =   1
      RowHeight       =   423
      Columns.Count   =   12
      Columns(0).Width=   3200
      Columns(0).Visible=   0   'False
      Columns(0).Caption=   "IdTipoPago"
      Columns(0).Name =   "IdTipoPago"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3387
      Columns(1).Caption=   "Pago"
      Columns(1).Name =   "DescripPago"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   1931
      Columns(2).Caption=   "Opcion"
      Columns(2).Name =   "Opcion"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   2196
      Columns(3).Caption=   "Importe"
      Columns(3).Name =   "Importe"
      Columns(3).Alignment=   1
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   6
      Columns(3).NumberFormat=   "CURRENCY"
      Columns(3).FieldLen=   256
      Columns(3).Locked=   -1  'True
      Columns(4).Width=   3200
      Columns(4).Caption=   "Referencia"
      Columns(4).Name =   "Referencia"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(4).Locked=   -1  'True
      Columns(5).Width=   3200
      Columns(5).Visible=   0   'False
      Columns(5).Caption=   "Resta"
      Columns(5).Name =   "Resta"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   6
      Columns(5).NumberFormat=   "CURRENCY"
      Columns(5).FieldLen=   256
      Columns(6).Width=   3200
      Columns(6).Visible=   0   'False
      Columns(6).Caption=   "IdAfiliacion"
      Columns(6).Name =   "IdAfiliacion"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   3200
      Columns(7).Visible=   0   'False
      Columns(7).Caption=   "LoteNumero"
      Columns(7).Name =   "LoteNumero"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   3200
      Columns(8).Visible=   0   'False
      Columns(8).Caption=   "OperacionNumero"
      Columns(8).Name =   "OperacionNumero"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   3200
      Columns(9).Visible=   0   'False
      Columns(9).Caption=   "ImporteRecibido"
      Columns(9).Name =   "ImporteRecibido"
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      Columns(10).Width=   3200
      Columns(10).Caption=   "FechaOperacion"
      Columns(10).Name=   "FechaOperacion"
      Columns(10).DataField=   "Column 10"
      Columns(10).DataType=   8
      Columns(10).FieldLen=   256
      Columns(11).Width=   3200
      Columns(11).Caption=   "NoCuenta"
      Columns(11).Name=   "NoCuenta"
      Columns(11).DataField=   "Column 11"
      Columns(11).DataType=   8
      Columns(11).FieldLen=   256
      _ExtentX        =   10821
      _ExtentY        =   3413
      _StockProps     =   79
      Caption         =   "Pagos"
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
   Begin VB.CommandButton cmdInsertaPagoEx 
      Caption         =   "Inserta Pago"
      Height          =   495
      Left            =   7200
      TabIndex        =   64
      Top             =   6840
      Width           =   1095
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssCmbOpcionPago 
      Height          =   375
      Left            =   9600
      TabIndex        =   62
      Top             =   6000
      Visible         =   0   'False
      Width           =   2415
      DataFieldList   =   "Column 0"
      _Version        =   196616
      DataMode        =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   3651
      Columns(0).Caption=   "Descripcion"
      Columns(0).Name =   "Descripcion"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Visible=   0   'False
      Columns(1).Caption=   "IdFormadepagoOpcion"
      Columns(1).Name =   "IdFormadepagoOpcion"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   4260
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin VB.ComboBox cmbPeriodoCalculo 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   61
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdModRen 
      Caption         =   "Modifica Ren."
      Height          =   495
      Left            =   7320
      TabIndex        =   60
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton cmdQuitaInteres 
      Caption         =   "Quita Intereses"
      Height          =   615
      Left            =   3240
      TabIndex        =   59
      Top             =   7920
      Width           =   975
   End
   Begin VB.CommandButton cmdObser 
      Caption         =   "Obser."
      Height          =   495
      Left            =   11520
      TabIndex        =   58
      Top             =   7680
      Width           =   855
   End
   Begin VB.CheckBox chkDatosFac 
      Caption         =   "Modificar datos factura"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   53
      Top             =   3840
      Width           =   2055
   End
   Begin VB.CheckBox chkEliminaTodos 
      Caption         =   "Eliminar todos"
      Height          =   495
      Left            =   2160
      TabIndex        =   51
      Top             =   7920
      Width           =   855
   End
   Begin VB.TextBox txtObserva 
      Height          =   495
      Left            =   7200
      MaxLength       =   150
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   49
      Top             =   7680
      Width           =   4215
   End
   Begin VB.CommandButton cmdEliminaPago 
      Caption         =   "Elimina Pago"
      Height          =   495
      Left            =   8520
      TabIndex        =   44
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton cmdInsertaPago 
      Caption         =   "Inserta Pago"
      Height          =   495
      Left            =   7200
      TabIndex        =   43
      Top             =   6840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtImportePago 
      Height          =   405
      Left            =   12240
      MaxLength       =   20
      TabIndex        =   39
      Top             =   6240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtRefer 
      Height          =   405
      Left            =   13560
      MaxLength       =   20
      TabIndex        =   38
      Top             =   6240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo sscmbFormaPago 
      Height          =   375
      Left            =   7200
      TabIndex        =   37
      Top             =   6000
      Visible         =   0   'False
      Width           =   2295
      DataFieldList   =   "Column 0"
      AllowInput      =   0   'False
      _Version        =   196616
      DataMode        =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RowHeight       =   423
      Columns.Count   =   3
      Columns(0).Width=   4286
      Columns(0).Caption=   "FormaPago"
      Columns(0).Name =   "FormaPago"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Visible=   0   'False
      Columns(1).Caption=   "IdFormaPago"
      Columns(1).Name =   "IdFormaPago"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   3200
      Columns(2).Visible=   0   'False
      Columns(2).Caption=   "RequiereReferencia"
      Columns(2).Name =   "RequiereReferencia"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   17
      Columns(2).FieldLen=   256
      _ExtentX        =   4048
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin VB.CommandButton cmdDescuento 
      Caption         =   "Descuento"
      Height          =   615
      Left            =   4560
      TabIndex        =   35
      Top             =   7920
      Width           =   975
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Eliminar"
      Height          =   615
      Left            =   1200
      TabIndex        =   34
      Top             =   7920
      Width           =   855
   End
   Begin VB.CommandButton cmdIns 
      Caption         =   "&Insertar"
      Height          =   615
      Left            =   120
      TabIndex        =   33
      Top             =   7920
      Width           =   855
   End
   Begin VB.CommandButton cmdFacturar 
      Caption         =   "&Facturar"
      Height          =   615
      Left            =   5880
      TabIndex        =   32
      Top             =   7920
      Width           =   1095
   End
   Begin VB.Frame frmDatosFactura 
      Caption         =   "Datos de Facturación"
      Height          =   3615
      Left            =   120
      TabIndex        =   8
      Top             =   4200
      Width           =   6975
      Begin VB.OptionButton optTipoPer 
         Caption         =   "Persona Moral"
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   2520
         TabIndex        =   73
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton optTipoPer 
         Caption         =   "Persona Física"
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   72
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.TextBox txtFacTelefono 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   18
         Top             =   3240
         Width           =   1695
      End
      Begin VB.TextBox txtFacDelOMuni 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4680
         MaxLength       =   50
         TabIndex        =   14
         Top             =   2160
         Width           =   1935
      End
      Begin SSDataWidgets_B.SSDBCombo ssCmbFacEstado 
         Height          =   255
         Left            =   4680
         TabIndex        =   17
         Top             =   2760
         Width           =   2055
         DataFieldList   =   "Column 0"
         _Version        =   196616
         DataMode        =   2
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColorEven   =   0
         BackColorOdd    =   16777215
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   3200
         Columns(0).Caption=   "Estado"
         Columns(0).Name =   "Estado"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Visible=   0   'False
         Columns(1).Caption=   "IdEstado"
         Columns(1).Name =   "IdEstado"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   93
         BackColor       =   -2147483643
         Enabled         =   0   'False
      End
      Begin VB.TextBox txtFacCP 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3720
         MaxLength       =   5
         TabIndex        =   16
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox txtFacRFC 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4800
         MaxLength       =   20
         TabIndex        =   11
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtFacCiudad 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   15
         Top             =   2760
         Width           =   3255
      End
      Begin VB.TextBox txtFacColonia 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   13
         Top             =   2160
         Width           =   4455
      End
      Begin VB.TextBox txtFacDireccion 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   12
         Top             =   1560
         Width           =   6495
      End
      Begin VB.TextBox txtFacNombre 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         MaxLength       =   80
         TabIndex        =   10
         Top             =   960
         Width           =   4455
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "Teléfono"
         Height          =   255
         Left            =   360
         TabIndex        =   56
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label lblTipoDir 
         Enabled         =   0   'False
         Height          =   255
         Left            =   5760
         TabIndex        =   55
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblIdDireccion 
         Enabled         =   0   'False
         Height          =   255
         Left            =   4800
         TabIndex        =   54
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Estado"
         Height          =   255
         Left            =   4800
         TabIndex        =   25
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label Label8 
         Caption         =   "C.P."
         Height          =   255
         Left            =   3720
         TabIndex        =   24
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "RFC"
         Height          =   255
         Left            =   4680
         TabIndex        =   23
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Ciudad"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Delegación/Municipio"
         Height          =   255
         Left            =   4680
         TabIndex        =   21
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "Colonia"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Direccion"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdCalcular 
      Caption         =   "&Calcular"
      Default         =   -1  'True
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton cmdAyudaClave 
      Height          =   375
      Left            =   1800
      Picture         =   "frmFacturacion.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   240
      Width           =   495
   End
   Begin MSComCtl2.DTPicker dtpFechaCalc 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd/MMM/yyyy"
      Format          =   50528259
      CurrentDate     =   38097
   End
   Begin SSDataWidgets_B.SSDBGrid ssdbgFactura 
      Height          =   3495
      Left            =   2520
      TabIndex        =   3
      Top             =   480
      Width           =   12375
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   23
      SelectTypeRow   =   1
      BackColorOdd    =   12632256
      RowHeight       =   423
      Columns.Count   =   23
      Columns(0).Width=   6482
      Columns(0).Caption=   "Concepto"
      Columns(0).Name =   "Concepto"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   4683
      Columns(1).Caption=   "Nombre"
      Columns(1).Name =   "Nombre"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   1879
      Columns(2).Caption=   "Periodo"
      Columns(2).Name =   "Periodo"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(2).Locked=   -1  'True
      Columns(3).Width=   873
      Columns(3).Caption=   "Cant."
      Columns(3).Name =   "Cantidad"
      Columns(3).Alignment=   1
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   2
      Columns(3).FieldLen=   3
      Columns(3).Locked=   -1  'True
      Columns(3).Mask =   "###"
      Columns(4).Width=   2011
      Columns(4).Caption=   "Importe"
      Columns(4).Name =   "Importe"
      Columns(4).Alignment=   1
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   6
      Columns(4).NumberFormat=   "CURRENCY"
      Columns(4).FieldLen=   256
      Columns(4).Locked=   -1  'True
      Columns(4).Mask =   "999,999,999.99"
      Columns(5).Width=   1720
      Columns(5).Caption=   "Intereses"
      Columns(5).Name =   "Intereses"
      Columns(5).Alignment=   1
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   6
      Columns(5).NumberFormat=   "CURRENCY"
      Columns(5).FieldLen=   256
      Columns(5).Locked=   -1  'True
      Columns(6).Width=   1349
      Columns(6).Caption=   "Desc."
      Columns(6).Name =   "Descuento"
      Columns(6).Alignment=   1
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   5
      Columns(6).NumberFormat=   "##0.00\%"
      Columns(6).FieldLen=   6
      Columns(6).Locked=   -1  'True
      Columns(6).Nullable=   0
      Columns(7).Width=   2117
      Columns(7).Caption=   "Total"
      Columns(7).Name =   "Total"
      Columns(7).Alignment=   1
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   6
      Columns(7).NumberFormat=   "CURRENCY"
      Columns(7).FieldLen=   256
      Columns(7).Locked=   -1  'True
      Columns(8).Width=   1984
      Columns(8).Caption=   "Clave"
      Columns(8).Name =   "Clave"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(8).Locked=   -1  'True
      Columns(9).Width=   3200
      Columns(9).Visible=   0   'False
      Columns(9).Caption=   "IvaPor"
      Columns(9).Name =   "IvaPor"
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   5
      Columns(9).FieldLen=   256
      Columns(10).Width=   3200
      Columns(10).Visible=   0   'False
      Columns(10).Caption=   "Iva"
      Columns(10).Name=   "Iva"
      Columns(10).DataField=   "Column 10"
      Columns(10).DataType=   8
      Columns(10).FieldLen=   256
      Columns(11).Width=   3200
      Columns(11).Visible=   0   'False
      Columns(11).Caption=   "IvaDescuento"
      Columns(11).Name=   "IvaDescuento"
      Columns(11).DataField=   "Column 11"
      Columns(11).DataType=   8
      Columns(11).FieldLen=   256
      Columns(12).Width=   3200
      Columns(12).Visible=   0   'False
      Columns(12).Caption=   "IvaIntereses"
      Columns(12).Name=   "IvaIntereses"
      Columns(12).DataField=   "Column 12"
      Columns(12).DataType=   8
      Columns(12).FieldLen=   256
      Columns(13).Width=   3200
      Columns(13).Visible=   0   'False
      Columns(13).Caption=   "DescMonto"
      Columns(13).Name=   "DescMonto"
      Columns(13).DataField=   "Column 13"
      Columns(13).DataType=   6
      Columns(13).NumberFormat=   "CURRENCY"
      Columns(13).FieldLen=   256
      Columns(14).Width=   3200
      Columns(14).Visible=   0   'False
      Columns(14).Caption=   "IdMember"
      Columns(14).Name=   "IdMember"
      Columns(14).DataField=   "Column 14"
      Columns(14).DataType=   8
      Columns(14).FieldLen=   256
      Columns(15).Width=   3200
      Columns(15).Visible=   0   'False
      Columns(15).Caption=   "NoFamiliar"
      Columns(15).Name=   "NoFamiliar"
      Columns(15).DataField=   "Column 15"
      Columns(15).DataType=   8
      Columns(15).FieldLen=   256
      Columns(16).Width=   3200
      Columns(16).Visible=   0   'False
      Columns(16).Caption=   "FormaPago"
      Columns(16).Name=   "FormaPago"
      Columns(16).DataField=   "Column 16"
      Columns(16).DataType=   2
      Columns(16).FieldLen=   256
      Columns(17).Width=   3200
      Columns(17).Visible=   0   'False
      Columns(17).Caption=   "IdTipoUsuario"
      Columns(17).Name=   "IdTipoUsuario"
      Columns(17).DataField=   "Column 17"
      Columns(17).DataType=   3
      Columns(17).FieldLen=   256
      Columns(18).Width=   3200
      Columns(18).Visible=   0   'False
      Columns(18).Caption=   "TipoCargo"
      Columns(18).Name=   "TipoCargo"
      Columns(18).DataField=   "Column 18"
      Columns(18).DataType=   8
      Columns(18).FieldLen=   256
      Columns(19).Width=   3200
      Columns(19).Visible=   0   'False
      Columns(19).Caption=   "Auxiliar"
      Columns(19).Name=   "Auxiliar"
      Columns(19).DataField=   "Column 19"
      Columns(19).DataType=   8
      Columns(19).FieldLen=   256
      Columns(20).Width=   3200
      Columns(20).Visible=   0   'False
      Columns(20).Caption=   "FacORec"
      Columns(20).Name=   "FacORec"
      Columns(20).DataField=   "Column 20"
      Columns(20).DataType=   8
      Columns(20).FieldLen=   256
      Columns(21).Width=   3200
      Columns(21).Visible=   0   'False
      Columns(21).Caption=   "IdInstructor"
      Columns(21).Name=   "IdInstructor"
      Columns(21).DataField=   "Column 21"
      Columns(21).DataType=   8
      Columns(21).FieldLen=   256
      Columns(22).Width=   3200
      Columns(22).Visible=   0   'False
      Columns(22).Caption=   "Unidad"
      Columns(22).Name=   "Unidad"
      Columns(22).DataField=   "Column 22"
      Columns(22).DataType=   8
      Columns(22).FieldLen=   256
      _ExtentX        =   21828
      _ExtentY        =   6165
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
   Begin VB.TextBox txtClave 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label21 
      Caption         =   "Opcion pago"
      Height          =   255
      Left            =   9720
      TabIndex        =   63
      Top             =   5760
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblInscripcion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9840
      TabIndex        =   57
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lblRen 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   12840
      TabIndex        =   52
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label19 
      Caption         =   "Observaciones"
      Height          =   255
      Left            =   7200
      TabIndex        =   50
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Label lblPorPagar 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   13560
      TabIndex        =   48
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label lblPagado 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   13560
      TabIndex        =   47
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      Caption         =   "Por pagar:"
      Height          =   255
      Left            =   13560
      TabIndex        =   46
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Caption         =   "Total Pagado:"
      Height          =   375
      Left            =   13680
      TabIndex        =   45
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label Label16 
      Caption         =   "Importe"
      Height          =   255
      Left            =   12360
      TabIndex        =   42
      Top             =   5760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label15 
      Caption         =   "Referencia"
      Height          =   255
      Left            =   13800
      TabIndex        =   41
      Top             =   6000
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label14 
      Caption         =   "Forma de pago"
      Height          =   255
      Left            =   7320
      TabIndex        =   40
      Top             =   5760
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label13 
      Caption         =   "Fecha de Cálculo"
      Height          =   255
      Left            =   240
      TabIndex        =   31
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Total"
      Height          =   255
      Left            =   12240
      TabIndex        =   30
      Top             =   8040
      Width           =   1095
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "Descuento"
      Height          =   255
      Left            =   12240
      TabIndex        =   29
      Top             =   7680
      Width           =   1095
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "SubTotal"
      Height          =   255
      Left            =   12480
      TabIndex        =   28
      Top             =   7320
      Width           =   855
   End
   Begin VB.Label lblSubTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13440
      TabIndex        =   27
      Top             =   7320
      Width           =   1455
   End
   Begin VB.Label lblDescuento 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13440
      TabIndex        =   26
      Top             =   7680
      Width           =   1455
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13440
      TabIndex        =   7
      Top             =   8040
      Width           =   1455
   End
   Begin VB.Label lblNombreUsu 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   6
      Top             =   120
      Width           =   6735
   End
   Begin VB.Label Label1 
      Caption         =   "Clave"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
End
Attribute VB_Name = "frmFacturacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lIdTitular As Long
Dim lIdFamilia As Long

Dim sCargos(1000) As String 'Arreglo para facturar
Dim lIndexsCargos As Long   'Indice del arreglo de facturación
Dim lPeriodos As Long       'Número de periodos adeudados
Dim dFechaUltimoPago As Date 'Fecha del ultimo pago del usuario
Dim dPeriodo As Date        'Periodo actual del calculo
Dim dPeriodoIni As Date     'Fecha de inicio del periodo actual
Dim lFormaPago As Integer   'Periodicidad del pago

'Variables por Renglon
Dim dCuota As Double        'Cuota como esta en el Historico, sin IVA
Dim dMonto As Double        'Monto unitario de la cuota
Dim dCantidad As Double     'Cantidad
Dim nPorAus As Double       'Porcentaje de ausencia
Dim dImporte As Double      'Monto * Cantidad
Dim dIntereses As Double    'Intereses por renglon
Dim dDescuentoPor As Double 'Descuento por renglon en porcentaje
Dim dDescuento As Double    'Descuento por renglon en pesos
Dim dTotal As Double        'Importe + Intereses - Descuento
Dim dIvaPor                 'Iva en porcentaje
Dim dIva                    'Iva del Importe en pesos
Dim dIvaIntereses           'Iva de los Intereses
Dim dIvaDescuento           'Iva del descuento


'Variables totales
Dim siTotImporte As Double      'Suma de todos los importes
Dim siTotIntereses As Double     'Suma de todos los intereses
Dim siTotDescuento As Double     'Suma de todos los descuentos
Dim siTotTotal As Double         'Suma de todos los Totales
Dim siTotIva As Double           'Suma de todos los ivas
Dim siTotIvaIntereses As Double  'Suma de todos los ivas de intereses

Dim siTotalPagado As Double
Dim siTotalPorPagar As Double

'Para controlar los direccionados
Dim boDireccionado As Boolean
Dim boYaFacturado As Boolean    'Controla cuando esta con datos cargados
Dim lPublicoGeneral As Boolean

Private Sub chkDatosFac_Click()
    
    Me.optTipoPer(0).Enabled = chkDatosFac.Value
    Me.optTipoPer(1).Enabled = chkDatosFac.Value
    Me.txtFacNombre.Enabled = chkDatosFac.Value
    Me.txtFacRFC.Enabled = chkDatosFac.Value
    Me.txtFacDireccion.Enabled = chkDatosFac.Value
    Me.txtFacColonia.Enabled = chkDatosFac.Value
    Me.txtFacDelOMuni.Enabled = chkDatosFac.Value
    Me.txtFacCiudad.Enabled = chkDatosFac.Value
    Me.txtFacCP.Enabled = chkDatosFac.Value
    Me.ssCmbFacEstado.Enabled = chkDatosFac.Value
    Me.txtFacTelefono.Enabled = chkDatosFac.Value
End Sub

Private Sub chkDireccionar_Click()
    Dim byRespuesta As Byte
    Dim strValAnt As String
    Dim byValAnt As Byte
    
    'Evita la llamada recursiva
    If boDireccionado Then
        Exit Sub
    End If
    
    'Guarda el valor actual
    byValAnt = Me.chkDireccionar.Value
    
    boDireccionado = True
    
    If boYaFacturado Then
        byRespuesta = MsgBox("Si selecciona Direccionamiento se reiniciará" & vbCrLf & "el cálculo ¿Desea Continuar?", vbQuestion Or vbYesNo)
        If byRespuesta = vbYes Then
            strValAnt = Trim(Me.txtClave.Text)
            
            LimpiaTodo
            boDireccionado = True
            Me.chkDireccionar.Value = byValAnt
            Me.txtClave.Text = strValAnt
           
        Else
            Me.chkDireccionar.Value = IIf(Me.chkDireccionar.Value, 0, 1)
        End If
    End If
    
    Me.cmbTipoDireccionar.Visible = chkDireccionar.Value
    
    boDireccionado = False
End Sub

Private Sub cmdAyudaClave_Click()
    
    Dim frmAyuClave As frmAyudaClave
    Dim sCveAnt As String
    
    sCveAnt = Trim(Me.txtClave.Text)
    
    Set frmAyuClave = New frmAyudaClave
    
    frmAyuClave.Show 1
    
    
    If Trim(Me.txtClave.Text) <> sCveAnt Then
        Me.cmdCalcular.Value = True
    End If
End Sub

Private Sub cmdCalcular_Click()
    Dim adoRcsDireccionado As ADODB.Recordset
    Dim adoRcsDireccion As ADODB.Recordset
    Dim adoRcsFac As ADODB.Recordset
    Dim lidMember As Long
    Dim i As Long
    Dim sRenglon As Variant
    Dim lRespuesta As Integer
    Dim sOldClave As String
    
    If Me.lblNombreUsu <> vbNullString Then
        lRespuesta = MsgBox("¿Desea recalcular?", vbQuestion Or vbYesNo, "Facturación")
        If lRespuesta <> vbYes Then
            Exit Sub
        End If
        sOldClave = Me.txtClave.Text
        LimpiaTodo
        Me.txtClave.Text = sOldClave
    End If
    
    If Me.txtClave.Text = vbNullString Then
        MsgBox "Elegir un Usuario! ", vbInformation, "Facturación"
        Me.txtClave.SetFocus
        Exit Sub
    End If
        
    Me.txtClave.Text = Val(Trim(Me.txtClave.Text))
    
    'Busca nombre de usuario
    If Me.lblNombreUsu = vbNullString Then
        strSQL = "SELECT A_PATERNO, A_MATERNO, NOMBRE, IdMember, Inscripcion"
        strSQL = strSQL & " FROM USUARIOS_CLUB"
        strSQL = strSQL & " WHERE NoFamilia=" & Me.txtClave.Text & " AND IdMember=IdTitular"
        
        Set adoRcsFac = New ADODB.Recordset
        adoRcsFac.ActiveConnection = Conn
        adoRcsFac.CursorLocation = adUseServer
        adoRcsFac.CursorType = adOpenForwardOnly
        adoRcsFac.LockType = adLockReadOnly
        adoRcsFac.Open strSQL
        
        If Not adoRcsFac.EOF Then
            Me.lblNombreUsu = adoRcsFac!A_Paterno & " " & adoRcsFac!A_Materno & " " & adoRcsFac!Nombre
            Me.txtFacNombre.Text = Me.lblNombreUsu
            lidMember = adoRcsFac!Idmember
            
            Me.lblInscripcion = IIf(IsNull(adoRcsFac!inscripcion), "", Trim(adoRcsFac!inscripcion))
        Else
            MsgBox "No existe el Usuario!", vbInformation, "Facturacion"
        
        End If
        
        lIdFamilia = Val(Me.txtClave.Text)
        lIdTitular = lidMember
        
        adoRcsFac.Close
        
        If Me.lblNombreUsu = vbNullString Then
            Me.txtClave.SelStart = 0
            Me.txtClave.SelLength = Len(Me.txtClave.Text)
            Me.txtClave.SetFocus
            Exit Sub
        End If
        
        'Verifica si está direccionado
        #If SqlServer_ Then
            strSQL = "SELECT TIPODIRECCIONADO"
            strSQL = strSQL & " FROM DIRECCIONADOS"
            strSQL = strSQL & " WHERE"
            strSQL = strSQL & " IdMember=" & lidMember
            strSQL = strSQL & " AND Activo=1"
        #Else
            strSQL = "SELECT TIPODIRECCIONADO"
            strSQL = strSQL & " FROM DIRECCIONADOS"
            strSQL = strSQL & " WHERE"
            strSQL = strSQL & " IdMember=" & lidMember
            strSQL = strSQL & " AND Activo=-1"
        #End If

        Set adoRcsDireccionado = New ADODB.Recordset
        adoRcsDireccionado.ActiveConnection = Conn
        adoRcsDireccionado.CursorLocation = adUseServer
        adoRcsDireccionado.CursorType = adOpenForwardOnly
        adoRcsDireccionado.LockType = adLockReadOnly
        adoRcsDireccionado.Open strSQL

        If Not adoRcsDireccionado.EOF Then
            If Not IsNull(adoRcsDireccionado!TipoDireccionado) Then
                Select Case adoRcsDireccionado!TipoDireccionado
                    Case "TC"
                        Me.cmbTipoDireccionar.Text = "T.CREDITO"
                    Case "TD"
                        Me.cmbTipoDireccionar.Text = "T.DEBITO"
                End Select
                Me.chkDireccionar.Value = 1
            End If
        End If

        adoRcsDireccionado.Close
        Set adoRcsDireccionado = Nothing
        
        'Busca direccion
        'Primero busca direccion fiscal
        
        strSQL = "SELECT RazonSocial, RFC, CALLE, COLONIA, DELOMUNI, Ciudad, CODPOS, TEL1, TEL2, Estado, IdDireccion, IdTipoDireccion, TipoPersona"
        strSQL = strSQL & " FROM DIRECCIONES DIR"
        strSQL = strSQL & " WHERE IDMEMBER=" & lidMember
        strSQL = strSQL & " AND IDTIPODIRECCION=3"
        
        Set adoRcsDireccion = New ADODB.Recordset
        With adoRcsDireccion
            .ActiveConnection = Conn
            .CursorLocation = adUseServer
            .CursorType = adOpenForwardOnly
            .LockType = adLockReadOnly
            .Open strSQL
        End With
        
        If Not adoRcsDireccion.EOF Then
            Me.txtFacNombre.Text = IIf(IsNull(adoRcsDireccion!RazonSocial), Me.lblNombreUsu, adoRcsDireccion!RazonSocial)
            Me.txtFacRFC.Text = IIf(IsNull(adoRcsDireccion!rfc), vbNullString, adoRcsDireccion!rfc)
            Me.txtFacDireccion = IIf(IsNull(adoRcsDireccion!calle), vbNullString, adoRcsDireccion!calle)
            Me.txtFacColonia = IIf(IsNull(adoRcsDireccion!colonia), vbNullString, adoRcsDireccion!colonia)
            Me.txtFacDelOMuni.Text = IIf(IsNull(adoRcsDireccion!DeloMuni), vbNullString, adoRcsDireccion!DeloMuni)
            Me.txtFacCiudad.Text = IIf(IsNull(adoRcsDireccion!Ciudad), vbNullString, adoRcsDireccion!Ciudad)
            Me.txtFacCP = IIf(IsNull(adoRcsDireccion!Codpos), vbNullString, Format(adoRcsDireccion!Codpos, "00000"))
            Me.ssCmbFacEstado.Text = IIf(IsNull(adoRcsDireccion!Estado), vbNullString, adoRcsDireccion!Estado)
            Me.txtFacTelefono.Text = IIf(IsNull(adoRcsDireccion!Tel1), vbNullString, adoRcsDireccion!Tel1)
            Me.lblIdDireccion = IIf(IsNull(adoRcsDireccion!IdDireccion), vbNullString, adoRcsDireccion!IdDireccion)
            Me.lblTipoDir = IIf(IsNull(adoRcsDireccion!IdTipoDireccion), vbNullString, adoRcsDireccion!IdTipoDireccion)
            If adoRcsDireccion!TipoPersona = "F" Then
                Me.optTipoPer(0).Value = True
            Else
                Me.optTipoPer(1).Value = True
            End If
        Else
            adoRcsDireccion.Close
            strSQL = "SELECT RazonSocial, RFC, CALLE, COLONIA, DELOMUNI, Ciudad, CODPOS, TEL1, TEL2, Estado, IdDireccion, IdTipoDireccion, TipoPersona"
            strSQL = strSQL & " FROM DIRECCIONES DIR"
            strSQL = strSQL & " WHERE IDMEMBER=" & lidMember
            strSQL = strSQL & " AND IDTIPODIRECCION=1"
            
            
            adoRcsDireccion.Open strSQL
            
            If Not adoRcsDireccion.EOF Then
                Me.txtFacNombre.Text = IIf(IsNull(adoRcsDireccion!RazonSocial), Me.lblNombreUsu, adoRcsDireccion!RazonSocial)
                Me.txtFacRFC.Text = IIf(IsNull(adoRcsDireccion!rfc), vbNullString, adoRcsDireccion!rfc)
                Me.txtFacDireccion = IIf(IsNull(adoRcsDireccion!calle), vbNullString, adoRcsDireccion!calle)
                Me.txtFacColonia = IIf(IsNull(adoRcsDireccion!colonia), vbNullString, adoRcsDireccion!colonia)
                Me.txtFacDelOMuni.Text = IIf(IsNull(adoRcsDireccion!DeloMuni), vbNullString, adoRcsDireccion!DeloMuni)
                Me.txtFacCiudad.Text = IIf(IsNull(adoRcsDireccion!Ciudad), vbNullString, adoRcsDireccion!Ciudad)
                Me.txtFacCP = IIf(IsNull(adoRcsDireccion!Codpos), vbNullString, Format(adoRcsDireccion!Codpos, "00000"))
                Me.ssCmbFacEstado.Text = IIf(IsNull(adoRcsDireccion!Estado), vbNullString, adoRcsDireccion!Estado)
                Me.txtFacTelefono.Text = IIf(IsNull(adoRcsDireccion!Tel1), vbNullString, adoRcsDireccion!Tel1)
                Me.lblIdDireccion.Caption = IIf(IsNull(adoRcsDireccion!IdDireccion), vbNullString, adoRcsDireccion!IdDireccion)
                Me.lblTipoDir.Caption = IIf(IsNull(adoRcsDireccion!IdTipoDireccion), vbNullString, adoRcsDireccion!IdTipoDireccion)
                If adoRcsDireccion!TipoPersona = "F" Then
                    Me.optTipoPer(0).Value = True
                Else
                    Me.optTipoPer(1).Value = True
                End If
            Else
                Me.txtFacNombre.Text = Me.lblNombreUsu
                Me.txtFacDireccion = vbNullString
                Me.txtFacColonia = vbNullString
                Me.txtFacDelOMuni.Text = vbNullString
                Me.txtFacCP = vbNullString
                Me.ssCmbFacEstado.Text = vbNullString
                Me.txtFacTelefono.Text = vbNullString
                Me.lblIdDireccion.Caption = vbNullString
                Me.lblTipoDir.Caption = vbNullString
                Me.optTipoPer(0).Value = True
            End If
            
        End If
        
        adoRcsDireccion.Close
        Set adoRcsDireccion = Nothing
        
        If Me.txtFacNombre.Text = "" Then
            Me.txtFacNombre.Text = Me.lblNombreUsu
        End If
        
        Screen.MousePointer = vbHourglass
        
        Screen.MousePointer = vbHourglass
        
        'Verifica los tipos de usuarios antes de facturar
        'por si su tipo cambio por la edad.
        ChecaTipoUsuario (lIdTitular)
        
        'Calcula cuotas de mantenimiento
        CalculaMantenimiento lidMember, Me.cmbPeriodoCalculo.ItemData(Me.cmbPeriodoCalculo.ListIndex), sCargos, lIndexsCargos
        
        'Calcula cuotas rentables
        CalculaRentables (lidMember)
        
        'Calcula cargos varios
        CalculaCargosVarios (lidMember), (Me.dtpFechaCalc.Value)
        
        'Calcula cargos de membresias
        CalculaCargosMembresia (lidMember), (Me.dtpFechaCalc.Value)
        
        If lIndexsCargos > 1 Then
            SortArray sCargos, lIndexsCargos - 1, False
        End If
        
        For i = 0 To lIndexsCargos - 1
            Me.ssdbgFactura.AddItem Mid(sCargos(i), 12)
        Next
        
        Me.ssdbgFactura.Update
        If Me.ssdbgFactura.Rows > 0 Then
            Me.ssdbgFactura.Bookmark = Me.ssdbgFactura.AddItemBookmark(0)
        End If
        
        Me.lblSubTotal = Format(siTotImporte, "#,##0.00")
        Me.lblDescuento = Format(siTotDescuento, "#,##0.00")
        Me.lblTotal = Format(siTotTotal, "#,##0.00")
        
        siTotalPagado = 0
        Me.lblPagado = Format(siTotalPagado, "#,##0.00")
        
        siTotalPorPagar = siTotTotal
        Me.lblPorPagar = Format(siTotalPorPagar, "#,##0.00")
        
        Me.lblRen = Format(Me.ssdbgFactura.Rows, "#0") & " Renglon(es)"
        
        If Me.ssdbgFactura.Rows = 0 Then
            MsgBox " No existen adeudos pendientes!", vbInformation, "Facturacion"
            Me.txtClave.SelStart = 0
            Me.txtClave.SelLength = Len(Me.txtClave.Text)
            Me.txtClave.SetFocus
        End If
        
        Me.chkDatosFac.Enabled = True
        
    End If

    'Checa si no tiene Mensajes
    
    ChecaMensajes (lIdTitular)
    
    Screen.MousePointer = vbDefault
    
    boYaFacturado = True
End Sub

Private Sub cmdCuponesDesc_Click()
    
    Dim lNumRen As Long
    
    Dim frmCupon As frmCuponDesc
    
    If Me.lblNombreUsu.Caption = "" Then
        MsgBox "No se ha selecionado ningún usuario!", vbExclamation, "Facturación"
        Exit Sub
    End If
    
    lNumRen = Me.ssdbgFactura.Rows
    
    
    Set frmCupon = New frmCuponDesc
    
    frmCupon.dMontoTotal = siTotalPorPagar
    
    frmCupon.Show vbModal
    
    If Me.ssdbgFactura.Rows > lNumRen Then
        Me.ssdbgFactura.Bookmark = Me.ssdbgFactura.GetBookmark(Me.ssdbgFactura.Rows - 1)
        If Me.ssdbgFactura.Rows > Me.ssdbgFactura.VisibleRows Then
            Me.ssdbgFactura.Scroll 0, 1
        End If
        
        RecalculaRenglon
        CalculaTotales
        
        siTotalPorPagar = siTotTotal - siTotalPagado
    
    
        Me.lblPagado = Format(siTotalPagado, "#,##0.00")
        Me.lblPorPagar = Format(siTotalPorPagar, "#,##0.00")
        
        Me.ssdbgFactura.Bookmark = Me.ssdbgFactura.Rows - 1
        
    End If
    
    
End Sub

Private Sub cmdDel_Click()
    
    Dim bTodos As Boolean
    Dim lCurrentRow As Long
    Dim vBookMark As Variant
    
    
    If Me.ssdbgFactura.Rows < 1 Then Exit Sub
    
    If Me.chkEliminaTodos.Value Then
        If MsgBox("¿Desea eliminar todos los registros?", vbQuestion Or vbYesNo, "Facturación") = vbNo Then
            Me.chkEliminaTodos.Value = 0
            Exit Sub
        End If
        bTodos = True
        Me.chkEliminaTodos.Value = 0
    End If
    
    
    lCurrentRow = Me.ssdbgFactura.AddItemRowIndex(Me.ssdbgFactura.Bookmark)
    
    Select Case lCurrentRow
        'Si es el ultimo renglon
        Case Is = Me.ssdbgFactura.Rows - 1
            vBookMark = Me.ssdbgFactura.AddItemBookmark(Me.ssdbgFactura.Rows - 2)
        'Si es un renglon intermedio
        Case Is < Me.ssdbgFactura.Rows - 1
            vBookMark = Me.ssdbgFactura.AddItemBookmark(lCurrentRow + 1)
        Case Else
            vBookMark = Me.ssdbgFactura.Bookmark
    End Select
    
    
    If bTodos Then
        Me.ssdbgFactura.RemoveAll
    Else
        Me.ssdbgFactura.RemoveItem (Me.ssdbgFactura.AddItemRowIndex(Me.ssdbgFactura.Bookmark))
    End If
    
    Me.ssdbgFactura.Update
    
    Select Case Me.ssdbgFactura.Rows
        Case 0
            Me.ssdbgFactura.Bookmark = Me.ssdbgFactura.AddItemBookmark(0)
        Case 1
            Me.ssdbgFactura.Bookmark = Me.ssdbgFactura.AddItemBookmark(0)
        Case Else
            Select Case lCurrentRow
                Case Is > (Me.ssdbgFactura.Rows - 1)
                    Me.ssdbgFactura.Bookmark = Me.ssdbgFactura.AddItemBookmark(Me.ssdbgFactura.Rows - 1)
                Case Is < (Me.ssdbgFactura.Rows - 1)
                    Me.ssdbgFactura.Bookmark = Me.ssdbgFactura.AddItemBookmark(lCurrentRow)
            End Select
    End Select
        
    
'    If Me.ssdbgFactura.Rows > 0 Then
'        Me.ssdbgFactura.Bookmark = vBookMark
'    Else
'        Me.ssdbgFactura.Bookmark = Me.ssdbgFactura.AddItemBookmark(0)
'    End If
'
    If Me.ssdbgFactura.Rows <= Me.ssdbgFactura.VisibleRows Then
        Me.ssdbgFactura.FirstRow = Me.ssdbgFactura.AddItemBookmark(0)
    End If
        
    
    CalculaTotales
    
    siTotalPorPagar = siTotTotal - siTotalPagado
    
    
    Me.lblPagado = Format(siTotalPagado, "#,##0.00")
    Me.lblPorPagar = Format(siTotalPorPagar, "#,##0.00")
    
    
End Sub

Private Sub cmdDescuento_Click()
    
    If Me.ssdbgFactura.Rows < 1 Then Exit Sub
    
    If Not ChecaSeguridad(Me.Name, Me.cmdDescuento.Name) Then
        Exit Sub
    End If
    
    
    frmDescFac.Show 1
    CalculaTotales
    
    siTotalPorPagar = siTotTotal - siTotalPagado
    
    
    Me.lblPagado = Format(siTotalPagado, "#,##0.00")
    Me.lblPorPagar = Format(siTotalPorPagar, "#,##0.00")
End Sub

Private Sub cmdEliminaPago_Click()
    If Me.ssdbgPagos.Rows < 1 Then
        Exit Sub
    End If
    
        
    siTotalPagado = siTotalPagado - Me.ssdbgPagos.Columns("Importe").Value
    siTotalPorPagar = siTotTotal - siTotalPagado
    
    Me.ssdbgPagos.RemoveItem Me.ssdbgPagos.AddItemRowIndex(Me.ssdbgPagos.Bookmark)
    
    Me.lblPagado = Format(siTotalPagado, "#,##0.00")
    Me.lblPorPagar = Format(siTotalPorPagar, "#,##0.00")
    
    If Me.ssdbgPagos.Rows Then
        Me.ssdbgPagos.Caption = "Pagos " & "(" & Me.ssdbgPagos.Rows & ")"
        
        If Me.ssdbgPagos.Rows <= Me.ssdbgPagos.VisibleRows Then
            Me.ssdbgPagos.FirstRow = 0
        End If
        
        Me.ssdbgPagos.Bookmark = Me.ssdbgPagos.AddItemBookmark(Me.ssdbgPagos.Rows - 1)
    Else
        Me.ssdbgPagos.Caption = "Pagos"
    End If
    
    
    
End Sub

Private Sub cmdEstacionamiento_Click()
    
    Dim lNumRen As Long
    
    Dim frmPark As frmEstacionamiento
    
    
    If Me.lblNombreUsu.Caption = "" Then
        MsgBox "No se ha selecionado ningún usuario!", vbExclamation, "Facturación"
        Exit Sub
    End If
    
    lNumRen = Me.ssdbgFactura.Rows
    
    
    Set frmPark = New frmEstacionamiento
    
    frmEstacionamiento.Show vbModal
    
    
    If Me.ssdbgFactura.Rows > lNumRen Then
        Me.ssdbgFactura.Bookmark = Me.ssdbgFactura.GetBookmark(Me.ssdbgFactura.Rows - 1)
        If Me.ssdbgFactura.Rows > Me.ssdbgFactura.VisibleRows Then
            Me.ssdbgFactura.Scroll 0, 1
        End If
        
        RecalculaRenglon
        CalculaTotales
        
        siTotalPorPagar = siTotTotal - siTotalPagado
    
    
        Me.lblPagado = Format(siTotalPagado, "#,##0.00")
        Me.lblPorPagar = Format(siTotalPorPagar, "#,##0.00")
        
        Me.ssdbgFactura.Bookmark = Me.ssdbgFactura.Rows - 1
        
    End If
    
    
    
    
End Sub

Private Sub cmdFacturar_Click()

    Dim lNumeroFacturaInicial As Long
    Dim lNumeroFacturaFinal As Long
    Dim lNumeroFactura As Long
    Dim lRenMaxFactura As Long          'Numero maximo de renglones
    Dim lNumerodeFacturas As Long       'Numero de facturas que se generaran
    
    
    '29/06/07
    Dim lNumeroFolioFacturaInicial As Long
    Dim lNumeroFolioFacturaFinal As Long
    Dim lNumeroFolioFactura As Long
    
    Dim lNumeroReciboInicial As Long
    Dim lNumeroReciboFinal  As Long
    Dim lNumeroRecibo As Long
    Dim lNumerodeRecibos As Long
    
    '29/06/07
    Dim lNumeroFolioReciboInicial As Long
    Dim lNumeroFolioReciboFinal As Long
    Dim lNumeroFolioRecibo As Long
    
    Dim sMensaje As String
    Dim iRespuesta As Integer
    
    Dim lFolio As Long
    Dim sSerie As String
    
    '23/07/2006
    Dim lTurno As Long
    
    Dim lIHeader As Long
    Dim lIDetalle As Long
    Dim lRowPointer As Long
    Dim lRowEnd As Long
    
    Dim lRowCount As Long
    
    Dim dTotalFactura As Double
    Dim dTotPagPorFactura As Double
    Dim lRowPagos As Long
    Dim iPagos As Long
    Dim dPagoFac As Double
    Dim dCompara As Double
    
    Dim lRenFac As Long
    Dim lRenRec As Long
    
    Dim sTipoDireccionado As String
    
    Dim sTipoPersona As String
    
    Dim adocmdFactura As ADODB.Command
    
    Dim iInitTrans As Integer
    
    Dim iResp As Integer
    
    Dim sFolioCFD As String
    Dim sSerieCFD As String
    Dim sNombreArcCfd As String
    
    Dim lNumeroCupones As Long
    
    Dim lHayMantenimiento As Boolean
    
    lPublicoGeneral = False
    
    lHayMantenimiento = False
    
    
    If Not ChecaSeguridad(Me.Name, Me.cmdFacturar.Name) Then
        Exit Sub
    End If
    
    Me.cmdFacturar.Enabled = False
    
    
    lRenMaxFactura = 300
    
    sTipoPersona = "F"
    
    sFolioCFD = vbNullString
    sSerieCFD = vbNullString
    sNombreArcCfd = vbNullString
    
    
    If Me.lblNombreUsu.Caption = "" Then
        MsgBox "No se ha selecionado ningún usuario!", vbExclamation, "Facturación"
        Me.cmdFacturar.Enabled = True
        Exit Sub
    End If
    
    
    If Round(siTotalPorPagar, 2) > 0 Then
        MsgBox " Faltan por pagar " & Format(siTotalPorPagar, "$#,##0.00"), vbExclamation, "Facturación"
        Me.cmdFacturar.Enabled = True
        Exit Sub
    End If
    
    If Round(siTotalPorPagar, 2) < 0 Then
        MsgBox "Se esta pagando de más " & Format(siTotalPorPagar * -1, "$#,##0.00"), vbExclamation, "Facturación"
        Me.cmdFacturar.Enabled = True
        Exit Sub
    End If
    
    If Me.chkDireccionar.Value And Me.cmbTipoDireccionar.Text = "" Then
        MsgBox "No se ha elegido la forma de direccionamiento!", vbExclamation, "Facturación"
        Me.cmdFacturar.Enabled = True
        Exit Sub
    End If
    
    If Me.txtFacNombre.Text = vbNullString Then
        MsgBox "El campo Nombre no puede quedar en blanco", vbExclamation, "Verifique"
        Me.cmdFacturar.Enabled = True
        Exit Sub
    End If
    
    If Me.txtFacDireccion.Text = vbNullString Then
        MsgBox "El campo Dirección no puede quedar en blanco", vbExclamation, "Verifique"
        Me.cmdFacturar.Enabled = True
        Exit Sub
    End If
    
    If Me.txtFacColonia.Text = vbNullString Then
        MsgBox "El campo Colonia no puede quedar en blanco", vbExclamation, "Verifique"
        Me.cmdFacturar.Enabled = True
        Exit Sub
    End If
    
    If Me.txtFacDelOMuni.Text = vbNullString Then
        MsgBox "El campo Delegación/Municipio no puede quedar en blanco", vbExclamation, "Verifique"
        Me.cmdFacturar.Enabled = True
        Exit Sub
    End If
    
    If Me.txtFacCiudad.Text = vbNullString Then
        MsgBox "El campo Ciudad no puede quedar en blanco", vbExclamation, "Verifique"
        Me.cmdFacturar.Enabled = True
        Exit Sub
    End If
    
    If Me.txtFacCP.Text = vbNullString Then
        Me.cmdFacturar.Enabled = True
        MsgBox "El campo Código Postal no puede quedar en blanco", vbExclamation, "Verifique"
        Exit Sub
    End If
    
    If Me.ssCmbFacEstado.Text = vbNullString Then
        Me.cmdFacturar.Enabled = True
        MsgBox "El campo Estado no puede quedar en blanco", vbExclamation, "Verifique"
        Exit Sub
    End If
    
    If Me.optTipoPer(0).Value Then 'Persona física
        If (Len(Me.txtFacRFC) <> 13) Then
            iResp = MsgBox("El RFC debe ser de 13 caracteres para personas físicas" & vbCrLf & "¿Desea emitir la factura para Publico en General?", vbYesNo + vbQuestion, "Confirme")
            If iResp = vbNo Then
                Me.cmdFacturar.Enabled = True
                Exit Sub
            End If
            lPublicoGeneral = True
        End If
    Else 'Persona Moral
        If (Len(Me.txtFacRFC) <> 12) Then
            iResp = MsgBox("El RFC debe ser de 12 caracteres para personas morales" & vbCrLf & "¿Desea emitir la factura para Publico en General?", vbYesNo + vbQuestion, "Confirme")
            If iResp = vbNo Then
                Me.cmdFacturar.Enabled = True
                Exit Sub
            End If
            lPublicoGeneral = True
        End If
    End If
        
    If Me.optTipoPer(0).Value Then
        sTipoPersona = "F"
    Else
        sTipoPersona = "M"
    End If
    
    If lPublicoGeneral Then
        Me.txtFacRFC.Text = "XAXX010101000"
    End If

    lTurno = OpenShiftF()
    
    If lTurno = 0 Then
        MsgBox "No hay turno abierto!", vbCritical, "Verifique"
        Me.cmdFacturar.Enabled = True
        Exit Sub
    End If
    
    If Me.txtObserva.Text = "" Then
        MsgBox "Faltan las observaciones", vbCritical, "Verifique"
        Me.cmdFacturar.Enabled = True
        Me.txtObserva.SetFocus
        Exit Sub
    End If
    
    lRenFac = 0
    lRenRec = 0
    
    For lIHeader = 0 To Me.ssdbgFactura.Rows - 1
        Me.ssdbgFactura.Bookmark = Me.ssdbgFactura.AddItemBookmark(lIHeader)
        If Me.ssdbgFactura.Columns("FacORec").Value = "F" Then
            lRenFac = lRenFac + 1
        Else
            lRenRec = lRenRec + 1
        End If
    Next
    
    lNumerodeFacturas = Int(lRenFac / lRenMaxFactura) + IIf(lRenFac Mod lRenMaxFactura, 1, 0)
    
    lNumerodeRecibos = Int(lRenRec / lRenMaxFactura) + IIf(lRenRec Mod lRenMaxFactura, 1, 0)
    
    If lNumerodeFacturas Then           'Si hay facturas
        If lNumerodeRecibos Then        'Facturas y recibos
            iRespuesta = MsgBox("Se generara(n) " & Format(lNumerodeFacturas, "##0") & " Factura(s) " & Chr(13) & "y " & Format(lNumerodeRecibos, "##0") & " Recibo(s) ", vbQuestion Or vbOKCancel, "Facturacion")
        Else                            'Solo facturas
            iRespuesta = MsgBox("Se generara(n) " & Format(lNumerodeFacturas, "##0") & " Factura(s) ", vbQuestion Or vbOKCancel, "Facturacion")
        End If
    Else                                'Si no hay facturas
        If lNumerodeRecibos Then          'Solo recibos
            iRespuesta = MsgBox("Se generara(n) " & Format(lNumerodeRecibos, "##0") & " Recibo(s) ", vbQuestion Or vbOKCancel, "Facturacion")
        Else
        End If
    End If
    
    If iRespuesta = 2 Then
        Me.cmdFacturar.Enabled = True
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    If Me.chkDireccionar.Value Then
        If Me.cmbTipoDireccionar.ListIndex = 0 Then
            sTipoDireccionado = "TC"
        Else
            sTipoDireccionado = "TD"
        End If
    Else
        sTipoDireccionado = ""
    End If
    
    MDIPrincipal.StatusBar1.Panels(1).Text = "Obteniendo folios"
    
    lNumeroFacturaInicial = GetFolio(lNumerodeFacturas, 0)
    lNumeroReciboInicial = GetFolio(lNumerodeRecibos, 1)
    
    '29/06/07
    lNumeroFolioFacturaInicial = GetFolioSerie(lNumerodeFacturas, sSerieFactura)
    lNumeroFolioReciboInicial = lNumeroReciboInicial
    
    If lNumerodeFacturas > 0 And lNumeroFacturaInicial = -1 Then
        Screen.MousePointer = vbDefault
        MsgBox "Error al obtener folio, reintente", vbCritical
        Exit Sub
    End If
    
    If lNumerodeRecibos > 0 And lNumeroReciboInicial = -1 Then
        Screen.MousePointer = vbDefault
        MsgBox "Error al obtener folio, reintente", vbCritical
        Exit Sub
    End If
    
    MDIPrincipal.StatusBar1.Panels(1).Text = "Guardando Factura(s)"
    
    Err.Clear
    Conn.Errors.Clear
    On Error GoTo Error_Catch
    iInitTrans = Conn.BeginTrans
    
    lFolio = 1
    sSerie = vbNullString
    
    lNumeroFactura = lNumeroFacturaInicial
    
    '29/06/07
    lNumeroFolioFactura = lNumeroFolioFacturaInicial
    
    Set adocmdFactura = New ADODB.Command
    adocmdFactura.ActiveConnection = Conn
    adocmdFactura.CommandType = adCmdText
    
    
    lRowPointer = 0
    
    For lIHeader = 1 To lNumerodeFacturas
    
        dTotalFactura = 0
        
        #If SqlServer_ Then
            strSQL = vbNullString
            strSQL = "SET DATEFORMAT dmy" & vbCrLf
        #Else
            strSQL = vbNullString
        #End If
    
        strSQL = strSQL & "INSERT INTO FACTURAS"
        strSQL = strSQL & " ( NumeroFactura,"
        strSQL = strSQL & " Folio,"
        strSQL = strSQL & " Serie,"
        strSQL = strSQL & " IdTitular,"
        strSQL = strSQL & " NoFamilia,"
        strSQL = strSQL & " FechaFactura,"
        strSQL = strSQL & " HoraFactura,"
        strSQL = strSQL & " NombreFactura,"
        strSQL = strSQL & " CalleFactura,"
        strSQL = strSQL & " ColoniaFactura,"
        strSQL = strSQL & " DelFactura,"
        strSQL = strSQL & " CiudadFactura,"
        strSQL = strSQL & " EstadoFactura,"
        strSQL = strSQL & " CodPos,"
        strSQL = strSQL & " RFC,"
        strSQL = strSQL & " Tel1,"
        strSQL = strSQL & " Observaciones,"
        strSQL = strSQL & " ImporteConLetra,"
        strSQL = strSQL & " Usuario,"
        '23/07/2006
        strSQL = strSQL & " Turno,"
        '29/06/2006
        strSQL = strSQL & " Caja,"
        strSQL = strSQL & " Direccionado,"
        strSQL = strSQL & " TipoPersona)"
        strSQL = strSQL & " VALUES ("
        strSQL = strSQL & lNumeroFactura & ","
        strSQL = strSQL & lNumeroFolioFactura & ","
        strSQL = strSQL & "'" & sSerieFactura & "', "
        strSQL = strSQL & lIdTitular & ", "
        strSQL = strSQL & lIdFamilia & ", "
        strSQL = strSQL & "'" & Format(Now, "dd/mm/yyyy") & "', "
        strSQL = strSQL & "'" & Format(Now, "Hh:Nn:Ss") & "', "
        strSQL = strSQL & "'" & Trim(Me.txtFacNombre.Text) & "', "
        strSQL = strSQL & "'" & Trim(Me.txtFacDireccion.Text) & "', "
        strSQL = strSQL & "'" & Trim(Me.txtFacColonia.Text) & "', "
        strSQL = strSQL & "'" & Trim(Me.txtFacDelOMuni.Text) & "', "
        strSQL = strSQL & "'" & Trim(Me.txtFacCiudad.Text) & "', "
        strSQL = strSQL & "'" & Trim(Me.ssCmbFacEstado.Text) & "', "
        strSQL = strSQL & "'" & Trim(Me.txtFacCP.Text) & "', "
        strSQL = strSQL & "'" & Trim(Me.txtFacRFC.Text) & "', "
        strSQL = strSQL & "'" & Trim(Me.txtFacTelefono.Text) & "', "
        strSQL = strSQL & "'" & Trim(Me.txtObserva.Text) & "', "
        strSQL = strSQL & "'" & Trim(vbNullString) & "',"
        strSQL = strSQL & "'" & Trim(sDB_User) & "',"
        '23/07/2006
        strSQL = strSQL & lTurno & ","
        '29/06/2007
        strSQL = strSQL & iNumeroCaja & ","
        strSQL = strSQL & "'" & sTipoDireccionado & "',"
        strSQL = strSQL & "'" & sTipoPersona & "')"
        
        adocmdFactura.CommandText = strSQL
        adocmdFactura.Execute
    
        lRowEnd = lRenMaxFactura * lIHeader - 1
        
        If lRowEnd > Me.ssdbgFactura.Rows - 1 Then
            lRowEnd = Me.ssdbgFactura.Rows - 1
        End If
        
        lRowCount = 1
    
        For lIDetalle = lRowPointer To Me.ssdbgFactura.Rows - 1
        
            If lRowCount > lRenMaxFactura Then
                Exit For
            End If
            
            Me.ssdbgFactura.Bookmark = Me.ssdbgFactura.AddItemBookmark(lIDetalle)
    
            If Me.ssdbgFactura.Columns("FacORec").Value = "F" Then
                #If SqlServer_ Then
                    strSQL = vbNullString
                    strSQL = "SET DATEFORMAT dmy" & vbCrLf
                #Else
                    strSQL = vbNullString
                #End If
                
                strSQL = strSQL & "INSERT INTO FACTURAS_DETALLE"
                strSQL = strSQL & " (NumeroFactura, "
                strSQL = strSQL & " Renglon, "
                strSQL = strSQL & " IdConcepto, "
                strSQL = strSQL & " IdMember, "
                strSQL = strSQL & " NumeroFamiliar, "
                strSQL = strSQL & " IdTipoUsuario, "
                strSQL = strSQL & " Periodo, "
                strSQL = strSQL & " FormaPago, "
                strSQL = strSQL & " Concepto, "
                strSQL = strSQL & " Cantidad, "
                strSQL = strSQL & " Importe, "
                strSQL = strSQL & " Intereses, "
                strSQL = strSQL & " DescuentoPorciento, "
                strSQL = strSQL & " Descuento, "
                strSQL = strSQL & " Total, "
                strSQL = strSQL & " IvaPorciento, "
                strSQL = strSQL & " Iva, "
                strSQL = strSQL & " IvaIntereses, "
                strSQL = strSQL & " IvaDescuento,"
                strSQL = strSQL & " TipoCargo,"
                strSQL = strSQL & " Auxiliar,"
                strSQL = strSQL & " IdInstructor,"
                strSQL = strSQL & " Unidad)"
                strSQL = strSQL & " VALUES ("
                strSQL = strSQL & lNumeroFactura & ","
                strSQL = strSQL & lRowCount & ","
                strSQL = strSQL & Me.ssdbgFactura.Columns("Clave").Value & ","
                strSQL = strSQL & Me.ssdbgFactura.Columns("IdMember").Value & ","
                strSQL = strSQL & Me.ssdbgFactura.Columns("NoFamiliar").Value & ","
                strSQL = strSQL & Me.ssdbgFactura.Columns("IdTipoUsuario").Value & ","
                strSQL = strSQL & "'" & Me.ssdbgFactura.Columns("Periodo").Value & "',"
                strSQL = strSQL & Me.ssdbgFactura.Columns("FormaPago").Value & ","
                strSQL = strSQL & "'" & Left$(Me.ssdbgFactura.Columns("Concepto").Value, 50) & "',"
                strSQL = strSQL & Me.ssdbgFactura.Columns("Cantidad").Value & ","
                strSQL = strSQL & Me.ssdbgFactura.Columns("Importe").Value & ","
                strSQL = strSQL & Me.ssdbgFactura.Columns("Intereses").Value & ","
                strSQL = strSQL & Me.ssdbgFactura.Columns("Descuento").Value & ","
                strSQL = strSQL & Me.ssdbgFactura.Columns("DescMonto").Value & ","
                strSQL = strSQL & Me.ssdbgFactura.Columns("Total").Value & ","
                strSQL = strSQL & Me.ssdbgFactura.Columns("IvaPor").Value & ","
                strSQL = strSQL & Me.ssdbgFactura.Columns("Iva").Value & ","
                strSQL = strSQL & Me.ssdbgFactura.Columns("IvaIntereses").Value & ","
                strSQL = strSQL & Me.ssdbgFactura.Columns("IvaDescuento").Value & ","
                strSQL = strSQL & Me.ssdbgFactura.Columns("TipoCargo").Value & ","
                strSQL = strSQL & "'" & Me.ssdbgFactura.Columns("Auxiliar").Value & "',"
                strSQL = strSQL & Val(Me.ssdbgFactura.Columns("IdInstructor").Value) & ","
                strSQL = strSQL & "'" & Me.ssdbgFactura.Columns("Unidad").Value & "')"
        
                adocmdFactura.CommandText = strSQL
                adocmdFactura.Execute
            
                dTotalFactura = dTotalFactura + Round(CDbl(Me.ssdbgFactura.Columns("Total").Value), 2)
                
                 If Me.ssdbgFactura.Columns("Clave").Value > 900 Then
                    lHayMantenimiento = True
                 End If
                
                lRowCount = lRowCount + 1
            End If
            
            lRowPointer = lRowPointer + 1
            
        Next
        
                
        lRowPagos = 1
        dTotPagPorFactura = 0
        Do Until Round(dTotPagPorFactura - dTotalFactura, 2) >= 0
        
            'Salta los que estan en cero
            For iPagos = 0 To Me.ssdbgPagos.Rows - 1
                Me.ssdbgPagos.Bookmark = Me.ssdbgPagos.AddItemBookmark(iPagos)
                If CDbl(Me.ssdbgPagos.Columns("Resta").Value) > 0 Then
                    Exit For
                End If
            Next
            
            Select Case Round(dTotalFactura - dTotPagPorFactura - CDbl(Me.ssdbgPagos.Columns("Resta").Value), 2)
                Case Is > 0 'No alcanza a cubrir
                    dPagoFac = CDbl(Me.ssdbgPagos.Columns("Resta").Value)
                    Me.ssdbgPagos.Columns("Resta").Value = 0
                Case Is = 0 'Cubre exacto
                    dPagoFac = CDbl(Me.ssdbgPagos.Columns("Resta").Value)
                    Me.ssdbgPagos.Columns("Resta").Value = 0
                Case Is < 0
                    dPagoFac = Round(dTotalFactura - dTotPagPorFactura, 2)
                    Me.ssdbgPagos.Columns("Resta").Value = Round(CDbl(Me.ssdbgPagos.Columns("Resta").Value) - dPagoFac, 2)
            End Select
            
            Me.ssdbgPagos.Update
        
            strSQL = "INSERT INTO PAGOS_FACTURA ("
            strSQL = strSQL & " NumeroFactura, "
            strSQL = strSQL & " Renglon, "
            strSQL = strSQL & " IdFormaPago, "
            strSQL = strSQL & " OpcionPago, "
            strSQL = strSQL & " Importe, "
            strSQL = strSQL & " Referencia, "
            strSQL = strSQL & " IdAfiliacion, "
            strSQL = strSQL & " LoteNumero, "
            strSQL = strSQL & " OperacionNumero, "
            strSQL = strSQL & " ImporteRecibido, "
            strSQL = strSQL & " FechaOperacion, "
            strSQL = strSQL & " NumeroCuenta) "
            strSQL = strSQL & " VALUES ("
            strSQL = strSQL & lNumeroFactura & ", "
            strSQL = strSQL & lRowPagos & ", "
            strSQL = strSQL & Me.ssdbgPagos.Columns("IdTipoPago").Value & ","
            strSQL = strSQL & "'" & Me.ssdbgPagos.Columns("Opcion").Value & "'" & ","
            strSQL = strSQL & dPagoFac & ","
            strSQL = strSQL & "'" & Me.ssdbgPagos.Columns("Referencia").Value & "',"
            strSQL = strSQL & IIf(Me.ssdbgPagos.Columns("IdAfiliacion").Value = vbNullString, 0, Me.ssdbgPagos.Columns("IdAfiliacion").Value) & ","
            strSQL = strSQL & "'" & Me.ssdbgPagos.Columns("LoteNumero").Value & "',"
            strSQL = strSQL & "'" & Me.ssdbgPagos.Columns("OperacionNumero").Value & "',"
            strSQL = strSQL & Me.ssdbgPagos.Columns("ImporteRecibido").Value & ","
            strSQL = strSQL & "'" & Me.ssdbgPagos.Columns("FechaOperacion").Value & "',"
            If Me.ssdbgPagos.Columns("NoCuenta").Value = vbNullString Then
                strSQL = strSQL & "Null" & ")"
            Else
                strSQL = strSQL & "'" & Me.ssdbgPagos.Columns("NoCuenta").Value & "')"
            End If
            adocmdFactura.CommandText = strSQL
            adocmdFactura.Execute
        
            dTotPagPorFactura = dTotPagPorFactura + dPagoFac
            lRowPagos = lRowPagos + 1
        
        Loop
        
        
        'Actualiza el total de la factura
        strSQL = "UPDATE FACTURAS SET Total=" & dTotalFactura
        strSQL = strSQL & " WHERE NumeroFactura=" & lNumeroFactura
        
        adocmdFactura.CommandText = strSQL
        adocmdFactura.Execute
        
        
        lNumeroFactura = lNumeroFactura + 1
        lNumeroFolioFactura = lNumeroFolioFactura + 1
   Next
   
   MDIPrincipal.StatusBar1.Panels(1).Text = "Guardando Recibo(s)"
   
'Para generar recibos
'------------------------------------------------------------------------------
    lFolio = 1
    sSerie = vbNullString
    
    lNumeroRecibo = lNumeroReciboInicial
    
    '29/06/07
    
    
    Set adocmdFactura = New ADODB.Command
    adocmdFactura.ActiveConnection = Conn
    adocmdFactura.CommandType = adCmdText
    
    lRowPointer = 0
    
    For lIHeader = 1 To lNumerodeRecibos
    
        dTotalFactura = 0
        
        #If SqlServer_ Then
            strSQL = vbNullString
            strSQL = "SET DATEFORMAT dmy" & vbCrLf
        #Else
            strSQL = vbNullString
        #End If
    
        strSQL = strSQL & "INSERT INTO RECIBOS"
        strSQL = strSQL & " ( NumeroRecibo,"
        strSQL = strSQL & " Folio,"
        strSQL = strSQL & " Serie,"
        strSQL = strSQL & " IdTitular,"
        strSQL = strSQL & " NoFamilia,"
        strSQL = strSQL & " FechaFactura,"
        strSQL = strSQL & " HoraFactura,"
        strSQL = strSQL & " NombreFactura,"
        strSQL = strSQL & " CalleFactura,"
        strSQL = strSQL & " ColoniaFactura,"
        strSQL = strSQL & " DelFactura,"
        strSQL = strSQL & " CiudadFactura,"
        strSQL = strSQL & " EstadoFactura,"
        strSQL = strSQL & " CodPos,"
        strSQL = strSQL & " RFC,"
        strSQL = strSQL & " Tel1,"
        strSQL = strSQL & " Observaciones,"
        strSQL = strSQL & " ImporteConLetra,"
        strSQL = strSQL & " Usuario,"
        '23/07/2006
        strSQL = strSQL & " Turno,"
        '29/06/2007
        strSQL = strSQL & " Caja,"
        strSQL = strSQL & " Direccionado)"
        strSQL = strSQL & " VALUES ("
        strSQL = strSQL & lNumeroRecibo & ","
        strSQL = strSQL & lNumeroRecibo & ","
        strSQL = strSQL & "'" & sSerie & "', "
        strSQL = strSQL & lIdTitular & ", "
        strSQL = strSQL & lIdFamilia & ", "
        strSQL = strSQL & "'" & Format(Now, "dd/mm/yyyy") & "', "
        strSQL = strSQL & "'" & Format(Now, "Hh:Nn:Ss") & "', "
        strSQL = strSQL & "'" & Trim(Me.txtFacNombre.Text) & "', "
        strSQL = strSQL & "'" & Trim(Me.txtFacDireccion.Text) & "', "
        strSQL = strSQL & "'" & Trim(Me.txtFacColonia.Text) & "', "
        strSQL = strSQL & "'" & Trim(Me.txtFacDelOMuni.Text) & "', "
        strSQL = strSQL & "'" & Trim(Me.txtFacCiudad.Text) & "', "
        strSQL = strSQL & "'" & Trim(Me.ssCmbFacEstado.Text) & "', "
        strSQL = strSQL & "'" & Trim(Me.txtFacCP.Text) & "', "
        strSQL = strSQL & "'" & Trim(Me.txtFacRFC.Text) & "', "
        strSQL = strSQL & "'" & Trim(Me.txtFacTelefono.Text) & "', "
        strSQL = strSQL & "'" & Trim(Me.txtObserva.Text) & "', "
        strSQL = strSQL & "'" & Trim(vbNullString) & "',"
        strSQL = strSQL & "'" & Trim(sDB_User) & "',"
        '23/07/2006
        strSQL = strSQL & lTurno & ","
        '29/06/2007
        strSQL = strSQL & iNumeroCaja & ","
        strSQL = strSQL & "'" & sTipoDireccionado & "')"
    
        adocmdFactura.CommandText = strSQL
        adocmdFactura.Execute
    
        lRowEnd = lRenMaxFactura * lIHeader - 1
        
        If lRowEnd > Me.ssdbgFactura.Rows - 1 Then
            lRowEnd = Me.ssdbgFactura.Rows - 1
        End If
        
        lRowCount = 1
    
        For lIDetalle = lRowPointer To Me.ssdbgFactura.Rows - 1
            
            If lRowCount > lRenMaxFactura Then
                Exit For
            End If
    
            Me.ssdbgFactura.Bookmark = Me.ssdbgFactura.AddItemBookmark(lIDetalle)
            
            If Me.ssdbgFactura.Columns("FacORec").Value = "R" Then
                #If SqlServer_ Then
                    strSQL = vbNullString
                    strSQL = "SET DATEFORMAT dmy" & vbCrLf
                #Else
                    strSQL = vbNullString
                #End If
                
                strSQL = strSQL & "INSERT INTO RECIBOS_DETALLE"
                strSQL = strSQL & " (NumeroRecibo, "
                strSQL = strSQL & " Renglon, "
                strSQL = strSQL & " IdConcepto, "
                strSQL = strSQL & " IdMember, "
                strSQL = strSQL & " NumeroFamiliar, "
                strSQL = strSQL & " IdTipoUsuario, "
                strSQL = strSQL & " Periodo, "
                strSQL = strSQL & " FormaPago, "
                strSQL = strSQL & " Concepto, "
                strSQL = strSQL & " Cantidad, "
                strSQL = strSQL & " Importe, "
                strSQL = strSQL & " Intereses, "
                strSQL = strSQL & " DescuentoPorciento, "
                strSQL = strSQL & " Descuento, "
                strSQL = strSQL & " Total, "
                strSQL = strSQL & " IvaPorciento, "
                strSQL = strSQL & " Iva, "
                strSQL = strSQL & " IvaIntereses, "
                strSQL = strSQL & " IvaDescuento,"
                strSQL = strSQL & " TipoCargo,"
                strSQL = strSQL & " Auxiliar,"
                strSQL = strSQL & " IdInstructor,"
                strSQL = strSQL & " Unidad)"
                strSQL = strSQL & " VALUES ("
                strSQL = strSQL & lNumeroRecibo & ","
                strSQL = strSQL & lRowCount & ","
                strSQL = strSQL & Me.ssdbgFactura.Columns("Clave").Value & ","
                strSQL = strSQL & Me.ssdbgFactura.Columns("IdMember").Value & ","
                strSQL = strSQL & Me.ssdbgFactura.Columns("NoFamiliar").Value & ","
                strSQL = strSQL & Me.ssdbgFactura.Columns("IdTipoUsuario").Value & ","
                strSQL = strSQL & "'" & Me.ssdbgFactura.Columns("Periodo").Value & "',"
                strSQL = strSQL & Me.ssdbgFactura.Columns("FormaPago").Value & ","
                strSQL = strSQL & "'" & Me.ssdbgFactura.Columns("Concepto").Value & "',"
                strSQL = strSQL & Me.ssdbgFactura.Columns("Cantidad").Value & ","
                strSQL = strSQL & Me.ssdbgFactura.Columns("Importe").Value & ","
                strSQL = strSQL & Me.ssdbgFactura.Columns("Intereses").Value & ","
                strSQL = strSQL & Me.ssdbgFactura.Columns("Descuento").Value & ","
                strSQL = strSQL & Me.ssdbgFactura.Columns("DescMonto").Value & ","
                strSQL = strSQL & Me.ssdbgFactura.Columns("Total").Value & ","
                strSQL = strSQL & Me.ssdbgFactura.Columns("IvaPor").Value & ","
                strSQL = strSQL & Me.ssdbgFactura.Columns("Iva").Value & ","
                strSQL = strSQL & Me.ssdbgFactura.Columns("IvaIntereses").Value & ","
                strSQL = strSQL & Me.ssdbgFactura.Columns("IvaDescuento").Value & ","
                strSQL = strSQL & Me.ssdbgFactura.Columns("TipoCargo").Value & ","
                strSQL = strSQL & "'" & Me.ssdbgFactura.Columns("Auxiliar").Value & "',"
                strSQL = strSQL & Val(Me.ssdbgFactura.Columns("IdInstructor").Value) & ","
                strSQL = strSQL & "'" & Me.ssdbgFactura.Columns("Unidad").Value & "')"
                adocmdFactura.CommandText = strSQL
                adocmdFactura.Execute
                
                dTotalFactura = dTotalFactura + Round(CDbl(Me.ssdbgFactura.Columns("Total").Value), 2)
                
                lRowCount = lRowCount + 1
            End If
            
            lRowPointer = lRowPointer + 1
            
        Next
        
        lRowPagos = 1
        dTotPagPorFactura = 0
        Do Until Round(dTotPagPorFactura - dTotalFactura, 2) >= 0
            'Salta los que estan en cero
            For iPagos = 0 To Me.ssdbgPagos.Rows - 1
                Me.ssdbgPagos.Bookmark = Me.ssdbgPagos.AddItemBookmark(iPagos)
                If CDbl(Me.ssdbgPagos.Columns("Resta").Value) > 0 Then
                    Exit For
                End If
            Next
            
            Select Case Round(dTotalFactura - dTotPagPorFactura - CDbl(Me.ssdbgPagos.Columns("Resta").Value), 2)
                Case Is > 0 'No alcanza a cubrir
                    dPagoFac = CDbl(Me.ssdbgPagos.Columns("Resta").Value)
                    Me.ssdbgPagos.Columns("Resta").Value = 0
                Case Is = 0 'Cubre exacto
                    dPagoFac = CDbl(Me.ssdbgPagos.Columns("Resta").Value)
                    Me.ssdbgPagos.Columns("Resta").Value = 0
                Case Is < 0
                    dPagoFac = dTotalFactura - dTotPagPorFactura
                    Me.ssdbgPagos.Columns("Resta").Value = Round(CDbl(Me.ssdbgPagos.Columns("Resta").Value) - dPagoFac, 2)
            End Select
            
            Me.ssdbgPagos.Update
        
            strSQL = "INSERT INTO PAGOS_RECIBO"
            strSQL = strSQL & " (NumeroRecibo, "
            strSQL = strSQL & " Renglon, "
            strSQL = strSQL & " IdFormaPago, "
            strSQL = strSQL & " OpcionPago, "
            strSQL = strSQL & " Importe, "
            strSQL = strSQL & " Referencia, "
            strSQL = strSQL & " IdAfiliacion, "
            strSQL = strSQL & " LoteNumero, "
            strSQL = strSQL & " OperacionNumero, "
            strSQL = strSQL & " ImporteRecibido, "
            strSQL = strSQL & " FechaOperacion) "
            strSQL = strSQL & " VALUES ("
            strSQL = strSQL & lNumeroRecibo & ", "
            strSQL = strSQL & lRowPagos & ", "
            strSQL = strSQL & Me.ssdbgPagos.Columns("IdTipoPago").Value & ","
            strSQL = strSQL & "'" & Me.ssdbgPagos.Columns("Opcion").Value & "'" & ","
            strSQL = strSQL & dPagoFac & ","
            strSQL = strSQL & "'" & Me.ssdbgPagos.Columns("Referencia").Value & "',"
            strSQL = strSQL & IIf(Me.ssdbgPagos.Columns("IdAfiliacion").Value = vbNullString, 0, Me.ssdbgPagos.Columns("IdAfiliacion").Value) & ","
            strSQL = strSQL & "'" & Me.ssdbgPagos.Columns("LoteNumero").Value & "',"
            strSQL = strSQL & "'" & Me.ssdbgPagos.Columns("OperacionNumero").Value & "',"
            strSQL = strSQL & Me.ssdbgPagos.Columns("ImporteRecibido").Value & ","
            strSQL = strSQL & "'" & Me.ssdbgPagos.Columns("FechaOperacion").Value & "')"
            
            adocmdFactura.CommandText = strSQL
            adocmdFactura.Execute
        
            dTotPagPorFactura = dTotPagPorFactura + dPagoFac
            lRowPagos = lRowPagos + 1
        
        Loop
        
        'Actualiza el total del recibo
        strSQL = "UPDATE RECIBOS SET Total=" & dTotalFactura
        strSQL = strSQL & " WHERE NumeroRecibo=" & lNumeroRecibo
        
        adocmdFactura.CommandText = strSQL
        adocmdFactura.Execute
        
        lNumeroRecibo = lNumeroRecibo + 1
   Next

'------------------------------------------------------------------------------
    'Hace commit
    Conn.CommitTrans
    
    iInitTrans = 0
   
   MDIPrincipal.StatusBar1.Panels(1).Text = "Actualizando Fechas"
   
   lNumFacIniImp = lNumeroFacturaInicial
   lNumFacFinImp = lNumeroFactura - 1
   
   lNumFolioFacIniImp = sSerieFactura & lNumeroFolioFacturaInicial
   lNumFolioFacFinImp = sSerieFactura & lNumeroFolioFactura - 1

   
   lNumRecIniImp = lNumeroReciboInicial
   lNumRecFinImp = lNumeroRecibo - 1
   
    Set adocmdFactura = Nothing
    
    MDIPrincipal.StatusBar1.Panels(1).Text = "Actualizando Fechas"
    DoEvents
    
    If lNumerodeFacturas > 0 Then
        ActualizaFechas lNumFacIniImp, lNumFacFinImp, 0
    End If
    
    If lNumerodeRecibos > 0 Then
        ActualizaFechas lNumRecIniImp, lNumRecFinImp, 1
    End If
    
    'Genera el CFD
    If lNumerodeFacturas > 0 Then
        MDIPrincipal.StatusBar1.Panels(1).Text = "Generando CFD en " & ObtieneParametro("URL_WS_CFD")
        DoEvents

        Select Case iNumeroCaja
            Case 1
                sSerieCFD = ObtieneParametro("SERIE_CFD_FACTURA_CAJA")
            Case 2
                sSerieCFD = ObtieneParametro("SERIE_CFD_FACTURA_DIRE")
            Case Else
            sSerieCFD = ObtieneParametro("SERIE_CFD_FACTURA_CAJA")
        End Select

        sFolioCFD = GeneraCFD(lNumFacIniImp, sSerieCFD, "ingreso")

        If Len(sFolioCFD) > 12 Then
        
            MsgBox "Ocurrio un error generando el CFD" & vbCrLf & sFolioCFD, vbCritical, "Error"
            
        ElseIf sFolioCFD = "" Then
        
        sFolioCFD = ErrorFolioCFD(lNumFacIniImp)
            MsgBox "Ocurrio un error generando el CFD" & vbCrLf & sFolioCFD & vbCrLf & "Favor de Cancelar y correguir.", vbCritical, "Error"
       
        Else
            If sFolioCFD <> vbNullString Then
                MDIPrincipal.StatusBar1.Panels(1).Text = "Actualizando FolioCFD"
                DoEvents
                If ActualizaFolioCFD(lNumFacIniImp, sFolioCFD, sSerieCFD, "F") = 0 Then
                End If
            End If
        End If
    End If

    MDIPrincipal.StatusBar1.Panels(1).Text = "Actualizando Direccion"
    
    If Me.chkDatosFac Then ActualizaDireccion (lIdTitular)
    
    If lNumerodeRecibos > 0 Then
        GeneraCupones "R", lNumRecIniImp, lNumRecFinImp, lNumeroCupones
    End If
    
    If lNumerodeFacturas > 0 Then
        GeneraCupones "F", lNumFacIniImp, lNumFacFinImp, lNumeroCupones
    End If
    
    Screen.MousePointer = vbDefault
    
'    If lNumerodeFacturas > 0 And iNumeroCaja = 1 Then
'
'        Dim frmImp As frmImpFac
'
'        Set frmImp = New frmImpFac
'        frmImp.cModo = "F"
'        frmImp.Tag = "F"
'        frmImp.lNumeroInicial = lNumFacIniImp
'        frmImp.lNumeroFinal = lNumFacFinImp
'        frmImp.Show 1
'
'        Set frmImp = Nothing
'
'    End If
   
    If lNumerodeFacturas > 0 And iNumeroCaja <> 2 Then
        Dim frmImpRec As frmImpFac
        
        Set frmImpRec = New frmImpFac
        frmImpRec.cModo = "R"
        frmImpRec.Tag = "R"
        frmImpRec.lNumeroInicial = lNumFacIniImp
        frmImpRec.lNumeroFinal = lNumFacFinImp
        frmImpRec.Caption = "Imprime Recibos"
        
        frmImpRec.Show 1
        
        Set frmImpRec = Nothing
        
    End If
    
    If lNumerodeRecibos > 0 And lNumeroCupones > 0 And iNumeroCaja <> 2 Then
        frmImpCupones.Tag = "R"
        frmImpCupones.Caption = "Imprime Cupones"
        frmImpCupones.Show 1
    End If
    
    If lNumerodeFacturas > 0 And lNumeroCupones > 0 And iNumeroCaja <> 2 Then
        frmImpCupones.Tag = "F"
        frmImpCupones.Caption = "Imprime Cupones"
        frmImpCupones.Show 1
    End If
    
    If sFolioCFD <> vbNullString And iNumeroCaja > 1 Then
        'MuestraCFD (lNumFacIniImp)
        MsgBox "Se genero CFD " & sSerieCFD & sFolioCFD, vbInformation, "Ok"
    End If
    
    MDIPrincipal.StatusBar1.Panels(1).Text = "Actualizando Acceso"
    
    If lNumerodeFacturas > 0 And lHayMantenimiento Then
        ActualizaAcceso lIdTitular
        ActivaAcceso lIdTitular
    End If
    
   LimpiaTodo
   
   Exit Sub
Error_Catch:

    Dim lI As Long
    Dim sCadError As String
    
    lI = 0
    sCadError = ""
    
    'Hace RollBack
    If iInitTrans > 0 Then
        Conn.RollbackTrans
        'Restablece pagos
        'RestablecePagos
    End If
    
    For lI = 0 To Conn.Errors.Count - 1
        sCadError = sCadError & Conn.Errors.Item(lI).Description & "(" & Conn.Errors.Item(lI).Number & ")" & vbLf
    Next
    
    If sCadError = vbNullString Then
        If Err.Number <> 0 Then
            sCadError = sCadError
        End If
    End If
    
    sCadError = sCadError & "Verifique el número de Folio Siguiente!"
    
    Screen.MousePointer = vbDefault
   
    MsgBox sCadError, vbCritical, "Error"
End Sub

Private Sub cmdIns_Click()
    
    Dim lNumRen As Long
    
    
    If Me.lblNombreUsu.Caption = "" Then
        MsgBox "No se ha selecionado ningún usuario!", vbExclamation, "Facturación"
        Exit Sub
    End If
    
    lNumRen = Me.ssdbgFactura.Rows
    
    frmFacInserta.Tag = lIdTitular
    frmFacInserta.Show 1
    
    
    If Me.ssdbgFactura.Rows > lNumRen Then
        Me.ssdbgFactura.Bookmark = Me.ssdbgFactura.GetBookmark(Me.ssdbgFactura.Rows - 1)
        If Me.ssdbgFactura.Rows > Me.ssdbgFactura.VisibleRows Then
            Me.ssdbgFactura.Scroll 0, 1
        End If
        
        RecalculaRenglon
        CalculaTotales
        
        siTotalPorPagar = siTotTotal - siTotalPagado
    
    
        Me.lblPagado = Format(siTotalPagado, "#,##0.00")
        Me.lblPorPagar = Format(siTotalPorPagar, "#,##0.00")
        
        Me.ssdbgFactura.Bookmark = Me.ssdbgFactura.Rows - 1
        
    End If
    
End Sub

Private Sub cmdInsertaPago_Click()

    Dim lIndex As Long
    Dim vBookMark As Variant
    Dim bFound As Boolean
    
    
    If Me.lblNombreUsu.Caption = "" Then
        MsgBox "No se ha selecionado ningún usuario!", vbExclamation, "Facturación"
        Exit Sub
    End If
    
    If Not CBool(Me.ssdbgFactura.Rows) Then
        MsgBox "No hay conceptos que pagar!", vbExclamation, "Facturación"
        Exit Sub
    End If
    
    If Me.ssCmbFormaPago.Text = vbNullString Then
        MsgBox "Selecione una forma de pago!", vbInformation, "Facturacion"
        Exit Sub
    End If
    
    'Si esta forma de pago tiene opciones y no se ha seleccionado
    If Me.ssCmbOpcionPago.Rows And Me.ssCmbOpcionPago.Text = vbNullString Then
        MsgBox "Selecione una opcion para esta forma de pago!", vbInformation, "Facturacion"
        Exit Sub
    End If
    
    If Val(Me.txtImportePago.Text) <= 0 Then
        MsgBox "El importe del pago debe ser mayor a 0!", vbInformation, "Facturacion"
        Me.txtImportePago.SetFocus
        Exit Sub
    End If
    
    
    If siTotTotal = 0 Then
        MsgBox "No hay nada por pagar!", vbExclamation, "Facturacion"
        Me.txtImportePago.Text = vbNullString
        Me.txtRefer.Text = vbNullString
        Me.ssCmbFormaPago.Text = vbNullString
        Exit Sub
    End If
    
    If siTotalPorPagar = 0 Then
        MsgBox "Ya se cubrió el importe totalmente!", vbExclamation, "Facturacion"
        Me.txtImportePago.Text = vbNullString
        Me.txtRefer.Text = vbNullString
        Me.ssCmbFormaPago.Text = vbNullString
        Exit Sub
    End If
    
    
   If Round(siTotalPagado + CSng(Me.txtImportePago.Text), 2) > Round(siTotTotal, 2) Then
        MsgBox "El importe del pago excedería el adeudo!", vbExclamation, "Facturacion"
        Me.txtImportePago.SelStart = 0
        Me.txtImportePago.SelLength = Len(Me.txtImportePago.Text)
        Me.txtImportePago.SetFocus
        Exit Sub
    End If
    
    'Si la forma de pago pide referencia checa que esta exista
    If Me.ssCmbFormaPago.Columns("RequiereReferencia").Value = 1 And Me.txtRefer.Text = vbNullString Then
        MsgBox "Esta forma de pago requiere referencia!", vbCritical, "Verifique"
        Me.txtRefer.SetFocus
        Exit Sub
    End If
    
    
    
    bFound = False
    vBookMark = Me.ssdbgPagos.GetBookmark(0)
    
    For lIndex = 0 To Me.ssdbgPagos.Rows - 1
        Me.ssdbgPagos.Bookmark = Me.ssdbgPagos.AddItemBookmark(lIndex)
        If ssdbgPagos.Columns("IdTipoPago").Value = Me.ssCmbFormaPago.Columns("IdFormaPago").Value And ssdbgPagos.Columns("Referencia").Value = Trim(Me.txtRefer.Text) Then
            bFound = True
            Exit For
        End If
    Next
    
    If bFound Then
        Me.ssdbgPagos.Columns("Importe").Value = CDbl(Me.ssdbgPagos.Columns("Importe").Value) + CDbl(Me.txtImportePago.Text)
        Me.ssdbgPagos.Columns("Resta").Value = Me.ssdbgPagos.Columns("Resta").Value + CDbl(Me.txtImportePago.Text)
        Me.ssdbgPagos.Update
    Else
        Me.ssdbgPagos.AddItem Me.ssCmbFormaPago.Columns("IdFormaPago").Value & vbTab & Me.ssCmbFormaPago.Columns("FormaPago").Value & vbTab & Me.ssCmbOpcionPago.Columns("Descripcion").Value & vbTab & Me.txtImportePago.Text & vbTab & Me.txtRefer & vbTab & Me.txtImportePago.Text
    End If
    
    siTotalPagado = siTotalPagado + Val(Me.txtImportePago.Text)
    siTotalPorPagar = siTotTotal - siTotalPagado
    
    
    Me.lblPagado = Format(siTotalPagado, "#,##0.00")
    Me.lblPorPagar = Format(siTotalPorPagar, "#,##0.00")
    
    
    Me.txtImportePago.Text = vbNullString
    Me.txtRefer.Text = vbNullString
    Me.ssCmbFormaPago.Text = vbNullString
    Me.ssCmbOpcionPago.Text = vbNullString
    
    If Me.ssdbgPagos.Rows Then
        Me.ssdbgPagos.Caption = "Pagos " & "(" & Me.ssdbgPagos.Rows & ")"
        Me.ssdbgPagos.Bookmark = Me.ssdbgPagos.Rows - 1
    Else
        Me.ssdbgPagos.Caption = "Pagos"
    End If
    
    
End Sub

Private Sub cmdInsertaPagoEx_Click()
    Dim frmFP As frmFacInsFormaPago
    
    Dim lNumRen As Long
    
    
    
    If Me.lblNombreUsu.Caption = "" Then
        MsgBox "No se ha selecionado ningún usuario!", vbExclamation, "Facturación"
        Exit Sub
    End If
    
    If Not CBool(Me.ssdbgFactura.Rows) Then
        MsgBox "No hay conceptos que pagar!", vbExclamation, "Facturación"
        Exit Sub
    End If
    
    
    
    
    lNumRen = Me.ssdbgPagos.Rows
    
    Set frmFP = New frmFacInsFormaPago
    
    frmFP.doImporte = siTotalPorPagar
    
    frmFP.Show vbModal
    
    
    'Si se inserto un renglon
    If Me.ssdbgPagos.Rows > lNumRen Then
    
        'Se ubica en el ultimo renglon
        Me.ssdbgPagos.Bookmark = Me.ssdbgPagos.AddItemBookmark(Me.ssdbgPagos.Rows - 1)
    
        siTotalPagado = siTotalPagado + Me.ssdbgPagos.Columns("Importe").Value
        siTotalPorPagar = siTotTotal - siTotalPagado
    
    
        Me.lblPagado = Format(siTotalPagado, "#,##0.00")
        Me.lblPorPagar = Format(siTotalPorPagar, "#,##0.00")
        
        If Me.ssdbgPagos.Rows Then
            Me.ssdbgPagos.Caption = "Pagos " & "(" & Me.ssdbgPagos.Rows & ")"
            Me.ssdbgPagos.Bookmark = Me.ssdbgPagos.Rows - 1
        Else
            Me.ssdbgPagos.Caption = "Pagos"
        End If
        
        
    End If
    
    
    
End Sub

Private Sub cmdPagoRapido_Click()
    Dim adorcsPago As ADODB.Recordset
    Dim adorcs As ADODB.Recordset
    
    Dim sFormaPago As String
    Dim sOpcionPago As String
    Dim sLoteNumero As String
    Dim sOperacionNumero As String
    
    Dim dFechaOPeracion As Date
    
    
    
    Set adorcs = New ADODB.Recordset
    adorcs.CursorLocation = adUseServer
        
    
    strSQL = "SELECT CFG_PAGO_RAPIDO.IdCaja, CFG_PAGO_RAPIDO.IdFormaPago, CFG_PAGO_RAPIDO.IdAfiliacion, CFG_PAGO_RAPIDO.LoteNumero, CFG_PAGO_RAPIDO.OperacionNumero, CFG_PAGO_RAPIDO.FechaOperacion, CFG_PAGO_RAPIDO.FechaAlta"
    strSQL = strSQL & " From CFG_PAGO_RAPIDO"
    strSQL = strSQL & " WHERE ("
    strSQL = strSQL & " ((CFG_PAGO_RAPIDO.IdCaja) =" & iNumeroCaja & ")"
    strSQL = strSQL & ")"

    Set adorcsPago = New ADODB.Recordset
    adorcsPago.CursorLocation = adUseServer
    
    adorcsPago.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    
    If adorcsPago!FechaAlta < Date Then
        MsgBox "La fecha de configuración del pago rápido es menor al día de hoy", vbExclamation + vbOKOnly, "Confirme"
    End If

    'Busca la descripción de la forma de pago
    strSQL = "SELECT DESCRIPCION"
    strSQL = strSQL & " FROM FORMA_PAGO"
    strSQL = strSQL & " WHERE ("
    strSQL = strSQL & " ((IdFormaPago)=" & adorcsPago!IdFormaPago & ")"
    strSQL = strSQL & ")"
    
    adorcs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not adorcs.EOF Then
        sFormaPago = Trim(adorcs!Descripcion)
    Else
        MsgBox "La forma de pago configurada no existe!", vbCritical, "Verifique"
        
        adorcs.Close
        Set adorcs = Nothing
        
        adorcsPago.Close
        Set adorcsPago = Nothing
        
        Exit Sub
        
        
    End If
    
    adorcs.Close
    
    'Busca la descripción de la afiliacion
    strSQL = "SELECT Modalidad"
    strSQL = strSQL & " FROM CT_AFILIACIONES"
    strSQL = strSQL & " WHERE ("
    strSQL = strSQL & " ((IdAfiliacion)=" & adorcsPago!IdAfiliacion & ")"
    strSQL = strSQL & ")"
    
    adorcs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not adorcs.EOF Then
        sOpcionPago = adorcs!Modalidad
    Else
'        MsgBox "La modalidad de pago configurada no existe!", vbCritical, "Verifique"
'
'        adorcs.Close
'        Set adorcs = Nothing
'
'        adorcsPago.Close
'        Set adorcsPago = Nothing
'
'        Exit Sub
    End If
    
    adorcs.Close
    Set adorcs = Nothing
    
    
    
    Me.ssdbgPagos.AddItem adorcsPago!IdFormaPago & vbTab & sFormaPago & vbTab & sOpcionPago & vbTab & siTotalPorPagar & vbTab & vbNullString & vbTab & siTotalPorPagar & vbTab & adorcsPago!IdAfiliacion & vbTab & adorcsPago!LoteNumero & vbTab & adorcsPago!OperacionNumero & vbTab & siTotalPorPagar & vbTab & adorcsPago!FechaOperacion
    
    
    
    adorcsPago.Close
    Set adorcsPago = Nothing
    
    
    Me.ssdbgPagos.Bookmark = Me.ssdbgPagos.AddItemBookmark(Me.ssdbgPagos.Rows - 1)
    
    siTotalPagado = siTotalPagado + Me.ssdbgPagos.Columns("Importe").Value
    siTotalPorPagar = siTotTotal - siTotalPagado
    
    
    Me.lblPagado = Format(siTotalPagado, "#,##0.00")
    Me.lblPorPagar = Format(siTotalPorPagar, "#,##0.00")
        
    If Me.ssdbgPagos.Rows Then
        Me.ssdbgPagos.Caption = "Pagos " & "(" & Me.ssdbgPagos.Rows & ")"
        Me.ssdbgPagos.Bookmark = Me.ssdbgPagos.Rows - 1
    Else
        Me.ssdbgPagos.Caption = "Pagos"
    End If
    
    
    
End Sub

Private Sub Form_Load()
    
    Me.Height = 9200
    Me.Width = 15255
        
    CentraForma MDIPrincipal, Me
    
    
    Me.dtpFechaCalc.Value = Date
  
    
    Me.lblSubTotal = Format(0, "#,##0.00")
    Me.lblDescuento = Format(0, "#,##0.00")
    Me.lblTotal = Format(0, "#,##0.00")
    
    boYaFacturado = False
    
    ''LlenaComboPeriodoCalculo
    LlenaComboPeriodo
    LlenaComboFormaPago
    LlenaComboEstados
    LlenaComboFormaDireccionar
    
    Me.cmbPeriodoCalculo.ListIndex = 0
End Sub

Private Sub cmdModRen_Click()
        
    If Not CBool(Me.ssdbgFactura.Rows) Then
        MsgBox "No hay renglones para modificar!", vbExclamation, "Facturación"
        Exit Sub
    End If
    
    If Not ChecaSeguridad(Me.Name, Me.cmdModRen.Name) Then
        Exit Sub
    End If
    
    
    frmFacModificaRen.Show 1
    
    RecalculaRenglon
    CalculaTotales
    
    siTotalPorPagar = siTotTotal - siTotalPagado
    
    
    Me.lblPagado = Format(siTotalPagado, "#,##0.00")
    Me.lblPorPagar = Format(siTotalPorPagar, "#,##0.00")
    
End Sub

Private Sub cmdObser_Click()
    Dim vCurBookMark As Variant
    Dim lI As Long
    Dim lIdMem As Long
    
    Dim sObser As String
    
    If Me.lblNombreUsu.Caption = "" Then
        MsgBox "No se ha selecionado ningún usuario!", vbExclamation, "Facturación"
        Exit Sub
    End If
    
    If Not CBool(Me.ssdbgPagos.Rows) Then
        MsgBox " No se ha insertado ninguna forma de pago! ", vbExclamation, "Facturación"
        Exit Sub
    End If
    
    
    If CBool(Me.chkDireccionar.Value) And Me.cmbTipoDireccionar.Text = "" Then
        MsgBox "No se ha elegido la forma de direccionamiento!", vbExclamation, "Facturación"
        Exit Sub
    End If
    
        
    vCurBookMark = ssdbgFactura.AddItemRowIndex(Me.ssdbgFactura.Bookmark)
    
    For lI = 0 To Me.ssdbgFactura.Rows - 1
    
        Me.ssdbgFactura.Bookmark = Me.ssdbgFactura.AddItemBookmark(lI)
        
        If Val(Me.ssdbgFactura.Columns("TipoCargo").Value) = 0 Then
            If Not (lIdMem <> 0 And lIdMem <> Me.ssdbgFactura.Columns("IdMember").Value) Then
                If lIdMem = 0 Then
                    lIdMem = Me.ssdbgFactura.Columns("IdMember").Value
                End If
                sObser = sObser & NombreMes(Month(CDate(Me.ssdbgFactura.Columns("Periodo").Value))) & " " & Right(Str(Year(CDate(Me.ssdbgFactura.Columns("Periodo").Value))), 2) & " "
            End If
        End If
    Next
    
'    If Me.chkDireccionar.Value And Len(sObser) Then
'        sObser = sObser & "DIRECCIONADO " & Trim(Me.cmbTipoDireccionar.Text) & " "
'    End If
    
    Me.ssdbgFactura.Bookmark = vCurBookMark
    
    If sObser <> "" Then
        sObser = "PAGO MANT. " & Trim(sObser)
    End If
    
    If Me.ssdbgPagos.Rows Then
        For lI = 0 To Me.ssdbgPagos.Rows - 1
            Me.ssdbgPagos.Bookmark = Me.ssdbgPagos.AddItemBookmark(lI)
            sObser = sObser & " " & Trim(Me.ssdbgPagos.Columns("DescripPago").Value) & " " & Trim(Me.ssdbgPagos.Columns("Referencia").Value) & " " & Me.ssdbgPagos.Columns("Opcion").Value
        Next
    End If
    
    Me.txtObserva.Text = Trim(sObser)
    
End Sub

Private Sub cmdQuitaInteres_Click()
    Dim byResp As Byte
    Dim vCurBookMark As Variant
    Dim lI As Long
    Dim oBookMark As Variant
    
    
    If Me.ssdbgFactura.Rows < 1 Then Exit Sub
    
    
    If Not ChecaSeguridad(Me.Name, Me.cmdQuitaInteres.Name) Then
        Exit Sub
    End If
    
    byResp = MsgBox("¿Está seguro de que quiere" & vbLf & "remover los intereses?", vbQuestion Or vbYesNo, "Facturación")
    
    If byResp <> vbYes Then Exit Sub
    
    vCurBookMark = Me.ssdbgFactura.Bookmark
    
    For lI = 0 To Me.ssdbgFactura.Rows - 1
        
        oBookMark = Me.ssdbgFactura.AddItemBookmark(lI)
        
        Me.ssdbgFactura.Bookmark = oBookMark
        
        Me.ssdbgFactura.Columns("Intereses").Value = 0
        Me.ssdbgFactura.Update
        
        RecalculaRenglon
    Next
    
    CalculaTotales
    
    Me.ssdbgFactura.Bookmark = vCurBookMark
    
    siTotalPorPagar = siTotTotal - siTotalPagado
    
    
    Me.lblPagado = Format(siTotalPagado, "#,##0.00")
    Me.lblPorPagar = Format(siTotalPorPagar, "#,##0.00")
    
    
End Sub

Private Sub SortArray(ByRef sArray As Variant, ByVal lNumEle As Long, bDescendente As Boolean)
    Dim lIndice As Long, lIndice2 As Long, lPrimerElemento As Long
    Dim lDistancia As Long, lValor As Long
    Dim sElemento As String
    
    lPrimerElemento = LBound(sArray)
    
    Do
        lDistancia = lDistancia * 3 + 1
    Loop Until lDistancia > lNumEle
    Do
        lDistancia = lDistancia \ 3
        For lIndice = lDistancia + lPrimerElemento To lNumEle + lPrimerElemento - 1
            lValor = Val(Left(sArray(lIndice), 10))
            sElemento = sArray(lIndice)
            lIndice2 = lIndice
            Do While (Val(Left(sArray(lIndice2 - lDistancia), 10)) > lValor Xor bDescendente)
                sArray(lIndice2) = sArray(lIndice2 - lDistancia)
                lIndice2 = lIndice2 - lDistancia
                If lIndice2 - lDistancia < lPrimerElemento Then Exit Do
            Loop
            sArray(lIndice2) = sElemento
        Next
    Loop Until lDistancia = 1
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> 1 Then
        Me.Height = 9200
        Me.Width = 15255
    End If
End Sub

Private Sub ssCmbFormaPago_Click()
    Me.txtImportePago.Text = Format(siTotalPorPagar, "#0.00")
    
    LlenaComboOpcionPago Me.ssCmbFormaPago.Columns("IdFormapago").Value
    
    Me.cmdInsertaPago.Default = True
End Sub

Private Sub ssdbgFactura_AfterColUpdate(ByVal ColIndex As Integer)
    If ColIndex <> 7 Then
        RecalculaRenglon
        CalculaTotales
        
        siTotalPorPagar = siTotTotal - siTotalPagado
    
    
        Me.lblPagado = Format(siTotalPagado, "#,##0.00")
        Me.lblPorPagar = Format(siTotalPorPagar, "#,##0.00")
        
    End If
End Sub

Private Sub ssdbgFactura_Click()
    Debug.Print Me.ssdbgFactura.Rows
    Debug.Print Me.ssdbgFactura.row
End Sub

Private Sub CalculaTotales()
    
    Dim lCurRow As Long
    Dim lI As Long
    Dim oBookMark As Variant
    
    siTotImporte = 0
    siTotDescuento = 0
    siTotIva = 0
    siTotTotal = 0
    
    frmFacturacion.ssdbgFactura.Update
    
    
    lCurRow = frmFacturacion.ssdbgFactura.AddItemRowIndex(Me.ssdbgFactura.Bookmark)
    
    
    For lI = 0 To frmFacturacion.ssdbgFactura.Rows - 1
    
        oBookMark = frmFacturacion.ssdbgFactura.AddItemBookmark(lI)
    
        dMonto = CDbl(frmFacturacion.ssdbgFactura.Columns("Cantidad").CellValue(oBookMark)) * CDbl(frmFacturacion.ssdbgFactura.Columns("Importe").CellValue(oBookMark))
        dImporte = dMonto + CDbl(frmFacturacion.ssdbgFactura.Columns("Intereses").CellValue(oBookMark))
        dIntereses = CDbl(frmFacturacion.ssdbgFactura.Columns("Intereses").CellValue(oBookMark))
        dDescuento = dImporte * CDbl(frmFacturacion.ssdbgFactura.Columns("Descuento").CellValue(oBookMark)) / 100
        
        dImporte = dMonto + dIntereses
                
        dTotal = dImporte - dDescuento
        
        siTotImporte = siTotImporte + dImporte
        siTotDescuento = siTotDescuento + dDescuento
        siTotIva = siTotIva + dIva
        siTotTotal = siTotTotal + dTotal
        
    Next
    frmFacturacion.ssdbgFactura.Bookmark = lCurRow
    
    Me.lblSubTotal = Format(siTotImporte, "#,##0.00")
    Me.lblDescuento = Format(siTotDescuento, "#,##0.00")
    Me.lblTotal = Format(siTotTotal, "#,##0.00")
    
    Me.lblRen = Format(Me.ssdbgFactura.Rows, "#0") & " Renglon(es)"

    
End Sub
''Carga el periodo del pago
Private Sub LlenaComboPeriodo()
    Dim adorcsPeriodoPago As ADODB.Recordset
    
    strSQL = "SELECT PeriodoPago, Descripcion"
    strSQL = strSQL & " FROM PERIODO_PAGO"
    strSQL = strSQL & " ORDER BY PeriodoPago"
    
    
    Set adorcsPeriodoPago = New ADODB.Recordset
    adorcsPeriodoPago.CursorLocation = adUseServer
    
    adorcsPeriodoPago.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    Do While Not adorcsPeriodoPago.EOF
        Me.cmbPeriodoCalculo.AddItem adorcsPeriodoPago!Descripcion
        Me.cmbPeriodoCalculo.ItemData(Me.cmbPeriodoCalculo.NewIndex) = adorcsPeriodoPago!PeriodoPago
        adorcsPeriodoPago.MoveNext
    Loop
    
    adorcsPeriodoPago.Close
    Set adorcsPeriodoPago = Nothing
End Sub
Private Sub LlenaComboFormaPago()
    Dim adorcsFormaPago As ADODB.Recordset
    
    strSQL = "SELECT IdFormaPago, Descripcion, TieneOpcion"
    strSQL = strSQL & " FROM FORMA_PAGO"
    strSQL = strSQL & " ORDER BY IdFormaPago"
    
    
    Set adorcsFormaPago = New ADODB.Recordset
    adorcsFormaPago.CursorLocation = adUseServer
    
    adorcsFormaPago.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    Do While Not adorcsFormaPago.EOF
        Me.ssCmbFormaPago.AddItem adorcsFormaPago!Descripcion & vbTab & adorcsFormaPago!IdFormaPago & vbTab & IIf(adorcsFormaPago!TieneOpcion = 0, 0, 1)
        adorcsFormaPago.MoveNext
    Loop
    
    adorcsFormaPago.Close
    Set adorcsFormaPago = Nothing
    
End Sub

Private Sub LlenaComboOpcionPago(lIdFormaPago As Long)
    Dim adorcsOpcionPago As ADODB.Recordset
    
    strSQL = "SELECT IdFormadePagoOpcion, Descripcion"
    strSQL = strSQL & " FROM FORMA_PAGO_OPCION"
    strSQL = strSQL & " WHERE idFormaPago=" & lIdFormaPago
    strSQL = strSQL & " ORDER BY IdFormadePagoOpcion"
    
    
    Set adorcsOpcionPago = New ADODB.Recordset
    adorcsOpcionPago.CursorLocation = adUseServer
    
    adorcsOpcionPago.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    Me.ssCmbOpcionPago.RemoveAll
    
    Do While Not adorcsOpcionPago.EOF
        Me.ssCmbOpcionPago.AddItem adorcsOpcionPago!Descripcion & vbTab & adorcsOpcionPago!idformadepagoopcion
        adorcsOpcionPago.MoveNext
    Loop
    
    adorcsOpcionPago.Close
    Set adorcsOpcionPago = Nothing
    
End Sub


Private Sub txtClave_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            KeyAscii = 0
            SendKeys vbTab
        Case 8 ' Tecla backspace
            KeyAscii = KeyAscii
        Case 48 To 57 ' Numeros del 0 al 9
            KeyAscii = KeyAscii
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtFacCiudad_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtFacColonia_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtFacCP_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8 ' Tecla backspace
            KeyAscii = KeyAscii
        Case 48 To 57 ' Numeros del 0 al 9
            KeyAscii = KeyAscii
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtFacDelOMuni_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtFacDireccion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtFacNombre_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtFacRFC_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtFacTelefono_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8 ' Tecla backspace
            KeyAscii = KeyAscii
        Case 40 To 42 ' teclas () *
            KeyAscii = KeyAscii
        Case 45 ' Tecla -
            KeyAscii = KeyAscii
        Case 48 To 57 ' Numeros del 0 al 9
            KeyAscii = KeyAscii
        Case Else
            KeyAscii = 0
    End Select

End Sub

Private Sub txtImportePago_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8 ' Tecla backspace
            KeyAscii = KeyAscii
        Case 46 'punto decimal
            If InStr(Me.txtImportePago.Text, ".") Then
                KeyAscii = 0
            Else
                KeyAscii = KeyAscii
            End If
        Case 48 To 57 ' Numeros del 0 al 9
            KeyAscii = KeyAscii
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtRefer_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txtObserva_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            KeyAscii = 0
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Private Sub LimpiaTodo()
    Me.txtClave.Text = vbNullString
    Me.txtFacNombre.Text = vbNullString
    Me.txtFacRFC.Text = vbNullString
    Me.txtFacDireccion.Text = vbNullString
    Me.txtFacColonia.Text = vbNullString
    Me.txtFacCiudad.Text = vbNullString
    Me.txtFacDelOMuni.Text = vbNullString
    Me.txtFacCP.Text = vbNullString
    Me.ssCmbFacEstado.Text = vbNullString
    Me.txtFacTelefono.Text = vbNullString
    
    Me.txtObserva.Text = vbNullString
    
    Me.chkDatosFac.Value = False
    
    Me.ssdbgFactura.RemoveAll
    Me.ssdbgPagos.RemoveAll
   
    Me.lblNombreUsu = vbNullString
    Me.lblInscripcion = vbNullString
    
    'Me.dtpFechaCalc.Value = Date
   
    Me.txtClave.SetFocus
   
    siTotImporte = 0
    siTotDescuento = 0
    siTotIva = 0
    siTotTotal = 0
   
    Me.lblSubTotal = Format(siTotImporte, "#,##0.00")
    Me.lblDescuento = Format(siTotDescuento, "#,##0.00")
    Me.lblTotal = Format(siTotTotal, "#,##0.00")
    
    siTotalPorPagar = 0
    siTotalPagado = 0
    
    Me.lblPagado = Format(siTotalPagado, "#,##0.00")
    Me.lblPorPagar = Format(siTotalPorPagar, "#,##0.00")

    Me.lblRen = Format(Me.ssdbgFactura.Rows, "#0") & " Renglon(es)"
    
    Me.ssdbgPagos.Caption = "Pagos"
    
    boYaFacturado = False
    
    lPublicoGeneral = False
    
    Me.chkDireccionar.Value = 0
    
    Me.cmdCalcular.Default = True
    
    Me.cmdFacturar.Enabled = True
End Sub

Private Sub CalculaRentables(lidMember)

        Dim AdoRcsRentables As ADODB.Recordset
        Dim adoRcsFac As ADODB.Recordset
        
        Dim lFormaPago As Long
        
        
        Dim i As Long
        
        '14 de febrero 2006
        
        
        Dim boProporcional As Boolean
        Dim dUltimaFechaPago As Date
        Dim dUltimoDiaPeriodo As Date
        
        Dim lPerProporcion As Long
        Dim siProporcion As Single
        
        
        Dim sDescripcion
        
        #If SqlServer_ Then
            strSQL = "SELECT CONCEPTO_RENTABLE.IdConcepto, CONCEPTO_RENTABLE.Periodo, CONCEPTO_INGRESOS.Descripcion, CONCEPTO_INGRESOS.Impuesto1, CONCEPTO_INGRESOS.Impuesto2,CONCEPTO_INGRESOS.FacORec, CONCEPTO_INGRESOS.Unidad, RENTABLES.FechaPago, USUARIOS_CLUB.Nombre, USUARIOS_CLUB.A_Paterno, USUARIOS_CLUB.A_Materno, USUARIOS_CLUB.IdMember, USUARIOS_CLUB.NumeroFamiliar, USUARIOS_CLUB.IdTipoUsuario, RENTABLES.IdTipoRentable, RENTABLES.Numero, RENTABLES.PeriodoPago "
            strSQL = strSQL & " FROM ((USUARIOS_CLUB INNER JOIN RENTABLES ON USUARIOS_CLUB.IdMember = RENTABLES.IdUsuario) INNER JOIN CONCEPTO_RENTABLE ON RENTABLES.IdTipoRentable=CONCEPTO_RENTABLE.IdTipoRentable) INNER JOIN CONCEPTO_INGRESOS ON CONCEPTO_RENTABLE.IdConcepto = CONCEPTO_INGRESOS.IdConcepto"
            strSQL = strSQL & " WHERE USUARIOS_CLUB.IdTitular=" & lidMember
            strSQL = strSQL & "AND RENTABLES.FechaPago < " & "'" & Format(Me.dtpFechaCalc.Value, "yyyymmdd") & "'"
        #Else
            strSQL = "SELECT CONCEPTO_RENTABLE.IdConcepto, CONCEPTO_RENTABLE.Periodo, CONCEPTO_INGRESOS.Descripcion, CONCEPTO_INGRESOS.Impuesto1, CONCEPTO_INGRESOS.Impuesto2,CONCEPTO_INGRESOS.FacORec, CONCEPTO_INGRESOS.Unidad, RENTABLES.FechaPago, USUARIOS_CLUB.Nombre, USUARIOS_CLUB.A_Paterno, USUARIOS_CLUB.A_Materno, USUARIOS_CLUB.IdMember, USUARIOS_CLUB.NumeroFamiliar, USUARIOS_CLUB.IdTipoUsuario, RENTABLES.IdTipoRentable, RENTABLES.Numero, RENTABLES.PeriodoPago "
            strSQL = strSQL & " FROM ((USUARIOS_CLUB INNER JOIN RENTABLES ON USUARIOS_CLUB.IdMember = RENTABLES.IdUsuario) INNER JOIN CONCEPTO_RENTABLE ON RENTABLES.IdTipoRentable=CONCEPTO_RENTABLE.IdTipoRentable) INNER JOIN CONCEPTO_INGRESOS ON CONCEPTO_RENTABLE.IdConcepto = CONCEPTO_INGRESOS.IdConcepto"
            strSQL = strSQL & " WHERE USUARIOS_CLUB.IdTitular=" & lidMember
            strSQL = strSQL & " AND RENTABLES.FechaPago < " & "#" & Format(Me.dtpFechaCalc.Value, "mm/dd/yyyy") & "#"
        #End If
        Set AdoRcsRentables = New ADODB.Recordset
        
        AdoRcsRentables.ActiveConnection = Conn
        AdoRcsRentables.CursorLocation = adUseServer
        AdoRcsRentables.CursorType = adOpenForwardOnly
        AdoRcsRentables.LockType = adLockReadOnly
        AdoRcsRentables.Open strSQL
        
        Do While Not AdoRcsRentables.EOF
        
            
        
            dUltimaFechaPago = AdoRcsRentables!Fechapago
            dUltimoDiaPeriodo = UltimoDiaDelPeriodo(dUltimaFechaPago, AdoRcsRentables!PeriodoPago, False)
            
            If AdoRcsRentables!PeriodoPago = 12 Then
                lPerProporcion = DateDiff("m", dUltimaFechaPago, dUltimoDiaPeriodo) + 1
            Else
                lPerProporcion = DateDiff("d", dUltimaFechaPago, dUltimoDiaPeriodo)
            End If
            
            If dUltimaFechaPago = dUltimoDiaPeriodo Then
                If AdoRcsRentables!PeriodoPago = 12 Then
                    lPerProporcion = 12
                Else
                    lPerProporcion = Day(dUltimaFechaPago)
                End If
            End If
            
            If AdoRcsRentables!PeriodoPago = 12 Then
                siProporcion = lPerProporcion / 12
            Else
                siProporcion = lPerProporcion * (1 / Day(dUltimoDiaPeriodo))
            End If
            
'            If siProporcion < 1 Then
'                dPeriodo = DateAdd("yyyy", -1, dUltimoDiaPeriodo)
'                dUltimaFechaPago = dPeriodo
'            Else
'                dPeriodo = AdoRcsRentables!Fechapago
'            End If
                
             If dUltimaFechaPago = dUltimoDiaPeriodo Then
                dPeriodo = AdoRcsRentables!Fechapago
            Else
                If AdoRcsRentables!PeriodoPago = 12 Then
                    dPeriodo = DateAdd("yyyy", -1, dUltimoDiaPeriodo)
                    dUltimaFechaPago = dPeriodo
                Else
                    dPeriodo = DateAdd("m", -1, dUltimoDiaPeriodo)
                    dUltimaFechaPago = dPeriodo
                End If
            End If

            
            lPeriodos = CalculaPeriodos(dUltimaFechaPago, Me.dtpFechaCalc.Value, AdoRcsRentables!PeriodoPago)
            
            
            
            
            For i = 1 To lPeriodos
            
                If i > 1 Then
                    siProporcion = 1
                End If
            
            
                dPeriodo = DateAdd("m", i * AdoRcsRentables!PeriodoPago, dUltimaFechaPago)
                
                If AdoRcsRentables!PeriodoPago = 1 Then
                    If dPeriodo <> UltimoDiaDelPeriodo(dPeriodo, 1, False) Then
                        dPeriodo = UltimoDiaDelPeriodo(dPeriodo, 1, False)
                    End If
                End If
                
                
                #If SqlServer_ Then
                    strSQL = "SET DATEFORMAT dmy" & vbCrLf
                    strSQL = strSQL & " SELECT MONTO"
                    strSQL = strSQL & " FROM HISTORICO_RENTABLES"
                    strSQL = strSQL & " WHERE IdTipoRentable=" & AdoRcsRentables!idtiporentable
                    strSQL = strSQL & " AND Periodo=" & AdoRcsRentables!PeriodoPago
                    strSQL = strSQL & " AND VigenteDesde <=" & "'" & dPeriodo & "'"
                    strSQL = strSQL & " AND VigenteHasta >=" & "'" & dPeriodo & "'"
                #Else
                    strSQL = "SELECT MONTO"
                    strSQL = strSQL & " FROM HISTORICO_RENTABLES"
                    strSQL = strSQL & " WHERE IdTipoRentable=" & AdoRcsRentables!idtiporentable
                    strSQL = strSQL & " AND Periodo=" & AdoRcsRentables!PeriodoPago
                    strSQL = strSQL & " AND VigenteDesde <=" & "#" & dPeriodo & "#"
                    strSQL = strSQL & " AND VigenteHasta >=" & "#" & dPeriodo & "#"
                #End If
                Set adoRcsFac = New ADODB.Recordset
                adoRcsFac.ActiveConnection = Conn
                adoRcsFac.CursorLocation = adUseServer
                adoRcsFac.CursorType = adOpenForwardOnly
                adoRcsFac.LockType = adLockReadOnly
                adoRcsFac.Open strSQL
                
                If adoRcsFac.EOF Then
                    dCuota = 0
                    dCantidad = 1
                    dDescuentoPor = 0
                Else
                    dCuota = adoRcsFac!Monto
                    dCantidad = 1
                    dDescuentoPor = 0
                End If
                
                dIvaPor = AdoRcsRentables!Impuesto1 / 100
                lFormaPago = AdoRcsRentables!Periodo
                
                adoRcsFac.Close
                
                
                dMonto = Round(dCuota * dCantidad * siProporcion, 0)
                dIntereses = CalculaInteresesPasados(dPeriodo, Date, dMonto, 3)
                dIntereses = dIntereses + CalculaInteresesActuales(dPeriodo, Date, dMonto, 3)
                dDescuento = 0
                
                dIva = dMonto - Round(dMonto / (1 + dIvaPor), 2)
                dIvaDescuento = dDescuento - Round(dDescuento / (1 + dIvaPor), 2)
                dIvaIntereses = dIntereses - Round(dIntereses / (1 + dIvaPor), 2)
                
                dImporte = dMonto + dIntereses
                
                
                dTotal = dImporte - dDescuento
                
                
                
                siTotImporte = siTotImporte + dImporte
                siTotDescuento = siTotDescuento + dDescuento
                siTotIva = siTotIva + dIva
                siTotTotal = siTotTotal + dTotal
                
                sDescripcion = AdoRcsRentables!Descripcion & " " & Trim(AdoRcsRentables!Numero)
                
                If siProporcion < 1 Then
                    If AdoRcsRentables!PeriodoPago = 12 Then
                        sDescripcion = sDescripcion & " (" & lPerProporcion & "M)"
                    Else
                        sDescripcion = sDescripcion & " (" & lPerProporcion & "D)"
                    End If
                End If
                
                
                'Columnas del grid
                '0  Concepto
                '1  Nombre
                '2  Periodo
                '3  Cantidad
                '4  Importe
                '5  Intereses
                '6  Descuento
                '7  Total
                '8  Clave
                '9  IvaPor          No Visible
                '10  Iva            No Visible
                '11 IvaDescuento    No Visible
                '12 IvaIntereses    No Visible
                '13 DescMonto       No Visible
                '14 IdMember        No Visible
                '15 NoFamiliar      No Visible
                '16 Periodo         No Visible
                '17 IdTipoUsuario   No Visible
                '18 TipoCargo       No Visible
                '19 Auxiliar        No Visible
                
                
                sCargos(lIndexsCargos) = Format(Date2Days(dPeriodo), "0000000000") & vbTab & sDescripcion & vbTab & Trim(AdoRcsRentables!Numero) & vbTab & dPeriodo & vbTab & dCantidad & vbTab & dMonto & vbTab & dIntereses & vbTab & dDescuento & vbTab & dImporte & vbTab & AdoRcsRentables!IdConcepto & vbTab & dIvaPor & vbTab & dIva & vbTab & dIvaDescuento & vbTab & dIvaIntereses & vbTab & 0 & vbTab & AdoRcsRentables!Idmember & vbTab & AdoRcsRentables!NumeroFamiliar & vbTab & lFormaPago & vbTab & AdoRcsRentables!idtipousuario & vbTab & 1 & vbTab & Trim(AdoRcsRentables!Numero) & vbTab & AdoRcsRentables!FacORec & vbTab & "0" & vbTab & AdoRcsRentables!Unidad
                
                
                
                lIndexsCargos = lIndexsCargos + 1
            Next
            AdoRcsRentables.MoveNext
        Loop
        
        AdoRcsRentables.Close

        
        
        
        
End Sub

Private Sub CalculaMantenimiento(ByRef lidMember As Long, ByVal iPeriodoCalculo As Integer, ByRef aFacturacion() As String, ByRef lIndexFacturacion As Long)
    
    Dim AdoRcsUsuarios As ADODB.Recordset
    Dim adoRcsFac As ADODB.Recordset
    Dim adorcsAus As ADODB.Recordset
    
    Dim i As Long   'Variable para ciclos for
    
    
    
    Dim iUltimoDiadelMes As Integer
    Dim siProporcion As Single
    Dim iDiasProporcion As Integer
        
    '26/06/06 gpo
    Dim sngPorInter As Single 'Porcentaje de intereses
    Dim sInter As String
    
    
    sInter = ObtieneParametro("INTERES MENSUAL")
    
    If sInter = vbNullString Then
        sngPorInter = 3
    Else
        sngPorInter = CSng(sInter)
    End If
    
    #If SqlServer_ Then
        strSQL = "SELECT CONCEPTO_TIPO.IdConcepto, CONCEPTO_TIPO.Periodo, CONCEPTO_INGRESOS.Descripcion, CONCEPTO_INGRESOS.Impuesto1, CONCEPTO_INGRESOS.Impuesto2, CONCEPTO_INGRESOS.FacORec, CONCEPTO_INGRESOS.Unidad, FECHAS_USUARIO.FechaUltimoPago, USUARIOS_CLUB.Nombre, USUARIOS_CLUB.A_Paterno, USUARIOS_CLUB.A_Materno, USUARIOS_CLUB.IdMember, USUARIOS_CLUB.NumeroFamiliar, USUARIOS_CLUB.IdTipoUsuario, TIPO_USUARIO.Descripcion AS [TIPO_USUARIO.Descripcion] "
        strSQL = strSQL & " FROM (((USUARIOS_CLUB INNER JOIN CONCEPTO_TIPO ON USUARIOS_CLUB.IdTipoUsuario = CONCEPTO_TIPO.IdTipoUsuario) INNER JOIN CONCEPTO_INGRESOS ON CONCEPTO_TIPO.IdConcepto = CONCEPTO_INGRESOS.IdConcepto) INNER JOIN FECHAS_USUARIO ON (USUARIOS_CLUB.IdMember = FECHAS_USUARIO.IdMember) AND (CONCEPTO_TIPO.IdConcepto = FECHAS_USUARIO.IdConcepto)) LEFT JOIN TIPO_USUARIO ON USUARIOS_CLUB.IdTipoUsuario=TIPO_USUARIO.IdTipoUsuario"
        strSQL = strSQL & " WHERE USUARIOS_CLUB.IdTitular=" & lidMember
        strSQL = strSQL & " ORDER BY USUARIOS_CLUB.NumeroFamiliar"
    #Else
        strSQL = "SELECT CONCEPTO_TIPO.IdConcepto, CONCEPTO_TIPO.Periodo, CONCEPTO_INGRESOS.Descripcion, CONCEPTO_INGRESOS.Impuesto1, CONCEPTO_INGRESOS.Impuesto2, CONCEPTO_INGRESOS.FacORec, CONCEPTO_INGRESOS.Unidad, FECHAS_USUARIO.FechaUltimoPago, USUARIOS_CLUB.Nombre, USUARIOS_CLUB.A_Paterno, USUARIOS_CLUB.A_Materno, USUARIOS_CLUB.IdMember, USUARIOS_CLUB.NumeroFamiliar, USUARIOS_CLUB.IdTipoUsuario, TIPO_USUARIO.Descripcion "
        strSQL = strSQL & " FROM (((USUARIOS_CLUB INNER JOIN CONCEPTO_TIPO ON USUARIOS_CLUB.IdTipoUsuario = CONCEPTO_TIPO.IdTipoUsuario) INNER JOIN CONCEPTO_INGRESOS ON CONCEPTO_TIPO.IdConcepto = CONCEPTO_INGRESOS.IdConcepto) INNER JOIN FECHAS_USUARIO ON (USUARIOS_CLUB.IdMember = FECHAS_USUARIO.IdMember) AND (CONCEPTO_TIPO.IdConcepto = FECHAS_USUARIO.IdConcepto)) LEFT JOIN TIPO_USUARIO ON USUARIOS_CLUB.IdTipoUsuario=TIPO_USUARIO.IdTipoUsuario"
        strSQL = strSQL & " WHERE USUARIOS_CLUB.IdTitular=" & lidMember
        strSQL = strSQL & " ORDER BY USUARIOS_CLUB.NumeroFamiliar"
    #End If
    
    
    
    Set AdoRcsUsuarios = New ADODB.Recordset
        
    AdoRcsUsuarios.ActiveConnection = Conn
    AdoRcsUsuarios.CursorLocation = adUseServer
    AdoRcsUsuarios.CursorType = adOpenForwardOnly
    AdoRcsUsuarios.LockType = adLockReadOnly
    AdoRcsUsuarios.Open strSQL


    siTotImporte = 0
    siTotIntereses = 0
    siTotDescuento = 0
    siTotTotal = 0
    siTotIva = 0
    siTotIvaIntereses = 0
        
    lIndexsCargos = 0
        
    'Crea e inicializa adoRcsFac
    Set adoRcsFac = New ADODB.Recordset
    adoRcsFac.ActiveConnection = Conn
    adoRcsFac.CursorLocation = adUseServer
    adoRcsFac.CursorType = adOpenForwardOnly
    adoRcsFac.LockType = adLockReadOnly
    
    'Crea e inicializa adorcsAuse
    Set adorcsAus = New ADODB.Recordset
    adorcsAus.ActiveConnection = Conn
    adorcsAus.CursorLocation = adUseServer
    adorcsAus.CursorType = adOpenForwardOnly
    adorcsAus.LockType = adLockReadOnly

    Do While Not AdoRcsUsuarios.EOF
    
        siProporcion = 1
        
        dFechaUltimoPago = AdoRcsUsuarios!Fechaultimopago
        
        iUltimoDiadelMes = Day(UltimoDiaDelMes(dFechaUltimoPago))
        iDiasProporcion = iUltimoDiadelMes - Day(dFechaUltimoPago) + 1
        
        'Para el pago proporcional de los dias
        If iDiasProporcion > 1 Then
            siProporcion = (1 / iUltimoDiadelMes) * iDiasProporcion
        End If
        
        dPeriodo = dFechaUltimoPago
        dPeriodoIni = dPeriodo
        
        'Si la ultima fecha de pago no coincide con el el
        'ultimo dia del mes (siProporcion < 1 ) ajusta la fecha
        'para calcular el periodo completo y despues multiplica la cuota
        'por el factor de proporcion.
        
        If siProporcion < 1 Then
            
            dPeriodo = UltimoDiaDelMes(DateAdd("m", -1, dPeriodo))
            dPeriodoIni = dPeriodo
            
            dFechaUltimoPago = dPeriodo
            
            
        End If
        
        
        
        
        'lPeriodos = CalculaPeriodos(dFechaUltimopago, Me.dtpFechaCalc.Value, AdoRcsUsuarios!Periodo)
        lPeriodos = CalculaPeriodos(dFechaUltimoPago, Me.dtpFechaCalc.Value, iPeriodoCalculo)
        
            
        For i = 1 To lPeriodos
        
            If i > 1 Then
                siProporcion = 1
            End If
            
            dPeriodoIni = CDate(dPeriodo + 1)
            'dPeriodo = DateAdd("m", i * AdoRcsUsuarios!Periodo, dFechaUltimopago)
            dPeriodo = DateAdd("m", i * iPeriodoCalculo, dFechaUltimoPago)
            dPeriodo = UltimoDiaDelMes(dPeriodo)
            
            #If SqlServer_ Then
                strSQL = "SET DATEFORMAT mdy" & vbCrLf
                strSQL = strSQL & " SELECT MONTO, MONTODESCUENTO"
                strSQL = strSQL & " FROM HISTORICO_CUOTAS"
                strSQL = strSQL & " WHERE IdTipoUsuario=" & AdoRcsUsuarios!idtipousuario
                strSQL = strSQL & " AND Periodo=" & iPeriodoCalculo
                strSQL = strSQL & " AND VigenteDesde <=" & "'" & Format(dPeriodo, "yyyymmdd") & "'"
                strSQL = strSQL & " AND VigenteHasta >=" & "'" & Format(dPeriodo, "yyyymmdd") & "'"
            #Else
                strSQL = "SELECT MONTO, MONTODESCUENTO"
                strSQL = strSQL & " FROM HISTORICO_CUOTAS"
                strSQL = strSQL & " WHERE IdTipoUsuario=" & AdoRcsUsuarios!idtipousuario
                strSQL = strSQL & " AND Periodo=" & iPeriodoCalculo
                strSQL = strSQL & " AND VigenteDesde <=" & "#" & dPeriodo & "#"
                strSQL = strSQL & " AND VigenteHasta >=" & "#" & dPeriodo & "#"
            #End If
            
            adoRcsFac.Open strSQL
                
            If adoRcsFac.EOF Then
                dCuota = 0
                dCantidad = 1
                dDescuentoPor = 0
                dIvaPor = 0
                lFormaPago = 0
            Else
                If Me.chkDireccionar.Value Then
                    dCuota = adoRcsFac!MontoDescuento
                Else
                    dCuota = adoRcsFac!Monto
                End If
                dCantidad = 1
                dDescuentoPor = 0
                
            End If
            
            dIvaPor = AdoRcsUsuarios!Impuesto1 / 100
            lFormaPago = AdoRcsUsuarios!Periodo
            
            adoRcsFac.Close
                    
            'Para ausencias
            #If SqlServer_ Then
                strSQL = "SELECT FechaInicial, FechaFinal, Porcentaje"
                strSQL = strSQL & " FROM AUSENCIAS"
                strSQL = strSQL & " WHERE IdMember=" & AdoRcsUsuarios!Idmember
                strSQL = strSQL & " AND FechaInicial <='" & Format(dPeriodoIni, "yyyymmdd") & "'"
                'strSQL = strSQL & " AND FechaFinal >='" & Format(dPeriodoIni, "mm/dd/yyyy") & "'"
            #Else
                strSQL = "SELECT FechaInicial, FechaFinal, Porcentaje"
                strSQL = strSQL & " FROM AUSENCIAS"
                strSQL = strSQL & " WHERE IdMember=" & AdoRcsUsuarios!Idmember
                strSQL = strSQL & " AND FechaInicial <=#" & Format(dPeriodoIni, "mm/dd/yyyy") & "#"
                'strSQL = strSQL & " AND FechaFinal >=#" & Format(dPeriodoIni, "mm/dd/yyyy") & "#"
            #End If
            adorcsAus.Open strSQL
            
            If adorcsAus.EOF Then
                nPorAus = 0
            Else
                nPorAus = adorcsAus!Porcentaje / 100
            End If
            
            adorcsAus.Close
                
            dMonto = dCuota * dCantidad * (1 - nPorAus)
            
            'Toma en cuenta el factor de proporcion
            
            dMonto = Round(dMonto * siProporcion, 2)
            
            dIntereses = CalculaInteresesPasados(dPeriodo, Date, dMonto, sngPorInter)
            dIntereses = dIntereses + CalculaInteresesActuales(dPeriodo, Date, dMonto, sngPorInter)
            dDescuento = 0
                
            dIva = dMonto - Round(dMonto / (1 + dIvaPor), 2)
            dIvaDescuento = dDescuento - Round(dDescuento / (1 + dIvaPor), 2)
            dIvaIntereses = dIntereses - Round(dIntereses / (1 + dIvaPor), 2)
                
            dImporte = dMonto + dIntereses
                
                
            dTotal = dImporte - dDescuento
                
                
                
            siTotImporte = siTotImporte + dImporte
            siTotDescuento = siTotDescuento + dDescuento
            siTotIva = siTotIva + dIva
            siTotTotal = siTotTotal + dTotal
                
            'Columnas del grid
            '0  Concepto
            '1  Nombre
            '2  Periodo
            '3  Cantidad
            '4  Importe
            '5  Intereses
            '6  Descuento
            '7  Total
            '8  Clave
            '9  IvaPor          No Visible
            '10  Iva            No Visible
            '11 IvaDescuento    No Visible
            '12 IvaIntereses    No Visible
            '13 DescMonto       No Visible
            '14 IdMember        No Visible
            '15 NoFamiliar      No Visible
            '16 Periodo         No Visible
            '17 IdTipoUsuario   No Visible
            '18 TipoCargo       No Visible
            '19 Auxiliar        No Visible
                
            'sCargos(lIndexsCargos) = Format(Date2Days(dPeriodo) + Val(AdoRcsUsuarios!NumeroFamiliar), "0000000000") & vbTab & AdoRcsUsuarios.Fields("Tipo_Usuario.Descripcion") & IIf(nPorAus > 0, "/AUS", "") & IIf(siProporcion < 1, "/PROP. " & iDiasProporcion & " DIAS", "") & IIf(Me.chkDireccionar.Value = 1, " DIRECCIONADO", "") & vbTab & Trim(AdoRcsUsuarios!a_paterno) & " " & Trim(AdoRcsUsuarios!a_materno) & " " & Trim(AdoRcsUsuarios!Nombre) & vbTab & dPeriodo & vbTab & dCantidad & vbTab & dMonto & vbTab & dIntereses & vbTab & dDescuento & vbTab & dImporte & vbTab & AdoRcsUsuarios!Idconcepto & vbTab & dIvaPor & vbTab & dIva & vbTab & dIvaDescuento & vbTab & dIvaIntereses & vbTab & 0 & vbTab & AdoRcsUsuarios!IdMember & vbTab & AdoRcsUsuarios!NumeroFamiliar & vbTab & lFormaPago & vbTab & AdoRcsUsuarios!idtipousuario & vbTab & 0 & vbTab & vbNullString & vbTab & AdoRcsUsuarios!FacORec
            'sCargos(lIndexsCargos) = Format(Date2Days(dPeriodo) + Val(AdoRcsUsuarios!NumeroFamiliar), "0000000000") & vbTab & AdoRcsUsuarios.Fields("Tipo_Usuario.Descripcion") & IIf(nPorAus > 0, "/AUS", "") & IIf(siProporcion < 1, "/PROP. " & iDiasProporcion & " DIAS", "") & IIf(Me.chkDireccionar.Value = 1, " DIRECCIONADO", "") & vbTab & Trim(AdoRcsUsuarios!a_paterno) & " " & Trim(AdoRcsUsuarios!a_materno) & " " & Trim(AdoRcsUsuarios!Nombre) & vbTab & dPeriodo & vbTab & dCantidad & vbTab & dMonto & vbTab & dIntereses & vbTab & dDescuento & vbTab & dImporte & vbTab & AdoRcsUsuarios!Idconcepto & vbTab & dIvaPor & vbTab & dIva & vbTab & dIvaDescuento & vbTab & dIvaIntereses & vbTab & 0 & vbTab & AdoRcsUsuarios!IdMember & vbTab & AdoRcsUsuarios!NumeroFamiliar & vbTab & iPeriodoCalculo & vbTab & AdoRcsUsuarios!idtipousuario & vbTab & 0 & vbTab & vbNullString & vbTab & AdoRcsUsuarios!FacORec
            aFacturacion(lIndexFacturacion) = Format(Date2Days(dPeriodo) + Val(AdoRcsUsuarios!NumeroFamiliar), "0000000000") & vbTab & AdoRcsUsuarios.Fields("Tipo_Usuario.Descripcion") & IIf(nPorAus > 0, "/AUS", "") & IIf(siProporcion < 1, "/PROP. " & iDiasProporcion & " DIAS", "") & IIf(Me.chkDireccionar.Value = 1, " DIRECCIONADO", "") & vbTab & Trim(AdoRcsUsuarios!A_Paterno) & " " & Trim(AdoRcsUsuarios!A_Materno) & " " & Trim(AdoRcsUsuarios!Nombre) & vbTab & dPeriodo & vbTab & dCantidad & vbTab & dMonto & vbTab & dIntereses & vbTab & dDescuento & vbTab & dImporte & vbTab & AdoRcsUsuarios!IdConcepto & vbTab & dIvaPor & vbTab & dIva & vbTab & dIvaDescuento & vbTab & dIvaIntereses & vbTab & 0 & vbTab & AdoRcsUsuarios!Idmember & vbTab & AdoRcsUsuarios!NumeroFamiliar & vbTab & iPeriodoCalculo & vbTab & AdoRcsUsuarios!idtipousuario & vbTab & 0 & vbTab & vbNullString & vbTab & AdoRcsUsuarios!FacORec & vbTab & "0" & vbTab & AdoRcsUsuarios!Unidad
                
                
            'lIndexsCargos = lIndexsCargos + 1
            lIndexFacturacion = lIndexFacturacion + 1
        Next
        AdoRcsUsuarios.MoveNext
    Loop
        
        
    Set adoRcsFac = Nothing
    
    Set adorcsAus = Nothing
        
    AdoRcsUsuarios.Close
    Set AdoRcsUsuarios = Nothing
End Sub

Private Sub CalculaCargosVarios(lidMember As Long, dFechacalculo As Date)
    
    Dim adorcsCargosVarios As ADODB.Recordset
    
    
    
    dIntereses = 0
    dDescuento = 0
    dImporte = 0
    dIva = 0
    dIvaPor = 0
    
    #If SqlServer_ Then
        strSQL = "SELECT CARGOS_VARIOS.IdMember, IdCargoVario, CARGOS_VARIOS.IdConcepto, CARGOS_VARIOS.DescripcionCargo, Ordinal, NumeroDeCargos, FechaVencimiento,"
        strSQL = strSQL & "  Importe, Descripcion, Impuesto1, Impuesto2, FacORec, Unidad"
        strSQL = strSQL & " FROM CARGOS_VARIOS INNER JOIN CONCEPTO_INGRESOS ON CARGOS_VARIOS.IdConcepto=CONCEPTO_INGRESOS.IdConcepto"
        strSQL = strSQL & " WHERE CARGOS_VARIOS.IdMember=" & lidMember
        strSQL = strSQL & " AND CARGOS_VARIOS.FechaVencimiento <= '" & Format(dFechacalculo, "yyyymmdd") & "'"
        strSQL = strSQL & " AND CARGOS_VARIOS.Pagado=0"
        strSQL = strSQL & " ORDER BY IdCargoVario, FechaVencimiento"
    #Else
        strSQL = "SELECT CARGOS_VARIOS.IdMember, IdCargoVario, CARGOS_VARIOS.IdConcepto, CARGOS_VARIOS.DescripcionCargo, Ordinal, NumeroDeCargos, FechaVencimiento,"
        strSQL = strSQL & "  Importe, Descripcion, Impuesto1, Impuesto2, FacORec, Unidad"
        strSQL = strSQL & " FROM CARGOS_VARIOS INNER JOIN CONCEPTO_INGRESOS ON CARGOS_VARIOS.IdConcepto=CONCEPTO_INGRESOS.IdConcepto"
        strSQL = strSQL & " WHERE CARGOS_VARIOS.IdMember=" & lidMember
        strSQL = strSQL & " AND CARGOS_VARIOS.FechaVencimiento <= #" & Format(dFechacalculo, "mm/dd/yyyy") & "#"
        strSQL = strSQL & " AND CARGOS_VARIOS.Pagado=0"
        strSQL = strSQL & " ORDER BY IdCargoVario, FechaVencimiento"
    #End If
    
    Set adorcsCargosVarios = New ADODB.Recordset
    adorcsCargosVarios.CursorLocation = adUseServer
    
    adorcsCargosVarios.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    Do Until adorcsCargosVarios.EOF
        dMonto = adorcsCargosVarios!Importe
        dIvaPor = adorcsCargosVarios!Impuesto1 / 100
        dIva = dMonto - Round(dMonto / (1 + dIvaPor), 2)
        dIntereses = 0
        dDescuento = 0
        
        dIvaDescuento = dDescuento - Round(dDescuento / (1 + dIvaPor), 2)
        dIvaIntereses = dIntereses - Round(dIntereses / (1 + dIvaPor), 2)
                
        dImporte = dMonto + dIntereses
        
        dTotal = dImporte - dDescuento
        
        '                                                                                                       0                                                  1                      2                                             3           4                5                    6                    7                   8                                      9                 10             11                      12                      13          14                                    15          16          17          18          19                     20
        sCargos(lIndexsCargos) = Format(Date2Days(adorcsCargosVarios!FechaVencimiento), "0000000000") & vbTab & adorcsCargosVarios.Fields("DescripcionCargo") & vbTab & vbNullString & vbTab & adorcsCargosVarios!FechaVencimiento & vbTab & 1 & vbTab & dMonto & vbTab & dIntereses & vbTab & dDescuento & vbTab & dImporte & vbTab & adorcsCargosVarios!IdConcepto & vbTab & dIvaPor & vbTab & dIva & vbTab & dIvaDescuento & vbTab & dIvaIntereses & vbTab & 0 & vbTab & adorcsCargosVarios!Idmember & vbTab & 0 & vbTab & 1 & vbTab & 0 & vbTab & 2 & vbTab & adorcsCargosVarios!IdCargoVario & vbTab & adorcsCargosVarios!FacORec & vbTab & "0" & vbTab & adorcsCargosVarios!Unidad
        
        siTotImporte = siTotImporte + dImporte
        siTotDescuento = siTotDescuento + dDescuento
        siTotIva = siTotIva + dIva
        siTotTotal = siTotTotal + dTotal

        
        adorcsCargosVarios.MoveNext
        lIndexsCargos = lIndexsCargos + 1
    Loop
    
    adorcsCargosVarios.Close
    Set adorcsCargosVarios = Nothing
    
End Sub

Private Sub CalculaCargosMembresia(lidMember As Long, dFechacalculo As Date)
    
    Dim adorcsCargoMem As ADODB.Recordset
    Dim sNumeroPago As String
    
    
    dIntereses = 0
    dDescuento = 0
    dImporte = 0
    dIva = 0
    dIvaPor = 0
    
    #If SqlServer_ Then
        strSQL = "SELECT MEMBRESIAS.IdMembresia, MEMBRESIAS.idMember, DETALLE_MEM.NoPago, MEMBRESIAS.NumeroPagos, DETALLE_MEM.Monto, DETALLE_MEM.FechaVence, DETALLE_MEM.idReg, "
        strSQL = strSQL & " CONCEPTO_INGRESOS.IdConcepto, CONCEPTO_INGRESOS.Descripcion, CONCEPTO_INGRESOS.Impuesto1, CONCEPTO_INGRESOS.Impuesto2, CONCEPTO_INGRESOS.FacORec, CONCEPTO_INGRESOS.Unidad"
        strSQL = strSQL & " FROM DETALLE_MEM INNER JOIN MEMBRESIAS ON DETALLE_MEM.IdMembresia=MEMBRESIAS.IdMembresia, CONCEPTO_INGRESOS"
        strSQL = strSQL & " WHERE MEMBRESIAS.IdMember=" & lidMember
        strSQL = strSQL & " AND DETALLE_MEM.FechaVence <= '" & Format(dFechacalculo, "yyyymmdd") & "'"
        strSQL = strSQL & " AND DETALLE_MEM.FechaPago IS NULL"
        strSQL = strSQL & " AND CONCEPTO_INGRESOS.IdConcepto=" & "903"
        strSQL = strSQL & " AND DETALLE_MEM.Monto <> 0"
        strSQL = strSQL & " ORDER BY DETALLE_MEM.NoPago"
    #Else
        strSQL = "SELECT MEMBRESIAS.IdMembresia, MEMBRESIAS.idMember, DETALLE_MEM.NoPago, MEMBRESIAS.NumeroPagos, DETALLE_MEM.Monto, DETALLE_MEM.FechaVence, DETALLE_MEM.idReg, "
        strSQL = strSQL & " CONCEPTO_INGRESOS.IdConcepto, CONCEPTO_INGRESOS.Descripcion, CONCEPTO_INGRESOS.Impuesto1, CONCEPTO_INGRESOS.Impuesto2, CONCEPTO_INGRESOS.FacORec, CONCEPTO_INGRESOS.Unidad"
        strSQL = strSQL & " FROM DETALLE_MEM INNER JOIN MEMBRESIAS ON DETALLE_MEM.IdMembresia=MEMBRESIAS.IdMembresia, CONCEPTO_INGRESOS"
        strSQL = strSQL & " WHERE MEMBRESIAS.IdMember=" & lidMember
        strSQL = strSQL & " AND DETALLE_MEM.FechaVence <= #" & Format(dFechacalculo, "mm/dd/yyyy") & "#"
        strSQL = strSQL & " AND IsNull(DETALLE_MEM.FechaPago)"
        strSQL = strSQL & " AND CONCEPTO_INGRESOS.IdConcepto=" & "903"
        strSQL = strSQL & " AND DETALLE_MEM.Monto <> 0"
        strSQL = strSQL & " ORDER BY DETALLE_MEM.NoPago"
    #End If
    
    Set adorcsCargoMem = New ADODB.Recordset
    adorcsCargoMem.CursorLocation = adUseServer
    
    adorcsCargoMem.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    Do Until adorcsCargoMem.EOF
        dMonto = adorcsCargoMem!Monto
        dIvaPor = adorcsCargoMem!Impuesto1 / 100
        dIva = dMonto - Round(dMonto / (1 + dIvaPor), 2)
        dIntereses = 0
        dDescuento = 0
        
        dIvaDescuento = dDescuento - Round(dDescuento / (1 + dIvaPor), 2)
        dIvaIntereses = dIntereses - Round(dIntereses / (1 + dIvaPor), 2)
                
        dImporte = dMonto + dIntereses
        
        dTotal = dImporte - dDescuento
        
        If adorcsCargoMem!NoPago = 0 Then
            If adorcsCargoMem!NumeroPagos = 0 Then
                sNumeroPago = "CONTADO"
            Else
                sNumeroPago = "INV. INICIAL"
            End If
        Else
            sNumeroPago = Format(adorcsCargoMem!NoPago, "00") & "/" & Format(adorcsCargoMem!NumeroPagos, "00")
        End If
        '                                                                                                       0                                                  1                      2                                             3           4                5                    6                    7                   8                                      9                 10             11                      12                      13          14                                    15          16          17          18          19                     20
        sCargos(lIndexsCargos) = Format(Date2Days(adorcsCargoMem!FechaVence), "0000000000") & vbTab & adorcsCargoMem.Fields("Descripcion") & " " & sNumeroPago & vbTab & vbNullString & vbTab & Format(adorcsCargoMem!FechaVence, "dd/mm/yyyy") & vbTab & 1 & vbTab & dMonto & vbTab & dIntereses & vbTab & dDescuento & vbTab & dImporte & vbTab & adorcsCargoMem!IdConcepto & vbTab & dIvaPor & vbTab & dIva & vbTab & dIvaDescuento & vbTab & dIvaIntereses & vbTab & 0 & vbTab & adorcsCargoMem!Idmember & vbTab & 0 & vbTab & 1 & vbTab & 0 & vbTab & 3 & vbTab & adorcsCargoMem!IdReg & vbTab & adorcsCargoMem!FacORec & vbTab & "0" & vbTab & adorcsCargoMem!Unidad
        
        siTotImporte = siTotImporte + dImporte
        siTotDescuento = siTotDescuento + dDescuento
        siTotIva = siTotIva + dIva
        siTotTotal = siTotTotal + dTotal

        
        adorcsCargoMem.MoveNext
        lIndexsCargos = lIndexsCargos + 1
    Loop
    
    adorcsCargoMem.Close
    Set adorcsCargoMem = Nothing

End Sub

Private Sub LlenaComboEstados()
    Dim adorcsEstados As ADODB.Recordset
    
    strSQL = "SELECT cveEntFederativa, nomEntFederativa"
    strSQL = strSQL & " FROM ENTFEDERATIVA"
    strSQL = strSQL & " ORDER BY nomEntFederativa"
    
    
    Set adorcsEstados = New ADODB.Recordset
    adorcsEstados.CursorLocation = adUseServer
    
    adorcsEstados.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    Do While Not adorcsEstados.EOF
        Me.ssCmbFacEstado.AddItem adorcsEstados!nomEntFederativa & vbTab & adorcsEstados!cveEntFederativa
        adorcsEstados.MoveNext
    Loop
    
    adorcsEstados.Close
    Set adorcsEstados = Nothing
    
End Sub

Private Sub LlenaComboFormaDireccionar()

    Me.cmbTipoDireccionar.AddItem "T.CREDITO"
    Me.cmbTipoDireccionar.AddItem "T.DEBITO"

End Sub

Private Sub LlenaComboPeriodoCalculo()

    Me.cmbPeriodoCalculo.AddItem "MENSUAL"
    Me.cmbPeriodoCalculo.ItemData(Me.cmbPeriodoCalculo.NewIndex) = 1
    Me.cmbPeriodoCalculo.AddItem "ANUAL"
    Me.cmbPeriodoCalculo.ItemData(Me.cmbPeriodoCalculo.NewIndex) = 12

End Sub

Private Sub ChecaMensajes(lIdTitular As Long)
    Dim frmMen As frmMensajes
    Dim adorcsCheca As ADODB.Recordset
    Dim boSiHay As Boolean
    
    
    boSiHay = False
    
    strSQL = "SELECT IdMensaje"
    strSQL = strSQL & " FROM MENSAJES"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & " IdMember=" & lIdTitular
    strSQL = strSQL & " AND Leido=0"
    
    
    
    Set adorcsCheca = New ADODB.Recordset
    adorcsCheca.CursorLocation = adUseServer
    
    adorcsCheca.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not adorcsCheca.EOF Then boSiHay = True
    
    adorcsCheca.Close
    Set adorcsCheca = Nothing
    
    
    
    If boSiHay Then
        Set frmMen = New frmMensajes
    
        Load frmMen
    
        frmMen.lidMember = lIdTitular
        frmMen.Show vbModal
    End If
    
End Sub
