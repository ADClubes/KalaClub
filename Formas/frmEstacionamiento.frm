VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmEstacionamiento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calcula Estacionamiento"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   6645
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdIns 
      Caption         =   "Insertar"
      Height          =   495
      Left            =   5280
      TabIndex        =   5
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Importe a pagar"
      Height          =   1215
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   3375
      Begin VB.Label lblCtrl 
         Alignment       =   2  'Center
         Height          =   255
         Index           =   4
         Left            =   1920
         TabIndex        =   11
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblImporte 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.TextBox txtCtrl 
      Height          =   375
      Index           =   1
      Left            =   1680
      MaxLength       =   8
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox txtCtrl 
      Height          =   375
      Index           =   0
      Left            =   3360
      MaxLength       =   4
      TabIndex        =   2
      Top             =   330
      Width           =   615
   End
   Begin VB.CommandButton cmdCalcula 
      Caption         =   "Calcular"
      Default         =   -1  'True
      Height          =   495
      Left            =   3960
      TabIndex        =   4
      Top             =   2400
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd/mm/yyyy HH:mm"
      Format          =   101187585
      CurrentDate     =   39941
   End
   Begin VB.Label lblCtrl 
      Alignment       =   2  'Center
      Caption         =   "# Folio"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   8
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblCtrl 
      Alignment       =   2  'Center
      Caption         =   "Hora"
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   7
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblCtrl 
      Alignment       =   2  'Center
      Caption         =   "Fecha"
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblCtrl 
      Caption         =   "Entrada:"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "frmEstacionamiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dImporte As Double

Private Sub cmdCalcula_Click()
    Dim lHoras As Long
    Dim lMinutos As Long
    Dim lFraccion As Long
    Dim lDiferencia  As Long
    
    Dim dFechaInicial As Date
    Dim dFechaFinal As Date
    
    
    Dim iMinutos1 As Integer
    Dim iMinutos2 As Integer
    Dim iMinutos3 As Integer
    
    Dim dCosto1 As Double
    Dim dCosto2 As Double
    Dim dCosto3 As Double
    
    Dim iParcial1 As Integer
    Dim iParcial2 As Integer
    Dim iParcial3 As Integer
    
    
    
    
    iMinutos1 = 195
    iMinutos2 = 15
    iMinutos3 = 15
    
    dCosto1 = 15
    dCosto2 = 5
    dCosto3 = dCosto2 / 4
    
    
    
    If Len(Me.txtCtrl(0).Text) < 4 Then
        MsgBox "¡Capturar la hora en formato HHMM!", vbExclamation, "Verifique"
        Me.txtCtrl(0).SetFocus
        Exit Sub
    End If
    
    
    lHoras = Val(Left$(Me.txtCtrl(0).Text, 2))
    lMinutos = Val(Right$(Me.txtCtrl(0).Text, 2))
    
    
    If lHoras > 23 Then
        MsgBox "La hora debe estar entre 00 y 23", vbExclamation, "Verifique"
        Me.txtCtrl(0).SetFocus
        Exit Sub
    End If
    
    If lMinutos > 59 Then
        MsgBox "Los minutos debe estar entre 00 y 59", vbExclamation, "Verifique"
        Me.txtCtrl(0).SetFocus
        Exit Sub
    End If
    
    
    
    dFechaInicial = CDate(Me.DTPicker1.Value & " " & lHoras & ":" & lMinutos)
    dFechaFinal = Now
    
    
    lDiferencia = DateDiff("n", dFechaInicial, dFechaFinal)
    
    
    If lDiferencia < 0 Then
        MsgBox "¡La hora de entrada tiene que ser menor que la de salida!", vbExclamation, "Verifique"
        Me.txtCtrl(0).SetFocus
        Exit Sub
    End If
    
    
    If lDiferencia > iMinutos1 Then
        iParcial1 = 1
        lDiferencia = lDiferencia - iMinutos1
    Else
        iParcial1 = 1
        lDiferencia = 0
    End If
    
    
    
    
    If lDiferencia > iMinutos2 Then
        iParcial2 = Int(lDiferencia / iMinutos2)
        lDiferencia = lDiferencia - iParcial2 * iMinutos2
    ElseIf lDiferencia > 0 Then
        iParcial2 = 1
        lDiferencia = 0
    End If
    
    iParcial3 = Int(lDiferencia / iMinutos3)
    
    If lDiferencia Mod iMinutos3 > 0 Then
        iParcial3 = iParcial3 + 1
    End If
    
    
    dImporte = dCosto1 * iParcial1 + dCosto2 * iParcial2
    
    Me.lblImporte = Format(dImporte, "$#,##0.00")
    
    Me.lblCtrl(4).Caption = dCosto1 * iParcial1 & " + " & dCosto2 * iParcial2 & " + " & dCosto3 * iParcial3
    
    Me.txtCtrl(0).SetFocus
    
End Sub



Private Sub cmdIns_Click()

    Dim sCadGrid As String
    Dim sDescripcion As String
    Dim dMonto As Double
    Dim dIvaPor As Double
    Dim sFacORec As String
    Dim sUnidad As String
    Dim lConceptoIngreso As Long

    If dImporte = 0 Then
        Exit Sub
    End If
    
    If Me.txtCtrl(1).Text = vbNullString Then
        MsgBox "¡Es necesario capturar el # de folio!", vbExclamation, "Verifique"
        Me.txtCtrl(1).SetFocus
        Exit Sub
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
    '9  IvaPor
    '10 IVA
    '11 IvaDescuento
    '12 IvaInteres
    '13 DescMonto
    '14 IdMember
    '15 NoFamiliar
    '16 FormaPago
    '17 IdTipoUsuario
    '18 TipoCargo
    '19 Auxiliar
    '20 FacoRec
    '21 IdInstructor
    
    
    
    lConceptoIngreso = Val(ObtieneParametro("CONCEPTO_ESTACIONAMIENTO"))
    
    If lConceptoIngreso = 0 Then
        MsgBox "No hay concepto de ingresos configurado", vbExclamation, "Error"
        Exit Sub
    End If
    
    If (Not ObtieneDatosConceptoIngresos(lConceptoIngreso, sDescripcion, dMonto, dIvaPor, sFacORec, sUnidad)) Then
        sDescripcion = "ESTACIONAMIENTO"
        dIvaPor = 0
    End If
    
    sCadGrid = vbNullString
    

        sCadGrid = sCadGrid & sDescripcion & vbTab
        sCadGrid = sCadGrid & "" & vbTab
        sCadGrid = sCadGrid & Format(Date, "dd/mm/yyyy") & vbTab
        sCadGrid = sCadGrid & 1 & vbTab
        sCadGrid = sCadGrid & dImporte & vbTab
        sCadGrid = sCadGrid & 0 & vbTab
        sCadGrid = sCadGrid & 0 & vbTab
        sCadGrid = sCadGrid & dImporte & vbTab
        sCadGrid = sCadGrid & lConceptoIngreso & vbTab
        sCadGrid = sCadGrid & dIvaPor / 100 & vbTab
        sCadGrid = sCadGrid & Round(dImporte - (dImporte / ((100 + dIvaPor) / 100)), 2) & vbTab
        sCadGrid = sCadGrid & 0 & vbTab
        sCadGrid = sCadGrid & 0 & vbTab
        sCadGrid = sCadGrid & 0 & vbTab
        sCadGrid = sCadGrid & 0 & vbTab
        sCadGrid = sCadGrid & 1 & vbTab
        sCadGrid = sCadGrid & 1 & vbTab
        sCadGrid = sCadGrid & 0 & vbTab
        sCadGrid = sCadGrid & 5 & vbTab
        sCadGrid = sCadGrid & Trim(Me.txtCtrl(1).Text) & vbTab
        sCadGrid = sCadGrid & sFacORec & vbTab
        sCadGrid = sCadGrid & 0 & vbTab
        sCadGrid = sCadGrid & sUnidad
    
    
    frmFacturacion.ssdbgFactura.AddItem sCadGrid
    
    Unload Me
        
End Sub

Private Sub Form_Activate()
    CentraForma MDIPrincipal, Me
    
    Me.txtCtrl(0).SetFocus
End Sub

Private Sub Form_Load()
    Me.DTPicker1.Value = Date
    Me.txtCtrl(0).Text = Format(Now, "HhNn")
    
    Me.lblImporte = Format(0, "$#0.00")
    
    
    
    
End Sub



Private Sub txtCtrl_GotFocus(index As Integer)
    If index = 0 Then
        Me.txtCtrl(0).SelStart = 0
        Me.txtCtrl(0).SelLength = 4
    End If
End Sub

Private Sub txtCtrl_KeyPress(index As Integer, KeyAscii As Integer)
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
