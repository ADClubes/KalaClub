VERSION 5.00
Begin VB.Form frmDescFac 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Aplicar descuento"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkTodos 
      Caption         =   "Aplicar a todos los renglones"
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   1200
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   1988
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdAplicar 
      Caption         =   "Aplicar"
      Default         =   -1  'True
      Height          =   495
      Left            =   548
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtCantidad 
      Height          =   285
      Left            =   2160
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.OptionButton optPorPorcen 
      Caption         =   " Porcentaje"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.OptionButton optPorMonto 
      Caption         =   "Monto"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "frmDescFac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdAplicar_Click()

    Dim dPorcenDescuento As Double
    Dim dMontoDescuento As Double
    
    Dim lCurRow As Long
    Dim lI As Long
    
    
    If Me.optPorMonto.Value And Val(Me.txtCantidad.Text) > frmFacturacion.ssdbgFactura.Columns(5).Value Then
        MsgBox "El descuento no puede ser mayor que el importe!", vbCritical, "Facturacion"
        Me.txtCantidad.SetFocus
        Exit Sub
    End If
    
    If Me.optPorPorcen.Value And Val(Me.txtCantidad.Text) > 100 Then
        MsgBox "El descuento no puede ser mayor del 100%!", vbCritical, "Facturacion"
        Me.txtCantidad.SetFocus
        Exit Sub
    End If
    
    If Me.optPorPorcen.Value Then
        dPorcenDescuento = Round(Val(Me.txtCantidad.Text), 2)
    Else
        dPorcenDescuento = (Round(Val(Me.txtCantidad.Text), 2) / frmFacturacion.ssdbgFactura.Columns(5).Value) * 100
    End If
    
    If Me.chkTodos.Value Then
        lCurRow = frmFacturacion.ssdbgFactura.Row
        For lI = 0 To frmFacturacion.ssdbgFactura.Rows - 1
            frmFacturacion.ssdbgFactura.Row = lI
            frmFacturacion.ssdbgFactura.Columns("Descuento").Value = dPorcenDescuento
            RecalculaRenglon
        Next
        frmFacturacion.ssdbgFactura.Row = lCurRow
    Else
        frmFacturacion.ssdbgFactura.Columns("Descuento").Value = dPorcenDescuento
        RecalculaRenglon
    End If
    frmFacturacion.ssdbgFactura.Update
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Me.txtCantidad.SetFocus
End Sub

Private Sub Form_Load()
    Me.Height = 2760
    Me.Width = 3840
    
    Me.optPorPorcen.Value = True
    
    
    CENTRAFORMA MDIPrincipal, Me
    
    
End Sub



Private Sub TxtCantidad_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8 ' Tecla backspace
            KeyAscii = KeyAscii
        Case 46 'punto decimal
            If InStr(Me.txtCantidad.Text, ".") Then
                KeyAscii = 0
            Else
                KeyAscii = KeyAscii
            End If
        Case 48 To 57 ' Numeros del 0 al 9
            KeyAscii = KeyAscii
        Case Else
            KeyAscii = 0
            MsgBox "Solo admite números", vbInformation, "Productos"
    End Select
    
End Sub
