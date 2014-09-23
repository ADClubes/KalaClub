VERSION 5.00
Begin VB.Form frmFacModificaRen 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Modifica Renglón"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCalc 
      Caption         =   "Calculadora"
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtIntereses 
      Height          =   285
      Left            =   4200
      MaxLength       =   20
      TabIndex        =   6
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox txtImporte 
      Height          =   285
      Left            =   2040
      MaxLength       =   20
      TabIndex        =   4
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox txtCantidad 
      Height          =   285
      Left            =   240
      MaxLength       =   6
      TabIndex        =   2
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox txtConcepto 
      Height          =   285
      Left            =   240
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   6015
   End
   Begin VB.CommandButton cmdCancela 
      Caption         =   "Salir"
      Height          =   495
      Left            =   5040
      TabIndex        =   9
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   3480
      TabIndex        =   8
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Intereses"
      Height          =   255
      Left            =   4200
      TabIndex        =   7
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Importe"
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Cantidad"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Concepto"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmFacModificaRen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCalc_Click()
    
    Dim dRetVal As Double
    
    dRetVal = 0
    
    dRetVal = Shell("calc.exe")
    
    If Not CBool(dRetVal) Then
        MsgBox "Ocurrió un error al intentar cargar la calculadora", vbExclamation
    End If
    
    
    
End Sub

Private Sub cmdCancela_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()

    If Me.txtConcepto.Text = "" Then
        MsgBox "Indicar el concepto!", vbExclamation, "Facturación"
        Me.txtConcepto.SetFocus
        Exit Sub
    End If
    
    If Val(Me.txtCantidad.Text) <= 0 Then
        MsgBox "La cantidad debe ser mayor que cero!", vbExclamation, "Facturación"
        Me.txtCantidad.SetFocus
        Exit Sub
    End If
    
    If Val(Me.txtCantidad.Text) > 1000 Then
        MsgBox "La cantidad debe ser menor que 1,000!", vbExclamation, "Facturación"
        Me.txtCantidad.SetFocus
        Exit Sub
    End If
    
'    If Val(Me.txtImporte.Text) <= 0 Then
'        MsgBox "El Importe debe ser mayor que cero!", vbExclamation, "Facturación"
'        Me.txtImporte.SetFocus
'        Exit Sub
'    End If
    
    frmFacturacion.ssdbgFactura.Columns("Concepto").Value = Trim(Me.txtConcepto.Text)
    frmFacturacion.ssdbgFactura.Columns("Cant.").Value = Val(Me.txtCantidad.Text)
    frmFacturacion.ssdbgFactura.Columns("Importe").Value = Round(Val(Me.txtImporte.Text), 2)
    frmFacturacion.ssdbgFactura.Columns("Intereses").Value = Round(Val(Me.txtIntereses.Text), 2)
    frmFacturacion.ssdbgFactura.Update
    Unload Me
    
End Sub

Private Sub Form_Load()

    Me.Height = 2535
    Me.Width = 6735
    
    Me.Top = 5500
    Me.Left = 4000


    Me.txtConcepto.Text = frmFacturacion.ssdbgFactura.Columns("Concepto").Value
    Me.txtCantidad.Text = frmFacturacion.ssdbgFactura.Columns("Cant.").Value
    Me.txtImporte.Text = frmFacturacion.ssdbgFactura.Columns("Importe").Value
    Me.txtIntereses.Text = frmFacturacion.ssdbgFactura.Columns("Intereses").Value
    
    If sDB_NivelUser <> 0 And Not ConceptoModificable(Val(frmFacturacion.ssdbgFactura.Columns("Clave").Value)) Then
        Me.txtImporte.Enabled = False
        Me.txtIntereses.Enabled = False
    End If
    
End Sub



Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8 ' Tecla backspace
            KeyAscii = KeyAscii
        Case 48 To 57 ' Numeros del 0 al 9
            KeyAscii = KeyAscii
        Case Else
            KeyAscii = 0
    End Select
End Sub





Private Sub txtConcepto_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtImporte_KeyPress(KeyAscii As Integer)
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
    End Select
End Sub



Private Sub txtIntereses_KeyPress(KeyAscii As Integer)
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
    End Select
End Sub
