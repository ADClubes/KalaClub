VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Toalla"
      Height          =   615
      Left            =   1800
      TabIndex        =   1
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clase"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
    Dim lI As Long
    Dim lNoCupones As Long
    
    lNoCupones = 5
    
    
    For lI = 1 To lNoCupones
    
    
        Printer.PaintPicture LoadPicture("d:\kalaclub\recursos\logo_sportium_bn.jpg"), 600, 0, 3000, 1000
        
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        
        Printer.FontSize = 10
        Printer.FontBold = True
        Printer.Print "SPORTIUM COYOACAN"
        Printer.Print
        Printer.FontSize = 8
        Printer.FontBold = False
        Printer.Print "Válido por un(a) Clase de Natación"
        Printer.Print "Instructor(a): " & "ROMAN AGUILAR"
        Printer.Print "Vale " & lI; " de " & lNoCupones
        Printer.Print "Emitido: " & Format(Date, "dd/mmm/yy") & " Válido hasta: " & Format(Date + 45, "dd/mmm/yy")
        Printer.Print
        Printer.Print "Recibo #: " & "00234"
        Printer.Print
        Printer.EndDoc
        
    Next
    
    
End Sub

Private Sub Command2_Click()
Dim lI As Long
    Dim lNoCupones As Long
    
    lNoCupones = 5
    
    
    For lI = 1 To lNoCupones
    
    
        Printer.PaintPicture LoadPicture("d:\kalaclub\recursos\logo_sportium_bn.jpg"), 600, 0, 3000, 1000
        
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        
        Printer.FontSize = 10
        Printer.FontBold = True
        Printer.Print "SPORTIUM COYOACAN"
        Printer.Print
        Printer.FontSize = 8
        Printer.FontBold = False
        Printer.Print "VALIDO POR UNA TOALLA"
        Printer.Print "Vale " & lI; " de " & lNoCupones
        Printer.Print "Emitido: " & Format(Date, "dd/mmm/yy") & " Válido hasta: " & Format(Date + 45, "dd/mmm/yy")
        Printer.Print
        Printer.Print "Recibo #: " & "00234"
        Printer.FontSize = 7
        Printer.Print "-El usuario deberá devolver la toalla el mismo dia"
        Printer.Print "-El usuario deberá pagar el importe de la toalla si:"
        Printer.Print "-Está rota o manchada de tinte o grasa para zapatos"
        Printer.Print "-Es extraviada o no se devuelve al personal por cualquier"
        Printer.Print " otro motivo"
        Printer.Print
        Printer.EndDoc
        
    Next

End Sub
