VERSION 5.00
Begin VB.Form frmAcerca 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acerca de..."
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   ClipControls    =   0   'False
   Icon            =   "frmAcerca.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSistema 
      Appearance      =   0  'Flat
      Caption         =   "Sys &Info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   5700
      Picture         =   "frmAcerca.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   0
      Tag             =   "Aceptar"
      Top             =   1605
      Width           =   990
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   5700
      Picture         =   "frmAcerca.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "Aceptar"
      Top             =   2520
      Width           =   990
   End
   Begin VB.Label lblWarning 
      Height          =   795
      Left            =   210
      TabIndex        =   6
      Top             =   1755
      Width           =   5340
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1245
      Left            =   180
      Picture         =   "frmAcerca.frx":1016
      Stretch         =   -1  'True
      Top             =   150
      Width           =   1305
   End
   Begin VB.Label lblDesarrolladoPor 
      Caption         =   "Desarrollado por: KalaSystems, S.A. de C.V."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   210
      TabIndex        =   5
      Top             =   2700
      Width           =   3975
   End
   Begin VB.Label lblDescription 
      Caption         =   "Descripción de la aplicación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   1695
      TabIndex        =   4
      Tag             =   "Descripción de la aplicación"
      Top             =   180
      Width           =   3480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      Index           =   1
      X1              =   165
      X2              =   5500
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   165
      X2              =   5500
      Y1              =   1590
      Y2              =   1590
   End
   Begin VB.Label lblVersion 
      Caption         =   "Versión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1695
      TabIndex        =   3
      Tag             =   "Versión"
      Top             =   615
      Width           =   3315
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "Advertencia: ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   1695
      TabIndex        =   2
      Tag             =   "Advertencia: ..."
      Top             =   1035
      Width           =   3900
   End
End
Attribute VB_Name = "frmAcerca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ************************************************************************
' PANTALLA ACERCA DE LÁMINAS
' Objetivo: MUESTRA UNA PANTALLA CON EL COPY RIGHT Y EL SYSINFO
' Programado por:
' Fecha: 5 DE NOVIEMBRE DE 2002
' ************************************************************************

Option Explicit

' Opciones de seguridad de claves del registro...
Const Read_Control = &H20000
Const Key_Query_Value = &H1
Const Key_Set_Value = &H2
Const Key_Create_Sub_Key = &H4
Const Key_Enumerate_Sub_Keys = &H8
Const Key_Notify = &H10
Const Key_Create_Link = &H20
Const Key_All_Access = Key_Query_Value + Key_Set_Value + _
                       Key_Create_Sub_Key + Key_Enumerate_Sub_Keys + _
                       Key_Notify + Key_Create_Link + Read_Control
                     
' Tipos ROOT de claves del registro...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                     ' Cadena Unicode terminada en nulo
Const REG_DWORD = 4                  ' Número de 32 bits

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Sub cmdSistema_Click()
      Call StartSysInfo
End Sub



Private Sub Form_Load()
    Dim Warn As String
    MDIPrincipal.Caption = "KalaClub - Acerca de KalaClub"
    With frmAcerca
        .Left = (Screen.Width / 2) - (frmAcerca.Width / 2)
        .Top = (Screen.Height / 3) - (frmAcerca.Height / 2)
    End With
    lblVersion.Caption = "Versión " & App.Major & "." & App.Minor & "." & App.Revision
'    lblTitle.Caption = App.Title
    lblDescription.Caption = App.EXEName & " ©"
    ' © <-- ansi (184)
    lblDisclaimer.Caption = "Derechos Reservados 2003-2007"
    lblWarning.Caption = "ADVERTENCIA: Este programa está protegido por las leyes " & _
                                    "de derechos de autor. La reproducción o distribución no " & _
                                    "autorizada de este programa o de cualquier parte del " & _
                                    "mismo, puede dar lugar a responsabilidades civiles y " & _
                                    "criminales, que serán perseguidas."
End Sub

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIPrincipal.Caption = "KALACLUB"
End Sub

Public Sub StartSysInfo()
On Error GoTo SysInfoErr
Dim Lnumrc As Long
Dim LstrSysInfoPath As String
  ' Intenta obtener del Registro la ruta\nombre de programa en la
  ' información del sistema...
  If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, LstrSysInfoPath) Then
  ' Intenta obtener del Registro sólo la ruta de programa en la
  ' información del sistema...
  ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, LstrSysInfoPath) Then
      ' Valida la existencia de una versión de archivo de 32 bits conocida
      If (Dir(LstrSysInfoPath & "\MSINFO32.EXE") <> "") Then
          LstrSysInfoPath = LstrSysInfoPath & "\MSINFO32.EXE"
      ' Error - No se encuentra el archivo...
      Else
          GoTo SysInfoErr
      End If
  ' Error - No se encuentra la entrada del Registro...
  Else
    GoTo SysInfoErr
  End If
  Call Shell(LstrSysInfoPath, vbNormalFocus)
  Exit Sub
SysInfoErr:
    Beep
  MsgBox "La información del sistema no está disponible en este momento", vbOKOnly
End Sub

Public Function GetKeyValue(PnumKeyRoot As Long, PstrKeyName As String, PstrSubKeyRef As String, ByRef PstrKeyVal As String) As Boolean
  Dim LnumI As Long                               ' Contador del bucle
  Dim Lnumrc As Long                              ' Código de retorno
  Dim LnumhKey As Long                            ' Controlador de una clave del registro abierta
  Dim LnumhDepth As Long                          '
  Dim LnumKeyValType As Long                      ' Tipo de datos de una clave del Registro
  Dim LstrtmpVal As String                        ' Almacenamiento temporal de un valor de clave del Registro
  Dim LnumKeyValSize As Long                      ' Tamaño de una variable de clave del Registro
  '------------------------------------------------------------
  ' Abre RegKey bajo PnumKeyRoot {LnumhKey_LOCAL_MACHINE...}
  '------------------------------------------------------------
  Lnumrc = RegOpenKeyEx(PnumKeyRoot, PstrKeyName, 0, Key_All_Access, LnumhKey) 'Abre una clave del Registro
  If (Lnumrc <> ERROR_SUCCESS) Then GoTo GetKeyError     ' Trata el error...
    LstrtmpVal = String$(1024, 0)                          ' Asigna espacio de variable
    LnumKeyValSize = 1024                                  ' Marca tamaño variable
    '------------------------------------------------------------
    ' Recupera un valor de clave del Registro...
    '------------------------------------------------------------
    Lnumrc = RegQueryValueEx(LnumhKey, PstrSubKeyRef, 0, _
                       LnumKeyValType, LstrtmpVal, LnumKeyValSize)  ' Obtiene/crea valor de clave
    If (Lnumrc <> ERROR_SUCCESS) Then GoTo GetKeyError        ' Trata los errores
    If (Asc(Mid(LstrtmpVal, LnumKeyValSize, 1)) = 0) Then         ' Win95 agrega cadena terminada en Null...
      LstrtmpVal = Left(LstrtmpVal, LnumKeyValSize - 1)             ' Encontrado Null, extraer de la cadena
    Else                                                  ' WinNT NO termina la cadena con Null...
      LstrtmpVal = Left(LstrtmpVal, LnumKeyValSize)                 ' No encontrado Null, extraer sólo la cadena
    End If
    '--------------------------------------------------------------
    ' Determina el tipo del valor de la clave para su conversión...
    '--------------------------------------------------------------
    Select Case LnumKeyValType                                       ' Busca los tipos de datos...
      Case REG_SZ                                                                ' Tipo de datos de la clave del Registro
        PstrKeyVal = LstrtmpVal                                           ' Copia el valor de cadena
      Case REG_DWORD                                                    ' Tipo de datos de clave de Registro Double Word
        For LnumI = Len(LstrtmpVal) To 1 Step -1           ' Convierte cada bit
          PstrKeyVal = PstrKeyVal + Hex(Asc(Mid(LstrtmpVal, LnumI, 1))) ' Genera valor Char. a Char.
        Next
        PstrKeyVal = Format$("&h" + PstrKeyVal)            ' Convierte Double Word en String
    End Select
    GetKeyValue = True                                                      ' Devuelve éxito
    Lnumrc = RegCloseKey(LnumhKey)                           ' Cierra la clave del Registro
    Exit Function                                                                    ' Sale
  
GetKeyError:                                                        ' Terminación cuando se ha producido un error...
  PstrKeyVal = ""                                                   ' Cadena vacía como valor devuelto
  GetKeyValue = False                                        ' Devuelve fallo
  Lnumrc = RegCloseKey(LnumhKey)             ' Cierra la clave del Registro
End Function

