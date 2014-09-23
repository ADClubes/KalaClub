VERSION 5.00
Begin VB.Form FrmUsuaActiv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Usuarios Activos"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7305
   Icon            =   "FrmUsuaActiv.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   7305
   Begin VB.ListBox List1 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   6975
   End
End
Attribute VB_Name = "FrmUsuaActiv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim rs As New ADODB.Recordset
    Dim i, j As Long
    Dim Cadena
    Dim Maquina
    Dim Usuario
    Dim Activo
    Dim Estatus

    ReDim LBTab(5) As Long
    Me.Top = 0
    Me.Left = 0
    Me.Width = 7425
    Me.Height = 4155
    LBTab(0) = 80
    LBTab(1) = 160
    LBTab(2) = 240
    Me.List1.Clear
    'Clear Tabs
    Call SendMessageArray(List1.hwnd, LB_SETTABSTOPS, 0&, 0&)
    'Set Tabs
    Call SendMessageArray(List1.hwnd, LB_SETTABSTOPS, 3&, LBTab(0))
    ' The user roster is exposed as a provider-specific schema rowset
    ' in the Jet 4 OLE DB provider.  You have to use a GUID to
    ' reference the schema, as provider-specific schemas are not
    ' listed in ADO's type library for schema rowsets
    #If SqlServer_ Then
    
    #Else
        Set rs = Conn.OpenSchema(adSchemaProviderSpecific, , "{947bb102-5d43-11d1-bdbf-00c04fb92675}")
    #End If
    'Output the list of all users in the current database.
    ' Obtiene los encabezados
    Cadena = rs.Fields(0).Name & vbTab & rs.Fields(1).Name & vbTab & rs.Fields(2).Name & vbTab & rs.Fields(3).Name
    Me.List1.AddItem Cadena
    While Not rs.EOF
        Maquina = Trim(IIf(IsNull(rs.Fields(0)), "", rs.Fields(0)))
        Usuario = Trim(IIf(IsNull(rs.Fields(1)), "", rs.Fields(1)))
        Activo = Trim(IIf(IsNull(rs.Fields(2)), "", rs.Fields(2)))
        Estatus = IIf(IsNull(rs.Fields(3)), " ", rs.Fields(3))
        If IsNull(Estatus) Then Estatus = " "
        If Activo Then Activo = "Si "
        Maquina = Left(Maquina, Len(Maquina) - 1)
        Usuario = Left(Usuario, Len(Usuario) - 1)
        Activo = Left(Activo, Len(Activo) - 1)
        Estatus = Left(Estatus, Len(Estatus) - 1)
        Cadena = Maquina & vbTab & Usuario & vbTab & Activo & vbTab & Estatus
        Me.List1.AddItem Cadena
        rs.MoveNext
    Wend
    rs.Close
    Set rs = Nothing
End Sub
