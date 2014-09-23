Attribute VB_Name = "dbModule"
Public Sub LlenaSsDbGrid(ssdbg As SSOleDBGrid, adoConn As ADODB.Connection, sSqlQry As String, nCampos As Integer)
    Dim adorcs As ADODB.Recordset
    Dim lI As Integer
    Dim sString As String
    Set adorcs = Nothing
    
    Set adorcs = New ADODB.Recordset
                 
    adorcs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    ssdbg.RemoveAll
    
    Do Until adorcs.EOF
        
        sString = vbNullString
        
        For lI = 1 To nCampos
            sString = sString & adorcs.Fields(lI - 1)
            If lI < nCampos Then
                sString = sString & vbTab
            End If
        Next
        
        ssdbg.AddItem sString
        
        adorcs.MoveNext
    Loop

    adorcs.Close
    Set adorcs = Nothing
    
End Sub

Public Sub LlenaSsCombo(sscmb As SSOleDBCombo, adoConn As ADODB.Connection, sSqlQry As String, nCampos As Integer)
    Dim adorcs As ADODB.Recordset
    Dim lI As Integer
    Dim sString As String
    
    Set adorcs = New ADODB.Recordset
    
    adorcs.Open sSqlQry, adoConn, adOpenForwardOnly, adLockReadOnly
    
    
    sscmb.RemoveAll
    
    Do Until adorcs.EOF
        
        sString = vbNullString
        
        For lI = 1 To nCampos
            sString = sString & adorcs.Fields(lI - 1)
            If lI < nCampos Then
                sString = sString & vbTab
            End If
        Next
        
        sscmb.AddItem sString
        
        adorcs.MoveNext
    Loop
    
    adorcs.Close
    Set adorcs = Nothing
    
End Sub

Public Sub BuscaSSCombo(sscmb As SSOleDBCombo, sValor As String, nColumna As Integer)
    Dim lI As Long

    sscmb.MoveFirst
    
    For lI = 0 To sscmb.Rows - 1
        If sscmb.Columns(nColumna).Value = sValor Then
            'txtControl.Text = sscmb.Columns(1).Value
            Exit For
        End If
        sscmb.MoveNext
    Next


End Sub



