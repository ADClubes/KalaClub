Attribute VB_Name = "SubMain"
Sub Main()
    
    Dim Valor As Byte
    Dim Bandera As Boolean
    
    
    
    If App.PrevInstance Then
        MsgBox "¡ El programa ya está en ejecución !", vbInformation, "¡ Imposible ejecutar !"
        End
        Exit Sub
    End If
    
    
    Bandera = True
    Valor = Lee_Ini()
    
    Select Case Valor
        Case 1  'No existe INI
            If Not CREA_INI Then
                End
            End If
        Case 0    ' Existe INI y todo esta bien
            If Not Connection_DB() Then
                Bandera = False
                End
            Else
                frmPresentacion.Show vbModal
                
                If Not LoginOk Then
                    'Si la contraseña ya esta caducada
                    If ChangePassword Then
                    
                        Dim frmCamPass As frmChangePass
                    
                        Set frmCamPass = New frmChangePass
                    
                        frmCamPass.Show vbModal
                    
                    End If
                    
                
                
                    End
                End If
            End If
    End Select
    
    MDIPrincipal.Show
    
End Sub
