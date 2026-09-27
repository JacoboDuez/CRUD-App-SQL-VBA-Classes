#THis is the userform code
Private Sub btnHide_Click()
    Unload Me
End Sub
Private Sub btnRegisterNewVehicle_Click()
Dim bProceed As Boolean

bProceed = False
    
With fmCarRental
    .Hide
    If .tbDL <> "" And .tbLName <> "" And .tbName <> "" Then
        If .tbVBrand <> "" And .tbMY <> "" And .tbConfirmation <> "" Then
            If .lbVehicleModels.Value <> "" Then
                bProceed = True
            End If
        End If
    End If
    If bProceed Then
        Call RegisterRenter
    Else
        MsgBox " You are missing one of the basic inputs (Drivers License,LastName,Name,Brand,Model Year,COnfirmation #,ETC)"
        GoTo ErrorHandler
    End If

End With
    
ErrorHandler:

    
End Sub
