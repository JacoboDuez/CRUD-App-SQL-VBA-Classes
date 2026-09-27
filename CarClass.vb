Public lUpperMileLimit As Long
Public sLicensePlate As String
Public currentLocation As String
Public currentMileage As Variant
Public tripCost As Long
Public sRenterName As String ' Implemented GET AND LET
Public sRenterLName As String ' Implemented Get and LET
Public sRenterDriverLicense As String ' Implemented Get and LET
Public sVehicleModel As String ' Implemented GET and LET
Public sVehicleYear As Variant ' Implemented GET AND LET
Public sVehicleBrand As String ' Implemented GET AND LET
Public sConfirmationCode As String ' Implemented GET AND LET
Public bRented As Boolean  '
Public Sub Class_Initialize()

    Debug.Print ("The object is constructed")
    tripCost = 0
    sRenterName = ""
    sRenterLName = ""

End Sub
Public Sub Class_Terminate()
    
    Debug.Print ("THe object is destroyed")


End Sub
Public Property Let AddRenterLastName(ByVal xRenterLastName As Variant)
    sRenterLastName = xRenterLastName
End Property
Public Property Get GetLastName() As String
    GetLastName = sRenterLastName
End Property
Public Property Let AddRenterName(ByVal xRenterName As String)
    sRenterName = xRenterName
End Property
Public Property Get GetRenterName() As String
    GetRenterName = sRenterName
End Property
Public Property Get GetPlate() As String
    GetPlate = sLicensePlate
End Property
Public Property Let LicensePlate(ByVal licPlate As String)
    sLicensePlate = licPlate
End Property
Public Property Let AddMileage(ByVal xMileage As Variant)
    currentMileage = currentMileage + xMileage
End Property
Public Property Get GetMileage() As Variant
    GetMileage = currentMileage
End Property
Public Property Let UpperMileageLimit(ByVal sLimit As Long)
    lUpperMileLimit = sLimit
End Property
Public Property Get GetMileageLimit() As Long
    GetMileageLimit = lUpperMileLimit
End Property
Public Property Let AddDriversLicense(ByVal theLicense As String)
    sDriverLicense = sDriverLicense
End Property
Public Property Get GetDriversLicense() As String
    GetDriversLicense = sLicense
End Property
Public Property Let AddVehicleModel(ByVal sVehicName As String)
    sVehicleModel = sVehicName
End Property
Public Property Get GetVehicleModel() As String
    GetVehicleModel = sVehicleModel
End Property
Public Property Let addVehicleYear(ByVal sVehYear As Variant)
    sVehicleYear = sVehYear
End Property
Public Property Get GetVehicleYear() As Variant
    GetVehicleYear = sVehicleYear
End Property
Public Property Let addVehicleBrand(ByVal sVehBrand As Variant)
    sVehicleBrand = sVehBrand
End Property
Public Property Get getVehicleBrand() As Variant
    getVehicleBrand = sVehicleBrand
End Property
Public Property Let addConfirmationCode(ByVal sDeCode As Variant)
    sConfirmationCode = sDeCode
End Property
Public Property Get getConfirmCode() As Variant
    getConfirmCode = sConfirmationCode
End Property



