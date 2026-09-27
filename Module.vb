Public Sub RegisterRenter()
Dim Car As CarClass

    Set Car = New CarClass

    Car.AddDriversLicense = fmCarRental.tbDL
    Car.AddRenterName = fmCarRental.tbName.Value
    Car.AddRenterLastName = fmCarRental.tbLName.Value
    Car.AddVehicleModel = fmCarRental.lbVehicleModels.Value
    Car.addVehicleYear = fmCarRental.tbMY.Value
    Car.addConfirmationCode = fmCarRental.tbConfirmation.Value
    Car.addVehicleBrand = fmCarRental.tbVBrand.Value
    
    
    
    Debug.Print (Car.GetRenterName)
    Debug.Print (Car.GetLastName)
    Debug.Print (Car.GetDriversLicense)
    Debug.Print (Car.getConfirmCode)
    Debug.Print (Car.GetVehicleYear)
    Debug.Print (Car.GetVehicleModel)
    Debug.Print (Car.getVehicleBrand)
    
    

    
End Sub
Public Function generateAccessConnection(ByVal accessDBPath As String, ByVal accessDBFileName As String) As Boolean
Dim accConnection As Object, accRecordset As Object, sqlString As String


Set accConnection = CreateObject("ADODB.Connection")

Do While Err.Number <> 0
    Err.Clear
    Set accConnection = CreateObject("ADODB.Connection")
    
Loop

Set accRecordset = CreateObject("ADODB.Recordset")


Do While Err.Number <> 0
    Err.Clear
    Set accRecordset = CreateObject("ADODB.Recordset")
Loop


accConnection.Open "Provider:=Microsoft.ACE.OLEDB.12.0;Data Source:= accessDBPath & accessDBFileName;Persist Security Info=False;"



End Function
Public Sub testIt()

bResult = CreateAccessDatabase

End Sub
Public Function VerifyFileExistence(ByVal sPath As String, ByVal sFileFormat As String, Optional sActualFileName As String) As Boolean
Dim xFileName As String, bFound As Boolean

    Let xFileName = Dir(sPath)
    bFound = False
    
    Do While xFileName <> ""
        If InStr(1, UCase(xFileName), UCase(sFileFormat), vbBinaryCompare) >= 1 Then
            If Len(sActualFileName) > 0 Then
                If InStr(1, UCase(xFileName), UCase(sActualFileName), vbBinaryCompare) >= 1 Then
                    bFound = True
                End If
            Else
                bFound = True
            End If
        End If
        xFileName = Dir
    Loop
    
    VerifyFileExistence = bFound

End Function
Public Sub setPause(ByVal xSeconds As Byte)
Dim xCurrentTime As Variant, bStillWaiting As Boolean, runningTime As Variant


    Let xCurrentTime = WorksheetFunction.Text(Now, "ss") + xSeconds
    If xCurrentTime > 60 Then
        xCurrentTime = 60 - xCurrentTime + 1
    End If
    bStillWaiting = True


    Do While bStillWaiting
        runningTime = WorksheetFunction.Text(Now, "ss") + 1 - 1
        If runningTime >= xCurrentTime Then
            bStillWaiting = False
        End If
    Loop




End Sub
Public Sub deleteAllCoincidences(ByVal sPath As String, ByVal sFileName As String)
Dim xFileName As String


xFileName = Dir(sPath)
bDeletedFiles = 0

Do While xFileName <> ""
    If InStr(1, UCase(xFileName), UCase(sFileName), vbBinaryCompare) >= 1 Then
        Kill sPath & xFileName
        bDeletedFiles = bDeletedFiles + 1
    End If
    xFileName = Dir
Loop

    If bDeletedFiles > 0 Then
        Debug.Print (bDeletedFiles & " were deleted from the directory  : " & vbNewLine & sPath)
    End If
    


End Sub
Public Function CreateAccessDatabase() As Boolean
Dim fsObject As Object, oShell As Object, accApp As Object, sqlString As String, accCatalog As Object, dbName As String
Dim accConnection As Object


    
    Set oShell = CreateObject("Shell.Application")
    Let accessDBPath = "C:\Users\" & getNetworkUsername & "\VBA Apps\CRUD App\"
    Let dbName = "RockwareIMM1"
    Let serverSource = "JACOBODUEZ\DEMO"
    Debug.Print (accessDBPath)
    Let connectionString = "Provider=MSOLEDBSQL;Data Source=" & serverSource & ";Initial Catalog=" & dbName & ";Integrated Security=SSPI;"
    

    
    'Call createFullPath(accessDBPath)
    'Call deleteAllCoincidences(accessDBPath, dbName)
    'Shell """C:\Program Files\Microsoft Office\root\Office16\MSACCESS.EXE""", vbNormalFocus
    'Implement a pause procedure here
    Call setPause(1)
    
'
'    Set accApp = GetObject(, "Access.Application")
'
    Set accConnection = CreateObject("ADODB.Connection")
    
    Debug.Print (connectionString)
    accConnection.Open connectionString

    
    If Not VerifyFileExistence(accessDBPath, ".ACCDB", dbName) Then
        MsgBox "The Database could not be generated!", vbCritical, "Function Error at CreateAccessDatabase function"
        GoTo finalSection
    End If
    
    Set accCatalog = Nothing
    
    'accConnection.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & accessDBPath & dbName

  
    
  

    sqlString = "IF NOT EXISTS(SELECT * FROM [RockwareIMM1].dbo.tblVehicleRental)" & _
    "BEGIN " & _
    "CREATE TABLE [RockwareIMM1].dbo.tblVehicleRental ( RecordID  BIGINT PRIMARY KEY NOT NULL," & _
    "VehicleModel NVARCHAR (50)," & _
    "VehicleYear NVARCHAR(20)," & _
    "VehicleBrand NVARCHAR(50)," & _
    "VehicleSegment NVARCHAR(50)," & _
    "UpperMileageLimit NVARCHAR (50)," & _
    "LicensePlate NVARCHAR(50)," & _
    "MilageOnPurchase NVARCHAR (50)," & _
    "TotalMileageUsage NVARCHAR(50))" & _
    "END"
        
    Debug.Print (sqlString)
    
    accConnection.Execute sqlString
    

    sqlString = "INSERT INTO tblVehicleRental (VehicleModel,VehicleYear,UpperMileageLimit,LicensePlate) VALUES ('Porsche 911','2023','5000','CHARLIE9')"
    
    accConnection.Execute sqlString
    sqlString2 = "INSERT INTO tblVehicleRental (VehicleModel,VehicleYear,UpperMileageLimit,LicensePlate) VALUES ('4Runner','2025','15000','OREGON5')"
    accConnection.Execute sqlString2
    

   ' sqlString = sqlString & sqlString2
    
    'accConnection.Execute sqlString
 

    
finalSection:
    
    
    

End Function
Public Sub testDeletion()
Dim accessDBPath As String, dbName As String

Let accessDBPath = "C:\Users\" & getNetworkUsername & "\VBA Apps\CRUD App\"
Let dbName = "Testing.accdb"

If Not DropTable(accessDBPath, "tblVehicleRental", dbName) Then

Else
    
End If
    

End Sub
Public Function DropTable(ByVal dbPath As String, ByVal tbLName As String, ByVal dbName As String) As Boolean
Dim accApp As Object, accConnection As Object, sqlString As String, bStatus As Boolean
    
    
    bStatus = False
    
    Set accApp = GetObject(, "Access.Application")
    Set accConnection = CreateObject("ADODB.Connection")

    accConnection.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & dbName

    sqlString = "DROP Table " & tbLName
    accConnection.Execute sqlString
    
    Set accConnection = Nothing
    
    DropTable = bStatus


End Function
Public Function DeleteRecords(ByVal dbPath As String, ByVal tbLName As String) As Boolean
Dim bStatus As Boolean
    
    bStatus = False
    
    Set accApp = GetObject(, "Access.Application")
    Set accConnection = CreateObject("ADODB.Connection")
    

    

    DeleteRecords = bStatus
    
End Function
Public Function performRetrieve(ByVal dbPath As String, ByVal tbLName As String, ByVal dbName As String, Optional tempSh As Worksheet) As Boolean






End Function
Public Function CreateTable(ByVal dbPath As String, ByVal tbLName As String, ByVal dbName As String) As Boolean
Dim bCreatedSuccessfully As Boolean, sActualName As String, accConnection As Object, accRecordset As Object
Dim nTries As Byte

Let bCreatedSuccessfully = False


Select Case UCase(tbLName)
    Case Is = "tblVehicleRental"
        sActualTblName = "tblVehicleRental"
        sqlString = "CREATE TABLE tblVehicleRental ( RecordID  AUTOINCREMENT PRIMARY KEY," & _
            "VehicleModel TEXT (50)," & _
            "VehicleYear TEXT(20)," & _
            "VehicleBrand TEXT(50)," & _
            "VehicleSegment TEXT(50)," & _
            "UpperMileageLimit TEXT (50)," & _
            "LicensePlate TEXT (50)," & _
            "MilageOnPurchase TEXT (50)," & _
            "TotalMileageUsage TEXT(50))"
    Case Is = "tblVehicleHistoric"
        sActualTblName = "tblVehicleHistoric"
        sqlString = "CREATE TABLE tblVehicleHistoric (RecordID AUTOINCREMENT PRIMARY KEY NOT NULL," & _
        "vehicleModelYear Date," & _
        "vehicleMileageAtPurchase BigInt," & _
        "vehicle"
    Case Is = "tblRentersHistoric"
        sActualTblName = "tblVehicleHistoric"
    Case Is = "tblAccountsReceivable"
        sActualTblName = "tblAccountsReceivable"
    Case Is = "tblReservations"
        sActualTblName = "tblReservations"
    Case Is = "tblProfitandLoss"
        sActualTblName = "tblProfitandLoss"
End Select


    Set accConnection = CreateObject("ADODB.Connection")
    
    Do While Err.Number <> 0
           Err.Clear
           Set accConnection = CreateObject("ADODB.Application")
           nTries = nTries + 1
           If nTries > 3 Then
                Err.Clear
           End If
    Loop
    
    Set accRecordset = CreateObject("ADODB.Recordset")
    
    nTries = 0
    Do While Err.Number <> 0
           Err.Clear
           Set accRecordset = CreateObject("ADODB.Recordset")
           nTries = nTries + 1
           If nTries > 3 Then
                Err.Clear
           End If
    Loop
    
     
    accConnection.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & dbPath & dbName




End Function
Public Function ListTableNames(ByVal dbPath As String, ByVal dbName As String) As Variant
Dim accessRS As Object, accessTbls As Object, accessConnection As Object
Dim tblArr(1 To 1000) As Variant

Set accessConnection = CreateObject("ADODB.Connection")

    bExitForGood = False
    If Err.Number <> 0 Then
        bRetryConnection = True
        Do While bRetryConnection
            iTryCount = iTryCount + 1
            Err.Clear
            Set accessConnection = CreateObject("ADODB.Connection")
            If Err.Number = 0 Then
                bRetryConnection = False
            End If
            If iTryCount = 2 Then
                bRetryConnection = False
                bExitForGood = True
            End If
        Loop
    End If

    If bExitForGood Then
        GoTo ErrorHandler
    End If
    
    accessConnection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & dbName
    
    Set accessRS = CreateObject("ADODB.Recordset")
    
    
    tblId = 1
    Do Until accessRS.EOF
    

ErrorHandler:


End Function
Public Function createFullPath(ByVal xPath As String) As Boolean
Dim fsObject As Object, sPossiblePaths As Variant

Set fsObject = CreateObject("Scripting.FileSystemObject")
bPathExists = False

If Not fsObject.FolderExists(xPath) Then
    sPossiblePaths = Split(xPath, "\")
    For t = 0 To UBound(sPossiblePaths)
        If t <> UBound(sPossiblePaths) Then
            sConcatenatedPath = sConcatenatedPath & sPossiblePaths(t) & "\"
        Else
            sConcatenatedPath = sConcatenatedPath & sPossiblePaths(t) & "\"
        End If
        If Not fsObject.FolderExists(sConcatenatedPath) Then
            MkDir sConcatenatedPath
        End If
    Next t
End If


If Not fsObject.FolderExists(xPath) Then
    bPathExists = False
Else
    bPathExists = True
End If

createFullPath = bPathExists

End Function
Public Sub clearFolder(ByVal xPath As String)
Dim sFileName As String

sFileName = Dir(xPath)

Do While sFileName <> ""
    Kill xPath & sFileName
    sFileName = Dir
Loop


End Sub
Public Sub displayCarRentalForm()
Dim oControl As MSForms.Control, tSh As Worksheet


For Each oControl In fmCarRental.Controls
    If InStr(1, UCase(Left(oControl.Name, 3)), "LBL", vbBinaryCompare) >= 1 And InStr(1, UCase(TypeName(oControl)), "LABEL", vbBinaryCompare) >= 1 Then
            oControl.BackStyle = fmBackStyleTransparent
    End If
Next oControl

    fmCarRental.lbVehicleModels.AddItem "Porsche 911"
    fmCarRental.lbVehicleModels.AddItem "BMW X5"
    fmCarRental.lbVehicleModels.AddItem "Audi Q7"
    fmCarRental.lbVehicleModels.AddItem "Tesla Model Y"
    fmCarRental.lbVehicleModels.AddItem "Jeep Wrangler"
    fmCarRental.lbVehicleModels.AddItem "Mustang"
    fmCarRental.lbVehicleModels.AddItem "Citröen 5S"
    fmCarRental.lbVehicleModels.AddItem "Vauxhall Spider"
    fmCarRental.lbVehicleModels.AddItem "Kia Sportage"
    
    
    

fmCarRental.Width = 960
fmCarRental.Height = 530


fmCarRental.Show




End Sub
Public Function getNetworkUsername() As String
Dim wScriptObject As Object

Set wScriptObject = CreateObject("wScript.Network")

getNetworkUsername = LCase(wScriptObject.UserName)

End Function
Public Sub enableExcelFeatures()

Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
Application.DisplayStatusBar = True
Application.ScreenUpdating = True

End Sub
Public Sub disableExcelFeatures()

Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
Application.DisplayStatusBar = False
Application.ScreenUpdating = False

End Sub
