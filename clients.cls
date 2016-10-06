VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsClients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private colClients As Collection

Private Sub Class_Initialize()
    Set colClients = New Collection
End Sub

Private Sub Class_Terminate()
    Set colClients = Nothing
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = colClients.[_NewEnum]
End Property

Public Sub Add(obj As clsClientInfo)
    colClients.Add obj
End Sub

Public Sub Remove(Index As Variant)
    colClients.Remove Index
End Sub

Public Property Get Item(Index As Variant) As clsClientInfo
Attribute Item.VB_UserMemId = 0
    Set Item = colClients.Item(Index)
End Property

Property Get Count() As Long
    Count = colClients.Count
End Property

Public Sub Clear()
    Set colClients = New Collection
End Sub


Public Sub FillFromSheet(ByRef wks As Worksheet, ByVal vStatYear As Integer, ByVal vStatMonth As Integer, ByVal vBrand As String)
    Const cFirstRow = 4
  
    Dim f_i As Long, obj As clsClientInfo

    'Dim vStatYear As Integer
    'Dim vStatMonth As Integer
    'Dim vBrand As String
    Dim vBrandName As String
    Dim vTypeBusiness As String
    Dim vDatabaseClientNum As Integer
    Dim vDatabaseClientAndBrandNum As String
    Dim vUniverseCode As Variant
    Dim vUniversCodeAndBrand As String
    Dim vMregName As String
    Dim vExtMregName As String
    Dim vRegName As String
    Dim vFlsmName As String
    Dim vSecName As String
    Dim vSrepName As String
    Dim vClient As String
    Dim vClientAddres As String
    Dim vClientName As String
    Dim vChainName As String
    Dim vChainCode As Double
    Dim vGeoCity As String
    Dim vGeoReg As String
    Dim vClientType As String
    Dim vClientTypeRus As String
    Dim vClientTypeEng As String
    Dim vClientTypeEngChort As String
    Dim vClientTypeEngChain As String
    Dim vClubStatus As String
    Dim vEmotionStatus As String
    Dim vCnqFullDate As Date
    Dim vCnqYearDate As Date
    Dim vCnqYearGA As String
    Dim vCnqMonthNum As Integer
    Dim vCnqMonthNameRus As String
    Dim vCnqMonthNameEng As String
    Dim vMagMinPrice As Long
    Dim vMagMaxPrice As Long
    Dim vHairDressersPlaces As Long
    Dim vMagType As String
    Dim vMagTypePrice As String
    Dim vMagTypePlace As String
    Dim vMag As String
    Dim vWorkStatusNum As Integer
    Dim vWorkStatusName As String
    Dim vHairDressers As Integer
    Dim vLtmAvgCaVal As Double
    Dim vLtmAvgCaName As String
    Dim vLtmFrqOrders As String
    Dim vClientEvVal As Double
    Dim vClientEvName As String
    Dim vClientEcadCode As Variant
    Dim vMastersEducatedAllY As Integer
    Dim vMastersEducatedPY As Integer
    Dim vMastersEducatedTY As Integer
    Dim vHairdressersNum As Integer
    Dim vHairdressersWorkPlace As Integer
    Dim vPartnerName As String
    Dim vPartnerCode As Variant

    
    For f_i = cFirstRow To myLib.GetLastRow
    
        With ActiveSheet
            vTypeBusiness = myLib.GetTypeBusiness(vBrand)
            vDatabaseClientNum = .Cells(f_i, 1)
            vDatabaseClientAndBrandNum = vBrand & vDatabaseClientNum
            vUniverseCode = myLib.getUniversCode(vBrand, f_i, .Cells(f_i, 2))
            vMregName = .Cells(f_i, 4)
            vRegName = .Cells(f_i, 5)
            vExtMregName = myLib.GetMregExt(vMregName, vRegName)
            vFlsmName = .Cells(f_i, 165)
            vSecName = .Cells(f_i, 6)
            vSrepName = .Cells(f_i, 7)
            vClient = .Cells(f_i, 9)
            vClientAddres = .Cells(f_i, 12)
            vGeoCity = .Cells(f_i, 11)
            vClientName = myLib.GetClientName(vClient, vClientAddres, vGeoCity)
            vChainName = .Cells(f_i, 19)
            vChainCode = myLib.num2numNull(Cells(f_i, 20))
            vGeoReg = .Cells(f_i, 13)
            vClientType = .Cells(f_i, 18)
            vClientTypeRus = myLib.GetClntType(vClientType, 1)
            vClientTypeEng = myLib.GetClntType(vClientType, 2)
            vClientTypeEngChort = myLib.GetClntType(vClientType, 3)
            vClientTypeEngChain = myLib.GetClntType(vClientType, 4)
            vCnqYearDate = myLib.GetYearType(vStatYear, myLib.GetNum2num0(.Cells(f_i, 65)), 1)
            vCnqYearGA = myLib.GetYearType(vStatYear, myLib.GetNum2num0(.Cells(f_i, 65)), 3)
            vCnqMonthNameRus = .Cells(f_i, 64)
            vCnqMonthNum = myLib.GetMonthNumeric(vCnqMonthNameRus)
            vCnqMonthNameEng = myLib.GetNameMonthEN(vCnqMonthNum)
            vCnqFullDate = DateSerial(vCnqYearDate, vCnqMonthNum, 1)
            vMagMinPrice = myLib.GetNum2num0(.Cells(f_i, 23))
            vMagMaxPrice = myLib.GetNum2num0(.Cells(f_i, 25))
            vHairDressersPlaces = myLib.GetNum2num0(.Cells(f_i, 27))
            vMagTypePrice = myLib.GetMag(vMagMinPrice, vMagMaxPrice, vHairDressersPlaces, vTypeBusiness)
            vMagTypePlace = myLib.GetMag(vMagMinPrice, vMagMaxPrice, vHairDressersPlaces, "place")
            vMag = vMagTypePrice & vMagTypePlace
            vMagType = IIF(Len(vMag) = 2, vMag, Empty)
            vWorkStatusNum = .Cells(f_i, 8)
            vWorkStatusName = myLib.GetTypeDN(vWorkStatusNum)
            vClientEcadCode = .Cells(f_i, 29)
            vHairDressers = myLib.GetNum2num0(.Cells(f_i, 28))
            vPartnerCode = .Cells(f_i, 173)
            vPartnerName = .Cells(f_i, 167)
        End With

        Set obj = New clsClientInfo
        With obj
            .StatYear = vStatYear
            .StatMonth = vStatMonth
            .BrandName = vBrand
            .TypeBusiness = vTypeBusiness
            .DatabaseClientNum = vDatabaseClientNum
            .DatabaseClientAndBrandNum = vDatabaseClientAndBrandNum
            .UniverseCode = vUniverseCode
            .UniversCodeAndBrand = vBrand & vUniverseCode
            .MregName = vMregName
            .ExtMregName = vExtMregName
            .RegName = vRegName
            .FlsmName = vFlsmName
            .SecName = vSecName
            .SrepName = vSrepName
            .ClientName = vClientName
            .ChainName = vChainName
            .ChainCode = vChainCode
            .GeoCity = vGeoCity
            .GeoReg = vGeoReg
            .ClientTypeRus = vClientTypeRus
            .ClientTypeEng = vClientTypeEng
            .ClientTypeEngChort = vClientTypeEngChort
            .ClientTypeEngChain = vClientTypeEngChain
            .ClubStatus = Empty
            .EmotionStatus = Empty
            .CnqFullDate = vCnqFullDate
            .CnqYearDate = vCnqYearDate
            .CnqYearGA = vCnqYearGA
            .CnqMonthNum = vCnqMonthNum
            .CnqMonthNameRus = vCnqMonthNameRus
            .CnqMonthNameEng = vCnqMonthNameEng
            .MagType = vMagType
            .MagTypePrice = vMagTypePrice
            .MagTypePlace = vMagTypePlace
            .WorkStatusNum = vWorkStatusNum
            .WorkStatusName = vWorkStatusName
            .LtmAvgCaVal = Empty
            .LtmAvgCaName = Empty
            .LtmFrqOrders = Empty
            .ClientEvVal = Empty
            .ClientEvName = Empty
            .ClientEcadCode = vClientEcadCode
            .MastersEducatedAllY = Empty
            .MastersEducatedPY = Empty
            .MastersEducatedTY = Empty
            .HairdressersNum = vHairDressers
            .HairdressersWorkPlace = vHairDressersPlaces
            .PartnerName = vPartnerCode
            .PartnerCode = vPartnerName
        End With
        
        Me.Add obj
    Next
End Sub



