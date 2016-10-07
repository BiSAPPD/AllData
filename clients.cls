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


Public Sub FillFromSheet(ByRef wks As Worksheet, ByVal cStatYear As Integer, ByVal cStatMonth As Integer, ByVal cBrand As String)
    Const cFirstRow = 4
  
    Dim f_i As Long, obj As clsClientInfo, n as Long

   
    Dim clv As clsClientInfoValue
    Set clv = New clsClientInfoValue
    
    
    For f_i = cFirstRow To myLib.GetLastRow
    
        With clv
            .vTypeBusiness = myLib.GetTypeBusiness(cBrand)
            .vDatabaseClientNum = wks.Cells(f_i, 1)
            .vDatabaseClientAndBrandNum = cBrand & .vDatabaseClientNum
            .vUniverseCode = myLib.getUniversCode(cBrand, f_i, wks.Cells(f_i, 2))
            .vMregName = wks.Cells(f_i, 4)
            .vRegName = wks.Cells(f_i, 5)
            .vExtMregName = myLib.GetMregExt(.vMregName, .vRegName)
            .vFlsmName = wks.Cells(f_i, 165)
            .vSecName = wks.Cells(f_i, 6)
            .vSrepName = wks.Cells(f_i, 7)
            .vClient = wks.Cells(f_i, 9)
            .vClientAddres = wks.Cells(f_i, 12)
            .vGeoCity = wks.Cells(f_i, 11)
            .vClientName = myLib.GetClientName(.vClient, .vClientAddres, .vGeoCity)
            .vChainName = wks.Cells(f_i, 19)
            .vChainCode = myLib.GetNum2numNull(Cells(f_i, 20))
            .vGeoReg = wks.Cells(f_i, 13)
            .vClientType = wks.Cells(f_i, 18)
            .vClientTypeRus = myLib.GetClntType(.vClientType, 1)
            .vClientTypeEng = myLib.GetClntType(.vClientType, 2)
            .vClientTypeEngChort = myLib.GetClntType(.vClientType, 3)
            .vClientTypeEngChain = myLib.GetClntType(.vClientType, 4)
            .vCnqYearDate = myLib.GetYearType(cStatYear, myLib.GetNum2num0(wks.Cells(f_i, 65)), 1)
            .vCnqYearGA = myLib.GetYearType(cStatYear, myLib.GetNum2num0(wks.Cells(f_i, 65)), 3)
            .vCnqMonthNameRus = wks.Cells(f_i, 64)
            .vCnqMonthNum = myLib.GetMonthNumeric(.vCnqMonthNameRus)
            .vCnqMonthNameEng = myLib.GetNameMonthEN(.vCnqMonthNum)
            .vChainNameTop = myLib.getNmChainTop(.vChainName, .vChainCode, .vCnqMonthNameEng) 
            .vCnqFullDate = DateSerial(.vCnqYearDate, .vCnqMonthNum, 1)
            .vMagMinPrice = myLib.GetNum2num0(wks.Cells(f_i, 23))
            .vMagMaxPrice = myLib.GetNum2num0(wks.Cells(f_i, 25))
            .vHairDressersPlaces = myLib.GetNum2num0(wks.Cells(f_i, 27))
            .vMagTypePrice = myLib.GetMag(.vMagMinPrice, .vMagMaxPrice, .vHairDressersPlaces, .vTypeBusiness)
            .vMagTypePlace = myLib.GetMag(.vMagMinPrice, .vMagMaxPrice, .vHairDressersPlaces, "place")
            .vMag = .vMagTypePrice & .vMagTypePlace
            .vMagType = IIF(Len(.vMag) = 2, .vMag, Empty)
            .vWorkStatusNum = wks.Cells(f_i, 8)
            .vWorkStatusName = myLib.GetTypeDN(.vWorkStatusNum)
            .vClientEcadCode = wks.Cells(f_i, 29)
            .vHairDressers = myLib.GetNum2num0(wks.Cells(f_i, 28))
            .vPartnerCode = wks.Cells(f_i, 173)
            .vPartnerName = wks.Cells(f_i, 167)
            .vClubStatus = wks.Cells(f_i, 40)
            .vEmotionStatus = Empty
            .vClientTypeGA = myLib.GetYearType(cStatYear, myLib.GetNum2num0(wks.Cells(f_i, 65)), 2)
            .vLtmAvgCaVal = myLib.GetLTM(wks, f_i, cStatMonth, "avg_ca")
            .vLtmAvgCaName = myLib.GetLTM(wks, f_i, cStatMonth, "type_avg_ca") 
            .vLtmFrqOrders = myLib.GetLTM(wks, f_i, cStatMonth, "frqOrders")
            .vClientEvVal = myLib.GetNum2num0(wks.Cells(f_i, 92)) 
            .vClientEvName = myLib.getVectoreEV(.vClientEvVal)
            .vMastersEducatedAllY = myLib.GetNum2numNull(wks.Cells(f_i, 30))
            .vMastersEducatedPY = myLib.GetNum2numNull(wks.Cells(f_i, 31))
            .vMastersEducatedTY = myLib.GetNum2numNull(wks.Cells(f_i, 32))
        End With

        Dim clc As clsClientCAValue
        Set clc = New clsClientCAValue
        Const clmTYLor = 93, clmPYLor = 106     

        n = 0
         With clc
            n = n + 1
            .CA_TY_M1 = IIf(cStatMonth >= n, myLib.GetNum2numNull(wks.Cells(f_i, clmTYLor + n - 1)), Empty)
            .CA_PY_M1 = IIf(cStatMonth >= 1, myLib.GetNum2numNull(wks.Cells(f_i, clmPYLor + n - 1)), Empty)
            .CA_TY_YTD1 = IIf(cStatMonth >= n, .CA_TY_M1, Empty)
            .CA_PY_YTD1 = IIf(cStatMonth >= 1, .CA_PY_M1, Empty)

            n = n + 1
            .CA_TY_M2 = IIf(cStatMonth >= n, myLib.GetNum2numNull(wks.Cells(f_i, clmTYLor + n - 1)), Empty)
            .CA_PY_M2 = IIf(cStatMonth >= 1, myLib.GetNum2numNull(wks.Cells(f_i, clmPYLor + n - 1)), Empty)
            .CA_TY_YTD2 = IIf(cStatMonth >= n, .CA_TY_YTD1 + .CA_TY_M2, Empty)
            .CA_PY_YTD2 = IIf(cStatMonth >= 1, .CA_PY_YTD1 + .CA_PY_M2, Empty)

            n = n + 1
            .CA_TY_M3 = IIf(cStatMonth >= n, myLib.GetNum2numNull(wks.Cells(f_i, clmTYLor + n - 1)), Empty)
            .CA_PY_M3 = IIf(cStatMonth >= 1, myLib.GetNum2numNull(wks.Cells(f_i, clmPYLor + n - 1)), Empty)
            .CA_TY_YTD3 = IIf(cStatMonth >= n, .CA_TY_YTD2 + .CA_TY_M3, Empty)
            .CA_PY_YTD3 = IIf(cStatMonth >= 1, .CA_PY_YTD2 + .CA_PY_M3, Empty)
            .CA_TY_Q1 = .CA_TY_YTD3
            .CA_PY_Q1 = .CA_PY_YTD3

            n = n + 1
            .CA_TY_M4 = IIf(cStatMonth >= n, myLib.GetNum2numNull(wks.Cells(f_i, clmTYLor + n - 1)), Empty)
            .CA_PY_M4 = IIf(cStatMonth >= 1, myLib.GetNum2numNull(wks.Cells(f_i, clmPYLor + n - 1)), Empty)
            .CA_TY_YTD4 = IIf(cStatMonth >= n, .CA_TY_YTD3 + .CA_TY_M4, Empty)
            .CA_PY_YTD4 = IIf(cStatMonth >= 1, .CA_PY_YTD3 + .CA_PY_M4, Empty)

            n = n + 1
            .CA_TY_M5 = IIf(cStatMonth >= n, myLib.GetNum2numNull(wks.Cells(f_i, clmTYLor + n - 1)), Empty)
            .CA_PY_M5 = IIf(cStatMonth >= 1, myLib.GetNum2numNull(wks.Cells(f_i, clmPYLor + n - 1)), Empty)
            .CA_TY_YTD5 = IIf(cStatMonth >= n, .CA_TY_YTD4 + .CA_TY_M5, Empty)
            .CA_PY_YTD5 = IIf(cStatMonth >= 1, .CA_PY_YTD4 + .CA_PY_M5, Empty)

            n = n + 1
            .CA_TY_M6 = IIf(cStatMonth >= n, myLib.GetNum2numNull(wks.Cells(f_i, clmTYLor + n - 1)), Empty)
            .CA_PY_M6 = IIf(cStatMonth >= 1, myLib.GetNum2numNull(wks.Cells(f_i, clmPYLor + n - 1)), Empty)
            .CA_TY_YTD6 = IIf(cStatMonth >= n, .CA_TY_YTD5 + .CA_TY_M6, Empty)
            .CA_PY_YTD6 = IIf(cStatMonth >= 1, .CA_PY_YTD5 + .CA_PY_M6, Empty)
            .CA_TY_Q2 = .CA_TY_YTD6
            .CA_PY_Q2 = .CA_PY_YTD6

            n = n + 1
            .CA_TY_M7 = IIf(cStatMonth >= n, myLib.GetNum2numNull(wks.Cells(f_i, clmTYLor + n - 1)), Empty)
            .CA_PY_M7 = IIf(cStatMonth >= 1, myLib.GetNum2numNull(wks.Cells(f_i, clmPYLor + n - 1)), Empty)
            .CA_TY_YTD7 = IIf(cStatMonth >= n, .CA_TY_YTD6 + .CA_TY_M7, Empty)
            .CA_PY_YTD7 = IIf(cStatMonth >= 1, .CA_PY_YTD6 + .CA_PY_M7, Empty)

            n = n + 1
            .CA_TY_M8 = IIf(cStatMonth >= n, myLib.GetNum2numNull(wks.Cells(f_i, clmTYLor + n - 1)), Empty)
            .CA_PY_M8 = IIf(cStatMonth >= 1, myLib.GetNum2numNull(wks.Cells(f_i, clmPYLor + n - 1)), Empty)
            .CA_TY_YTD8 = IIf(cStatMonth >= n, .CA_TY_YTD7 + .CA_TY_M8, Empty)
            .CA_PY_YTD8 = IIf(cStatMonth >= 1, .CA_PY_YTD7 + .CA_PY_M8, Empty)

            n = n + 1
            .CA_TY_M9 = IIf(cStatMonth >= n, myLib.GetNum2numNull(wks.Cells(f_i, clmTYLor + n - 1)), Empty)
            .CA_PY_M9 = IIf(cStatMonth >= 1, myLib.GetNum2numNull(wks.Cells(f_i, clmPYLor + n - 1)), Empty)
            .CA_TY_YTD9 = IIf(cStatMonth >= n, .CA_TY_YTD8 + .CA_TY_M9, Empty)
            .CA_PY_YTD9 = IIf(cStatMonth >= 1, .CA_PY_YTD8 + .CA_PY_M9, Empty)
            .CA_TY_Q3 = .CA_TY_YTD9
            .CA_PY_Q3 = .CA_PY_YTD9

            n = n + 1
            .CA_TY_M10 = IIf(cStatMonth >= n, myLib.GetNum2numNull(wks.Cells(f_i, clmTYLor + n - 1)), Empty)
            .CA_PY_M10 = IIf(cStatMonth >= 1, myLib.GetNum2numNull(wks.Cells(f_i, clmPYLor + n - 1)), Empty)
            .CA_TY_YTD10 = IIf(cStatMonth >= n, .CA_TY_YTD9 + .CA_TY_M10, Empty)
            .CA_PY_YTD10 = IIf(cStatMonth >= 1, .CA_PY_YTD9 + .CA_PY_M10, Empty)

            n = n + 1
            .CA_TY_M11 = IIf(cStatMonth >= n, myLib.GetNum2numNull(wks.Cells(f_i, clmTYLor + n - 1)), Empty)
            .CA_PY_M11 = IIf(cStatMonth >= 1, myLib.GetNum2numNull(wks.Cells(f_i, clmPYLor + n - 1)), Empty)
            .CA_TY_YTD11 = IIf(cStatMonth >= n, .CA_TY_YTD10 + .CA_TY_M11, Empty)
            .CA_PY_YTD11 = IIf(cStatMonth >= 1, .CA_PY_YTD10 + .CA_PY_M11, Empty)

            n = n + 1
            .CA_TY_M12 = IIf(cStatMonth >= n, myLib.GetNum2numNull(wks.Cells(f_i, clmTYLor + n - 1)), Empty)
            .CA_PY_M12 = IIf(cStatMonth >= 1, myLib.GetNum2numNull(wks.Cells(f_i, clmPYLor + n - 1)), Empty)
            .CA_TY_YTD12 = IIf(cStatMonth >= n, .CA_TY_YTD11 + .CA_TY_M12, Empty)
            .CA_PY_YTD12 = IIf(cStatMonth >= 1, .CA_PY_YTD11 + .CA_PY_M12, Empty)
            .CA_TY_Q4 = .CA_TY_YTD12
            .CA_PY_Q4 = .CA_PY_YTD12
        End With


        Set obj = New clsClientInfo
        With obj
            .StatYear = cStatYear
            .StatMonth = cStatMonth
            .BrandName = cBrand
            .TypeBusiness = clv.vTypeBusiness
            .DatabaseClientNum = clv.vDatabaseClientNum
            .DatabaseClientAndBrandNum = clv.vDatabaseClientAndBrandNum
            .UniverseCode = clv.vUniverseCode
            .UniversCodeAndBrand = cBrand & clv.vUniverseCode
            .MregName = clv.vMregName
            .ExtMregName = clv.vExtMregName
            .RegName = clv.vRegName
            .FlsmName = clv.vFlsmName
            .SecName = clv.vSecName
            .SrepName = clv.vSrepName
            .ClientName = clv.vClientName
            .ChainName = clv.vChainName
            .ChainName = clv.vChainName
            .ChainCode = clv.vChainCode
            .ChainNameTop = clv.vChainNameTop
            .GeoCity = clv.vGeoCity
            .GeoReg = clv.vGeoReg
            .ClientTypeRus = clv.vClientTypeRus
            .ClientTypeEng = clv.vClientTypeEng
            .ClientTypeEngChort = clv.vClientTypeEngChort
            .ClientTypeEngChain = clv.vClientTypeEngChain
            .ClientTypeGA = clv.vClientTypeGA
            .ClubStatus = clv.vClubStatus
            .EmotionStatus = clv.vEmotionStatus
            .CnqFullDate = clv.vCnqFullDate
            .CnqYearDate = clv.vCnqYearDate
            .CnqYearGA = clv.vCnqYearGA
            .CnqMonthNum = clv.vCnqMonthNum
            .CnqMonthNameRus = clv.vCnqMonthNameRus
            .CnqMonthNameEng = clv.vCnqMonthNameEng
            .MagType = clv.vMagType
            .MagTypePrice = clv.vMagTypePrice
            .MagTypePlace = clv.vMagTypePlace
            .WorkStatusNum = clv.vWorkStatusNum
            .WorkStatusName = clv.vWorkStatusName
            .LtmAvgCaVal = clv.vLtmAvgCaVal
            .LtmAvgCaName = clv.vLtmAvgCaName
            .LtmFrqOrders = clv.vLtmFrqOrders
            .ClientEvVal = clv.vClientEvVal
            .ClientEvName = clv.vClientEvName
            .ClientEcadCode = clv.vClientEcadCode
            .MastersEducatedAllY = clv.vMastersEducatedAllY
            .MastersEducatedPY = clv.vMastersEducatedPY
            .MastersEducatedTY = clv.vMastersEducatedTY
            .HairdressersNum = clv.vHairDressers
            .HairdressersWorkPlace = clv.vHairDressersPlaces
            .PartnerName = clv.vPartnerCode
            .PartnerCode = clv.vPartnerName

            .CA_TY_M1 = clc.CA_TY_M1
            .CA_TY_M2 = clc.CA_TY_M2
            .CA_TY_M3 = clc.CA_TY_M3
            .CA_TY_M4 = clc.CA_TY_M4
            .CA_TY_M5 = clc.CA_TY_M5
            .CA_TY_M6 = clc.CA_TY_M6
            .CA_TY_M7 = clc.CA_TY_M7
            .CA_TY_M8 = clc.CA_TY_M8
            .CA_TY_M9 = clc.CA_TY_M9
            .CA_TY_M10 = clc.CA_TY_M10
            .CA_TY_M11 = clc.CA_TY_M11
            .CA_TY_M12 = clc.CA_TY_M12

            .CA_PY_M1 = clc.CA_PY_M1
            .CA_PY_M2 = clc.CA_PY_M2
            .CA_PY_M3 = clc.CA_PY_M3
            .CA_PY_M4 = clc.CA_PY_M4
            .CA_PY_M5 = clc.CA_PY_M5
            .CA_PY_M6 = clc.CA_PY_M6
            .CA_PY_M7 = clc.CA_PY_M7
            .CA_PY_M8 = clc.CA_PY_M8
            .CA_PY_M9 = clc.CA_PY_M9
            .CA_PY_M10 = clc.CA_PY_M10
            .CA_PY_M11 = clc.CA_PY_M11
            .CA_PY_M12 = clc.CA_PY_M12

            .CA_TY_YTD1 = clc.CA_TY_YTD1
            .CA_TY_YTD2 = clc.CA_TY_YTD2
            .CA_TY_YTD3 = clc.CA_TY_YTD3
            .CA_TY_YTD4 = clc.CA_TY_YTD4
            .CA_TY_YTD5 = clc.CA_TY_YTD5
            .CA_TY_YTD6 = clc.CA_TY_YTD6
            .CA_TY_YTD7 = clc.CA_TY_YTD7
            .CA_TY_YTD8 = clc.CA_TY_YTD8
            .CA_TY_YTD9 = clc.CA_TY_YTD9
            .CA_TY_YTD10 = clc.CA_TY_YTD10
            .CA_TY_YTD11 = clc.CA_TY_YTD11
            .CA_TY_YTD12 = clc.CA_TY_YTD12

            .CA_PY_YTD1 = clc.CA_PY_YTD1
            .CA_PY_YTD2 = clc.CA_PY_YTD2
            .CA_PY_YTD3 = clc.CA_PY_YTD3
            .CA_PY_YTD4 = clc.CA_PY_YTD4
            .CA_PY_YTD5 = clc.CA_PY_YTD5
            .CA_PY_YTD6 = clc.CA_PY_YTD6
            .CA_PY_YTD7 = clc.CA_PY_YTD7
            .CA_PY_YTD8 = clc.CA_PY_YTD8
            .CA_PY_YTD9 = clc.CA_PY_YTD9
            .CA_PY_YTD10 = clc.CA_PY_YTD10
            .CA_PY_YTD11 = clc.CA_PY_YTD11
            .CA_PY_YTD12 = clc.CA_PY_YTD12

            .CA_TY_Q1 = clc.CA_TY_Q1
            .CA_TY_Q2 = clc.CA_TY_Q2
            .CA_TY_Q3 = clc.CA_TY_Q3
            .CA_TY_Q4 = clc.CA_TY_Q4

            .CA_PY_Q1 = clc.CA_PY_Q1
            .CA_PY_Q2 = clc.CA_PY_Q2
            .CA_PY_Q3 = clc.CA_PY_Q3
            .CA_PY_Q4 = clc.CA_PY_Q4

        End With
        
        Me.Add obj
    Next
End Sub



