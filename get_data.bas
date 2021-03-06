Attribute VB_Name = "get_data_tr"
Sub get_data()
Dim nm_brand As String, patch As String, ShIn As String

Dim ThisYear As Integer, cd_ActualMonth As Integer
Dim LastRow As Long
Dim f_mnth As Integer


myLib.VBA_Start

nm_ActWb = ActiveWorkbook.Name
cd_ActualMonth = CInt(InputBox("Month"))
ThisYear = CInt(InputBox("YearEnd"))

ar_brand = Array("LP", "MX", "KR", "RD", "ES", "DE", "CR")
myLib.VBA_Start

Dim clnts As clsClients, clnt As clsClientInfo
Set clnts = New clsClients

For f_mnth = 1 To cd_ActualMonth
    For f_brnd = 0 To UBound(ar_brand)
        nm_brand = ar_brand(f_brnd)
        ShIn = nm_brand
        ShOut = "TR"
        patch = myLib.GetPatchHistTR(nm_brand, ThisYear, ThisYear, cd_ActualMonth, f_mnth)
        WbTR = myLib.OpenFile(patch, ShIn)
        Workbooks(WbTR).Activate
        Sheets(ShIn).Select
        clnts.FillFromSheet ActiveSheet, ThisYear, f_mnth, nm_brand
        
        Workbooks(WbTR).Close
        Workbooks(nm_ActWb).Activate
    Next f_brnd
Next f_mnth

myLib.CreateSh (ShOut)
myLib.sheetActivateCleer (ShOut)

i = 0
For Each clnt In clnts
    i = i + 1
    n = 0
    With clnt
        n = n + 1: Cells(i + 1, n) = .BrandName:                    If i = 1 Then Cells(1, n) = "BrandName"
        n = n + 1: Cells(i + 1, n) = .DatabaseClientNum:            If i = 1 Then Cells(1, n) = "DatabaseClientNum"
        n = n + 1: Cells(i + 1, n) = .StatYear:                     If i = 1 Then Cells(1, n) = "StatYear"
        n = n + 1: Cells(i + 1, n) = .StatMonth:                    If i = 1 Then Cells(1, n) = "StatMonth"
        n = n + 1: Cells(i + 1, n) = .BrandName:                    If i = 1 Then Cells(1, n) = "BrandName"
        n = n + 1: Cells(i + 1, n) = .TypeBusiness:                 If i = 1 Then Cells(1, n) = "TypeBusiness"
        n = n + 1: Cells(i + 1, n) = .DatabaseClientNum:            If i = 1 Then Cells(1, n) = "DatabaseClientNum"
        n = n + 1: Cells(i + 1, n) = .DatabaseClientAndBrandNum:    If i = 1 Then Cells(1, n) = "DatabaseClientAndBrandNum"
        n = n + 1: Cells(i + 1, n) = .UniverseCode:                 If i = 1 Then Cells(1, n) = "UniverseCode"
        n = n + 1: Cells(i + 1, n) = .UniversCodeAndBrand:          If i = 1 Then Cells(1, n) = "UniversCodeAndBrand"
        n = n + 1: Cells(i + 1, n) = .MregName:                     If i = 1 Then Cells(1, n) = "MregName"
        n = n + 1: Cells(i + 1, n) = .ExtMregName:                  If i = 1 Then Cells(1, n) = "ExtMregName"
        n = n + 1: Cells(i + 1, n) = .RegName:                      If i = 1 Then Cells(1, n) = "RegName"
        n = n + 1: Cells(i + 1, n) = .FlsmName:                     If i = 1 Then Cells(1, n) = "FlsmName"
        n = n + 1: Cells(i + 1, n) = .SecName:                      If i = 1 Then Cells(1, n) = "SecName"
        n = n + 1: Cells(i + 1, n) = .SrepName:                     If i = 1 Then Cells(1, n) = "SrepName"
        n = n + 1: Cells(i + 1, n) = .ClientName:                   If i = 1 Then Cells(1, n) = "ClientName"
        n = n + 1: Cells(i + 1, n) = .ChainName:                    If i = 1 Then Cells(1, n) = "ChainName"
        n = n + 1: Cells(i + 1, n) = .ChainCode:                    If i = 1 Then Cells(1, n) = "ChainCode"
        n = n + 1: Cells(i + 1, n) = .GeoCity:                      If i = 1 Then Cells(1, n) = "GeoCity"
        n = n + 1: Cells(i + 1, n) = .GeoReg:                       If i = 1 Then Cells(1, n) = "GeoReg"
        n = n + 1: Cells(i + 1, n) = .ClientTypeRus:                If i = 1 Then Cells(1, n) = "ClientTypeRus"
        n = n + 1: Cells(i + 1, n) = .ClientTypeEng:                If i = 1 Then Cells(1, n) = "ClientTypeEng"
        n = n + 1: Cells(i + 1, n) = .ClientTypeEngShort:           If i = 1 Then Cells(1, n) = "ClientTypeEngChort"
        n = n + 1: Cells(i + 1, n) = .ClientTypeEngChain:           If i = 1 Then Cells(1, n) = "ClientTypeEngChain"
        n = n + 1: Cells(i + 1, n) = .ClubStatus:                   If i = 1 Then Cells(1, n) = "ClubStatus"
        n = n + 1: Cells(i + 1, n) = .EmotionStatus:                If i = 1 Then Cells(1, n) = "EmotionStatus"
        n = n + 1: Cells(i + 1, n) = .CnqFullDate:                  If i = 1 Then Cells(1, n) = "CnqFullDate"
        n = n + 1: Cells(i + 1, n) = .CnqYearDate:                  If i = 1 Then Cells(1, n) = "CnqYearGA"
        n = n + 1: Cells(i + 1, n) = .CnqYearGA:                    If i = 1 Then Cells(1, n) = "CnqYearGA"
        n = n + 1: Cells(i + 1, n) = .CnqMonthNum:                  If i = 1 Then Cells(1, n) = "CnqMonthNum"
        n = n + 1: Cells(i + 1, n) = .CnqMonthNameRus:              If i = 1 Then Cells(1, n) = "CnqMonthNameRus"
        n = n + 1: Cells(i + 1, n) = .CnqMonthNameEng:              If i = 1 Then Cells(1, n) = "CnqMonthNameEng"
        n = n + 1: Cells(i + 1, n) = .MagType:                      If i = 1 Then Cells(1, n) = "MagType"
        n = n + 1: Cells(i + 1, n) = .MagTypePrice:                 If i = 1 Then Cells(1, n) = "MagTypePrice"
        n = n + 1: Cells(i + 1, n) = .MagTypePlace:                 If i = 1 Then Cells(1, n) = "MagTypePlace"
        n = n + 1: Cells(i + 1, n) = .WorkStatusNum:                If i = 1 Then Cells(1, n) = "WorkStatusNum"
        n = n + 1: Cells(i + 1, n) = .WorkStatusName:               If i = 1 Then Cells(1, n) = "WorkStatusName"
        n = n + 1: Cells(i + 1, n) = .LtmAvgCaVal:                  If i = 1 Then Cells(1, n) = "LtmAvgCaVal"
        n = n + 1: Cells(i + 1, n) = .LtmAvgCaName:                 If i = 1 Then Cells(1, n) = "LtmAvgCaName"
        n = n + 1: Cells(i + 1, n) = .LtmFrqOrders:                 If i = 1 Then Cells(1, n) = "LtmFrqOrders"
        n = n + 1: Cells(i + 1, n) = .ClientEvVal:                  If i = 1 Then Cells(1, n) = "ClientEvVal"
        n = n + 1: Cells(i + 1, n) = .ClientEvName:                 If i = 1 Then Cells(1, n) = "ClientEvName"
        n = n + 1: Cells(i + 1, n) = .ClientEcadCode:               If i = 1 Then Cells(1, n) = "ClientEcadCode"
        n = n + 1: Cells(i + 1, n) = .MastersEducatedAllY:          If i = 1 Then Cells(1, n) = "MastersEducatedAllY"
        n = n + 1: Cells(i + 1, n) = .MastersEducatedPY:            If i = 1 Then Cells(1, n) = "MastersEducatedPY"
        n = n + 1: Cells(i + 1, n) = .MastersEducatedTY:            If i = 1 Then Cells(1, n) = "MastersEducatedTY"
        n = n + 1: Cells(i + 1, n) = .HairdressersNum:              If i = 1 Then Cells(1, n) = "HairdressersNum"
        n = n + 1: Cells(i + 1, n) = .HairdressersWorkPlace:        If i = 1 Then Cells(1, n) = "HairdressersWorkPlace"
        n = n + 1: Cells(i + 1, n) = .PartnerName:                  If i = 1 Then Cells(1, n) = "PartnerName"
        n = n + 1: Cells(i + 1, n) = .PartnerCode:                  If i = 1 Then Cells(1, n) = "PartnerCode"
        
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_TY_M1): If i = 1 Then Cells(1, n) = "CA_TY_M1"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_TY_M2): If i = 1 Then Cells(1, n) = "CA_TY_M2"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_TY_M3): If i = 1 Then Cells(1, n) = "CA_TY_M3"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_TY_M4): If i = 1 Then Cells(1, n) = "CA_TY_M4"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_TY_M5): If i = 1 Then Cells(1, n) = "CA_TY_M5"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_TY_M6): If i = 1 Then Cells(1, n) = "CA_TY_M6"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_TY_M7): If i = 1 Then Cells(1, n) = "CA_TY_M7"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_TY_M8): If i = 1 Then Cells(1, n) = "CA_TY_M8"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_TY_M9): If i = 1 Then Cells(1, n) = "CA_TY_M9"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_TY_M10): If i = 1 Then Cells(1, n) = "CA_TY_M10"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_TY_M11): If i = 1 Then Cells(1, n) = "CA_TY_M11"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_TY_M12): If i = 1 Then Cells(1, n) = "CA_TY_M12"

        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_PY_M1): If i = 1 Then Cells(1, n) = "CA_PY_M1"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_PY_M2): If i = 1 Then Cells(1, n) = "CA_PY_M2"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_PY_M3): If i = 1 Then Cells(1, n) = "CA_PY_M3"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_PY_M4): If i = 1 Then Cells(1, n) = "CA_PY_M4"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_PY_M5): If i = 1 Then Cells(1, n) = "CA_PY_M5"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_PY_M6): If i = 1 Then Cells(1, n) = "CA_PY_M6"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_PY_M7): If i = 1 Then Cells(1, n) = "CA_PY_M7"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_PY_M8): If i = 1 Then Cells(1, n) = "CA_PY_M8"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_PY_M9): If i = 1 Then Cells(1, n) = "CA_PY_M9"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_PY_M10): If i = 1 Then Cells(1, n) = "CA_PY_M10"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_PY_M11): If i = 1 Then Cells(1, n) = "CA_PY_M11"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_PY_M12): If i = 1 Then Cells(1, n) = "CA_PY_M12"

        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_TY_YTD1): If i = 1 Then Cells(1, n) = "CA_TY_YTD1"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_TY_YTD2): If i = 1 Then Cells(1, n) = "CA_TY_YTD2"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_TY_YTD3): If i = 1 Then Cells(1, n) = "CA_TY_YTD3"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_TY_YTD4): If i = 1 Then Cells(1, n) = "CA_TY_YTD4"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_TY_YTD5): If i = 1 Then Cells(1, n) = "CA_TY_YTD5"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_TY_YTD6): If i = 1 Then Cells(1, n) = "CA_TY_YTD6"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_TY_YTD7): If i = 1 Then Cells(1, n) = "CA_TY_YTD7"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_TY_YTD8): If i = 1 Then Cells(1, n) = "CA_TY_YTD8"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_TY_YTD9): If i = 1 Then Cells(1, n) = "CA_TY_YTD9"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_TY_YTD10): If i = 1 Then Cells(1, n) = "CA_TY_YTD10"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_TY_YTD11): If i = 1 Then Cells(1, n) = "CA_TY_YTD11"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_TY_YTD12): If i = 1 Then Cells(1, n) = "CA_TY_YTD12"

        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_PY_YTD1): If i = 1 Then Cells(1, n) = "CA_PY_YTD1"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_PY_YTD2): If i = 1 Then Cells(1, n) = "CA_PY_YTD2"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_PY_YTD3): If i = 1 Then Cells(1, n) = "CA_PY_YTD3"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_PY_YTD4): If i = 1 Then Cells(1, n) = "CA_PY_YTD4"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_PY_YTD5): If i = 1 Then Cells(1, n) = "CA_PY_YTD5"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_PY_YTD6): If i = 1 Then Cells(1, n) = "CA_PY_YTD6"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_PY_YTD7): If i = 1 Then Cells(1, n) = "CA_PY_YTD7"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_PY_YTD8): If i = 1 Then Cells(1, n) = "CA_PY_YTD8"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_PY_YTD9): If i = 1 Then Cells(1, n) = "CA_PY_YTD9"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_PY_YTD10): If i = 1 Then Cells(1, n) = "CA_PY_YTD10"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_PY_YTD11): If i = 1 Then Cells(1, n) = "CA_PY_YTD11"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_PY_YTD12): If i = 1 Then Cells(1, n) = "CA_PY_YTD12"

        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_TY_Q1): If i = 1 Then Cells(1, n) = "CA_TY_Q1"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_TY_Q2): If i = 1 Then Cells(1, n) = "CA_TY_Q2"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_TY_Q3): If i = 1 Then Cells(1, n) = "CA_TY_Q3"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_TY_Q4): If i = 1 Then Cells(1, n) = "CA_TY_Q4"

        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_PY_Q1): If i = 1 Then Cells(1, n) = "CA_PY_Q1"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_PY_Q2): If i = 1 Then Cells(1, n) = "CA_PY_Q2"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_PY_Q3): If i = 1 Then Cells(1, n) = "CA_PY_Q3"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_PY_Q4): If i = 1 Then Cells(1, n) = "CA_PY_Q4"
    End With

Next
myLib.VBA_End
End Sub
    



