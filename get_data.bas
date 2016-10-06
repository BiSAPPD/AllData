Attribute VB_Name = "Module1"
Sub get_data()
Dim nm_brand As String, patch As String, ShIn As String

Dim ThisYear As Integer, cd_ActualMonth As Integer
Dim LastRow As Long

myLib.VBA_Start

nm_ActWb = ActiveWorkbook.Name
cd_ActualMonth = CInt(InputBox("Month"))
ThisYear = CInt(InputBox("YearEnd"))

ar_brand = Array("LP", "MX", "KR", "RD", "ES")
myLib.VBA_Start

Dim clnts As clsClients, clnt As clsClientInfo
Set clnts = New clsClients


For f_brnd = 0 To UBound(ar_brand)
    nm_brand = ar_brand(f_brnd)
    ShIn = nm_brand
    ShOut = "TR"
    patch = myLib.GetPatchHistTR(nm_brand, ThisYear, ThisYear, cd_ActualMonth, cd_ActualMonth)
    WbTR = myLib.OpenFile(patch, ShIn)
    Workbooks(WbTR).Activate
    Sheets(ShIn).Select
    clnts.FillFromSheet ActiveSheet, 2016, 8, nm_brand
    
    Workbooks(WbTR).Close
    Workbooks(nm_ActWb).Activate
Next f_brnd

myLib.CreateSh (ShOut)
myLib.sheetActivateCleer (ShOut)

i = 0
For Each clnt In clnts
    i = i + 1
    n = 0
    With clnt
        n = n + 1: Cells(i, n) = IIF(i = 1, "BrandName", .BrandName)
        n = n + 1: Cells(i, n) = IIF(i = 1, "DatabaseClientNum", .DatabaseClientNum)
        n = n + 1: Cells(i, n) = IIF(i = 1, "StatYear", .StatYear)
        n = n + 1: Cells(i, n) = IIF(i = 1, "StatMonth", .StatMonth)
        n = n + 1: Cells(i, n) = IIF(i = 1, "BrandName", .BrandName)
        n = n + 1: Cells(i, n) = IIF(i = 1, "TypeBusiness", .TypeBusiness)
        n = n + 1: Cells(i, n) = IIF(i = 1, "DatabaseClientNum", .DatabaseClientNum)
        n = n + 1: Cells(i, n) = IIF(i = 1, "DatabaseClientAndBrandNum", .DatabaseClientAndBrandNum)
        n = n + 1: Cells(i, n) = IIF(i = 1, "UniverseCode", .UniverseCode)
        n = n + 1: Cells(i, n) = IIF(i = 1, "UniversCodeAndBrand", .UniversCodeAndBrand)
        n = n + 1: Cells(i, n) = IIF(i = 1, "MregName", .MregName)
        n = n + 1: Cells(i, n) = IIF(i = 1, "ExtMregName", .ExtMregName)
        n = n + 1: Cells(i, n) = IIF(i = 1, "RegName", .RegName)
        n = n + 1: Cells(i, n) = IIF(i = 1, "FlsmName", .FlsmName)
        n = n + 1: Cells(i, n) = IIF(i = 1, "SecName", .SecName)
        n = n + 1: Cells(i, n) = IIF(i = 1, "SrepName", .SrepName)
        n = n + 1: Cells(i, n) = IIF(i = 1, "ClientName", .ClientName)
        n = n + 1: Cells(i, n) = IIF(i = 1, "ChainName", .ChainName)
        n = n + 1: Cells(i, n) = IIF(i = 1, "ChainCode", .ChainCode)
        n = n + 1: Cells(i, n) = IIF(i = 1, "GeoCity", .GeoCity)
        n = n + 1: Cells(i, n) = IIF(i = 1, "GeoReg", .GeoReg)
        n = n + 1: Cells(i, n) = IIF(i = 1, "ClientTypeRus", .ClientTypeRus)
        n = n + 1: Cells(i, n) = IIF(i = 1, "ClientTypeEng", .ClientTypeEng)
        n = n + 1: Cells(i, n) = IIF(i = 1, "ClientTypeEngChort", .ClientTypeEngChort)
        n = n + 1: Cells(i, n) = IIF(i = 1, "ClientTypeEngChain", .ClientTypeEngChain)
        n = n + 1: Cells(i, n) = IIF(i = 1, "ClubStatus", .ClubStatus)
        n = n + 1: Cells(i, n) = IIF(i = 1, "EmotionStatus", .EmotionStatus)
        n = n + 1: Cells(i, n) = IIF(i = 1, "CnqFullDate", .CnqFullDate)
        n = n + 1: Cells(i, n) = IIF(i = 1, "CnqYearGA", .CnqYearDate)
        n = n + 1: Cells(i, n) = IIF(i = 1, "CnqYearGA", .CnqYearGA)
        n = n + 1: Cells(i, n) = IIF(i = 1, "CnqMonthNum", .CnqMonthNum)
        n = n + 1: Cells(i, n) = IIF(i = 1, "CnqMonthNameRus", .CnqMonthNameRus)
        n = n + 1: Cells(i, n) = IIF(i = 1, "CnqMonthNameEng", .CnqMonthNameEng)
        n = n + 1: Cells(i, n) = IIF(i = 1, "MagType", .MagType)
        n = n + 1: Cells(i, n) = IIF(i = 1, "MagTypePrice", .MagTypePrice)
        n = n + 1: Cells(i, n) = IIF(i = 1, "MagTypePlace", .MagTypePlace)
        n = n + 1: Cells(i, n) = IIF(i = 1, "WorkStatusNum", .WorkStatusNum)
        n = n + 1: Cells(i, n) = IIF(i = 1, "WorkStatusName", .WorkStatusName)
        n = n + 1: Cells(i, n) = IIF(i = 1, "LtmAvgCaVal", .LtmAvgCaVal)
        n = n + 1: Cells(i, n) = IIF(i = 1, "LtmAvgCaName", .LtmAvgCaName)
        n = n + 1: Cells(i, n) = IIF(i = 1, "LtmFrqOrders", .LtmFrqOrders)
        n = n + 1: Cells(i, n) = IIF(i = 1, "ClientEvVal", .ClientEvVal)
        n = n + 1: Cells(i, n) = IIF(i = 1, "ClientEvName", .ClientEvName)
        n = n + 1: Cells(i, n) = IIF(i = 1, "ClientEcadCode", .ClientEcadCode)
        n = n + 1: Cells(i, n) = IIF(i = 1, "MastersEducatedAllY", .MastersEducatedAllY)
        n = n + 1: Cells(i, n) = IIF(i = 1, "MastersEducatedPY", .MastersEducatedPY)
        n = n + 1: Cells(i, n) = IIF(i = 1, "MastersEducatedTY", .MastersEducatedTY)
        n = n + 1: Cells(i, n) = IIF(i = 1, "HairdressersNum", .HairdressersNum)
        n = n + 1: Cells(i, n) = IIF(i = 1, "HairdressersWorkPlace", .HairdressersWorkPlace)
        n = n + 1: Cells(i, n) = IIF(i = 1, "PartnerName", .PartnerName)
        n = n + 1: Cells(i, n) = IIF(i = 1, "PartnerCode", .PartnerCode)
    End With

Next
myLib.VBA_End
End Sub
    


