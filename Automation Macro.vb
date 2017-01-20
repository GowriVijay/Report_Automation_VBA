Sub Excel_Macro()
Dim StartTime, StopTime, ElapsedTime As Variant
Dim IntialName As String
Dim sFileSaveName As Variant
StartTime = Time
    Call NPS
    Call Comp_Buyer
    Call Comp_Seller
    Call PPTx
StopTime = Time
ElapsedTime = (StopTime - StartTime) * 24 * 60 * 60
MsgBox "Time taken = " & ElapsedTime & " Secs."
IntialName = "Competitve Buyer - "
sFileSaveName = Application.GetSaveAsFilename(InitialFileName:=InitialName, fileFilter:="Excel Files (*.xlsm), *.xlsm")
If sFileSaveName <> False Then
ActiveWorkbook.SaveAs sFileSaveName
Workbooks.Open Filename:="C:\Users\vthangamuthu\Documents\Vijay\Sam\NPS Deck Reformatting\Q4'16 Data\Macro\Competitive Analysis Macro.xlsm"
ThisWorkbook.Close
ActiveWorkbook.Worksheets("Introduction Page").Activate
Application.DisplayFullScreen = True
End If
End Sub


'*****************************************************************
Public Sub NPS()
Dim NPS_WB As Variant
Dim R, C As Integer
Dim Buyer, B2C, C2C As Single
Dim NPS As Workbook
Dim Slide2, Slide3 As Worksheet

'***************** This Section of the code is used to import the NPS Raw Data into the Competitve Analysis Macro
Qtr = InputBox("Enter the Quarter that you are Reporting this data on (e.g., Q3'16): ", "Reporting Quarter", "Q4'16")
MsgBox "Go ahead and select NPS " & Qtr & " Raw Data"
NPS_WB = Application.GetOpenFilename(FileFilter:="Excel Files,*.xl*;*.xm*")

If NPS_WB <> False Then
    Set NPS = Workbooks.Open(NPS_WB)
Else
    MsgBox "NPS file is not Chosen, Kindly Re-run the Macro!!!"
    Exit Sub
End If

'****** Switching off Screen Updating
Application.ScreenUpdating = False

'************ Slide 2a - US NPS **********
NPS.Worksheets("Big 4").Activate
C = Cells(1, Columns.Count).End(xlToLeft).Column - 1
Buyer = Cells(2, C).Value
B2C = Cells(4, C).Value
C2C = Cells(5, C).Value

Set Slide2 = ThisWorkbook.Worksheets("Slide 2")
Slide2.Activate
R = Application.Match(Qtr, Worksheets("Slide 2").Columns(1), 0)
Cells(R, 2).Value = Buyer
Cells(R, 3).Value = B2C
Cells(R, 4).Value = C2C

'************ Slide 2b - UK NPS **********
NPS.Worksheets("Big 4").Activate
Buyer = Cells(8, C).Value
B2C = Cells(10, C).Value
C2C = Cells(11, C).Value

Slide2.Activate
Cells(R, 8).Value = Buyer
Cells(R, 10).Value = B2C
Cells(R, 11).Value = C2C

'************ Slide 3a - DE NPS **********
NPS.Worksheets("Big 4").Activate
Buyer = Cells(14, C).Value
B2C = Cells(16, C).Value
C2C = Cells(17, C).Value

Set Slide3 = ThisWorkbook.Worksheets("Slide 3")
Slide3.Activate
Cells(R, 2).Value = Buyer
Cells(R, 4).Value = B2C
Cells(R, 5).Value = C2C

'************ Slide 3b - AU NPS **********
NPS.Worksheets("Big 4").Activate
Buyer = Cells(20, C).Value
B2C = Cells(22, C).Value
C2C = Cells(23, C).Value

Slide3.Activate
Cells(R, 8).Value = Buyer
Cells(R, 9).Value = B2C
Cells(R, 10).Value = C2C

NPS.Activate
ActiveWorkbook.Close SaveChanges:=False
MsgBox "NPS Data is updated... Go ahead and choose Competitve Buyer Data"
Application.ScreenUpdating = True

End Sub

'**************************************************
Public Sub Comp_Buyer()
Dim Buyer_WB, Horizontal, Fashion, Electronics, Parts As Variant
Dim Horizontal_1, Fashion_1, Electronics_1, Parts_1 As Variant
'Dim R, C As Integer
Dim Buyer As Workbook
Dim Slide5, Slide6, Slide7, Slide8, Slide9 As Worksheet

'***************** This Section of the code is used to import the Buyer Raw Data into the Competitve Analysis Macro
'Qtr = InputBox("Enter the Quarter that you are Reporting this data on (e.g., Q3'16): ", "Reporting Quarter", "Q4'16")
MsgBox "Go ahead and select Buyer " & Qtr & " Raw Data"
Buyer_WB = Application.GetOpenFilename(FileFilter:="Excel Files,*.xl*;*.xm*")

If Buyer_WB <> False Then
    Set Buyer = Workbooks.Open(Buyer_WB)
Else
    MsgBox "Buyer file is not Chosen, Kindly Re-run the Macro!!!"
    Exit Sub
End If

'****** Switching off Screen Updating
Application.ScreenUpdating = False

'************ Slide 5a - Unaided Consideration ***********

'*********************************************************
'***************Not yet Updated***************************
'*********************************************************

'************ Slide 5b - First Destination when Shopping **********
Buyer.Worksheets("US-Screen").Activate
Horizontal = Range("B33:B35").Value
Fashion = Range("F33:F35").Value
Electronics = Range("J33:J35").Value
Parts = Range("N33:N35").Value

Set Slide5 = ThisWorkbook.Worksheets("Slide 5")
Slide5.Activate
Range("H2:J2").Value = Application.WorksheetFunction.Transpose(Horizontal)
Range("H3:J3").Value = Application.WorksheetFunction.Transpose(Fashion)
Range("H4:J4").Value = Application.WorksheetFunction.Transpose(Electronics)
Range("H5:J5").Value = Application.WorksheetFunction.Transpose(Parts)

'************ Slide 6a - Overall Satisfaction **********
Buyer.Worksheets("US-Main").Activate
Horizontal = Range("B42:E42").Value
Fashion = Range("H42:K42").Value
Electronics = Range("N42:Q42").Value
Parts = Range("T42:W42").Value

Set Slide6 = ThisWorkbook.Worksheets("Slide 6")
Slide6.Activate
Range("B14:E14").Value = Horizontal
Range("B15:E15").Value = Fashion
Range("B16:E16").Value = Electronics
Range("B17:E17").Value = Parts

'************ Slide 6b - Purchase Intent **********
Buyer.Worksheets("US-Main").Activate
Horizontal = Range("B142:E142").Value
Fashion = Range("H142:K142").Value
Electronics = Range("N142:Q142").Value
Parts = Range("T142:W142").Value

Slide6.Activate
Range("H14:K14").Value = Horizontal
Range("H15:K15").Value = Fashion
Range("H16:K16").Value = Electronics
Range("H17:K17").Value = Parts

'************ Slide 7 - First Five Buyer attributes **********
Buyer.Worksheets("Attribute").Activate
Horizontal = Range("B3:B7").Value
Fashion = Range("H3:H7").Value
Electronics = Range("N3:N7").Value
Parts = Range("T3:T7").Value

Horizontal_1 = Range("D3:D7").Value
Fashion_1 = Range("J3:J7").Value
Electronics_1 = Range("P3:P7").Value
Parts_1 = Range("V3:V7").Value

Set Slide7 = ThisWorkbook.Worksheets("Slide 7")
Slide7.Activate
Range("B3:F3").Value = Application.WorksheetFunction.Transpose(Horizontal)
Range("B6:F6").Value = Application.WorksheetFunction.Transpose(Fashion)
Range("B9:F9").Value = Application.WorksheetFunction.Transpose(Electronics)
Range("B12:F12").Value = Application.WorksheetFunction.Transpose(Parts)

Range("B4:F4").Value = Application.WorksheetFunction.Transpose(Horizontal_1)
Range("B7:F7").Value = Application.WorksheetFunction.Transpose(Fashion_1)
Range("B10:F10").Value = Application.WorksheetFunction.Transpose(Electronics_1)
Range("B13:F13").Value = Application.WorksheetFunction.Transpose(Parts_1)

'************ Slide 8 - Next Five Buyer attributes **********
Buyer.Worksheets("Attribute").Activate
Horizontal = Range("B8:B12").Value
Fashion = Range("H8:H12").Value
Electronics = Range("N8:N12").Value
Parts = Range("T8:T12").Value

Horizontal_1 = Range("D8:D12").Value
Fashion_1 = Range("J8:J12").Value
Electronics_1 = Range("P8:P12").Value
Parts_1 = Range("V8:V12").Value

Set Slide8 = ThisWorkbook.Worksheets("Slide 8")
Slide8.Activate
Range("B3:F3").Value = Application.WorksheetFunction.Transpose(Horizontal)
Range("B6:F6").Value = Application.WorksheetFunction.Transpose(Fashion)
Range("B9:F9").Value = Application.WorksheetFunction.Transpose(Electronics)
Range("B12:F12").Value = Application.WorksheetFunction.Transpose(Parts)

Range("B4:F4").Value = Application.WorksheetFunction.Transpose(Horizontal_1)
Range("B7:F7").Value = Application.WorksheetFunction.Transpose(Fashion_1)
Range("B10:F10").Value = Application.WorksheetFunction.Transpose(Electronics_1)
Range("B13:F13").Value = Application.WorksheetFunction.Transpose(Parts_1)

'************ Slide 9 - Last Five Buyer attributes **********
Buyer.Worksheets("Attribute").Activate
Horizontal = Range("B13:B17").Value
Fashion = Range("H13:H17").Value
Electronics = Range("N13:N17").Value
Parts = Range("T13:T17").Value

Horizontal_1 = Range("D13:D17").Value
Fashion_1 = Range("J13:J17").Value
Electronics_1 = Range("P13:P17").Value
Parts_1 = Range("V13:V17").Value

Set Slide9 = ThisWorkbook.Worksheets("Slide 9")
Slide9.Activate
Range("B3:F3").Value = Application.WorksheetFunction.Transpose(Horizontal)
Range("B6:F6").Value = Application.WorksheetFunction.Transpose(Fashion)
Range("B9:F9").Value = Application.WorksheetFunction.Transpose(Electronics)
Range("B12:F12").Value = Application.WorksheetFunction.Transpose(Parts)

Range("B4:F4").Value = Application.WorksheetFunction.Transpose(Horizontal_1)
Range("B7:F7").Value = Application.WorksheetFunction.Transpose(Fashion_1)
Range("B10:F10").Value = Application.WorksheetFunction.Transpose(Electronics_1)
Range("B13:F13").Value = Application.WorksheetFunction.Transpose(Parts_1)

'************ Slide 16 - Base for all 15 Buyer attributes **********
Buyer.Worksheets("Attribute").Activate
NSize = Range("B41:I55").Value

Set Slide16 = ThisWorkbook.Worksheets("Slide 16")
Slide16.Activate
Range("C4:J18").Value = NSize

'************ Final Touch-ups
Buyer.Activate
ActiveWorkbook.Close SaveChanges:=False
MsgBox "Buyer Data is updated... Go ahead and choose Competitve Seller Data"
Application.ScreenUpdating = True

End Sub

'**********************************************************
Public Sub Comp_Seller()
Dim Seller_WB, eBay, Amazon, Facebook, Craigslist As Variant
Dim FM, TM, Satisfaction, SI As Variant
'Dim R, C As Integer
Dim Seller As Workbook
Dim Slide11, Slide12, Slide13, Slide14, Slide9 As Worksheet

'***************** This Section of the code is used to import the Seller Raw Data into the Competitve Analysis Macro
'Qtr = InputBox("Enter the Quarter that you are Reporting this data on (e.g., Q3'16): ", "Reporting Quarter", "Q4'16")
MsgBox "Go ahead and select Seller " & Qtr & " Raw Data"
Seller_WB = Application.GetOpenFilename(FileFilter:="Excel Files,*.xl*;*.xm*")

If Seller_WB <> False Then
    Set Seller = Workbooks.Open(Seller_WB)
Else
    MsgBox "Seller file is not Chosen, Kindly Re-run the Macro!!!"
    Exit Sub
End If

'****** Switching off Screen Updating
Application.ScreenUpdating = False

'************ Slide 11a - Incidence Rate ***********
Seller.Worksheets("C2C FINAL").Activate
FM = Range("B30:B34").Value
TM = Range("B50:B54").Value
Set Slide11 = ThisWorkbook.Worksheets("Slide 11")
Slide11.Activate
Range("B2:B6").Value = FM
Range("D2:D6").Value = TM

'************ Slide 11b - Overall Satisfaction ***********
Seller.Worksheets("C2C FINAL").Activate
Satisfaction = Range("B88:K88").Value
Slide11.Activate
Range("G9:P9").Value = Satisfaction

'************ Slide 11c - Selling Intent ***********
Seller.Worksheets("C2C FINAL").Activate
Satisfaction = Range("B137:K137").Value
Slide11.Activate
Range("G11:P11").Value = Satisfaction

'************ Slide 12 - First Five Seller attributes **********
Seller.Worksheets("Attribute").Activate
eBay = Range("B3:B7").Value
Amazon = Range("D3:D7").Value
Facebook = Range("F3:F7").Value
Craigslist = Range("H3:H7").Value

Set Slide12 = ThisWorkbook.Worksheets("Slide 12")
Slide12.Activate
Range("B2:F2").Value = Application.WorksheetFunction.Transpose(eBay)
Range("B4:F4").Value = Application.WorksheetFunction.Transpose(Amazon)
Range("B6:F6").Value = Application.WorksheetFunction.Transpose(Facebook)
Range("B8:F8").Value = Application.WorksheetFunction.Transpose(Craigslist)

'************ Slide 13 - Next Five Seller attributes **********
Seller.Worksheets("Attribute").Activate
eBay = Range("B8:B12").Value
Amazon = Range("D8:D12").Value
Facebook = Range("F8:F12").Value
Craigslist = Range("H8:H12").Value

Set Slide13 = ThisWorkbook.Worksheets("Slide 13")
Slide13.Activate
Range("B2:F2").Value = Application.WorksheetFunction.Transpose(eBay)
Range("B4:F4").Value = Application.WorksheetFunction.Transpose(Amazon)
Range("B6:F6").Value = Application.WorksheetFunction.Transpose(Facebook)
Range("B8:F8").Value = Application.WorksheetFunction.Transpose(Craigslist)

'************ Slide 14 - Last Four Seller attributes **********
Seller.Worksheets("Attribute").Activate
eBay = Range("B13:B16").Value
Amazon = Range("D13:D16").Value
Facebook = Range("F13:F16").Value
Craigslist = Range("H13:H16").Value


Set Slide14 = ThisWorkbook.Worksheets("Slide 14")
Slide14.Activate
Range("B2:E2").Value = Application.WorksheetFunction.Transpose(eBay)
Range("B4:E4").Value = Application.WorksheetFunction.Transpose(Amazon)
Range("B6:E6").Value = Application.WorksheetFunction.Transpose(Facebook)
Range("B8:E8").Value = Application.WorksheetFunction.Transpose(Craigslist)


'************ Final Touch-ups
Seller.Activate
ActiveWorkbook.Close SaveChanges:=False
MsgBox "Seller Data is updated..."
Application.ScreenUpdating = True

End Sub
'***************************************************************************************
'Ensure you have Reference 'Microsoft PowerPoint 16.0 object library

Public Sub PPTx()
Dim PPT As PowerPoint.Application
Dim oPPTFile As PowerPoint.Presentation
Dim Data, ws As Excel.Worksheet
Dim wb As Excel.Workbook
Dim cht As Chart
Dim R, C As Integer
Dim PPT_Tmplt As Variant

Application.ScreenUpdating = False

Set PPT = New PowerPoint.Application
'PPT.Visible = msoTrue
'PPT.ActiveWindow.ViewType = ppViewSlide

Set Data = ThisWorkbook.Worksheets("Introduction Page")
'Set oPPTFile = PPT.Presentations.Open("C:\Users\vthangamuthu\Documents\Vijay\Sam\NPS Deck Reformatting\Q4'16 Data\Macro\Template.pptx")
PPT_Tmplt = Application.GetOpenFilename
Set oPPTFile = PPT.Presentations.Open(PPT_Tmplt)

'******************** Slide2 ********************
Set Slide2 = ThisWorkbook.Worksheets("Slide 2")
Slide2.Activate
R = Cells(Rows.Count, 2).End(xlUp).Row 'To choose the last Row
Rs = R - 11 'R - 11 will give the same number of data points as in the Template for updating
'US - NPS
Range(Cells(Rs, 1), Cells(R, 4)).Copy
Set wb = PPT.ActivePresentation.Slides(2).Shapes("Chart 35").Chart.ChartData.Workbook
wb.Worksheets(1).Cells(2, 1).PasteSpecial ppPasteText
'UK - NPS
Range(Cells(Rs, 7), Cells(R, 11)).Copy
Set wb = PPT.ActivePresentation.Slides(2).Shapes("Chart 43").Chart.ChartData.Workbook
wb.Worksheets(1).Cells(2, 1).PasteSpecial ppPasteText

'******************** Slide3 ********************
Set Slide3 = ThisWorkbook.Worksheets("Slide 3")
Slide3.Activate
R = Cells(Rows.Count, 2).End(xlUp).Row 'To choose the last Row
Rs = R - 11 'R - 11 will give the same number of data points as in the Template for updating
'DE - NPS
Range(Cells(Rs, 1), Cells(R, 5)).Copy
Set wb = PPT.ActivePresentation.Slides(3).Shapes("Chart 47").Chart.ChartData.Workbook
wb.Worksheets(1).Cells(2, 1).PasteSpecial ppPasteText
'AU - NPS
Range(Cells(Rs, 7), Cells(R, 10)).Copy
Set wb = PPT.ActivePresentation.Slides(3).Shapes("Chart 49").Chart.ChartData.Workbook
wb.Worksheets(1).Cells(2, 1).PasteSpecial ppPasteText

'******************** Slide5a ********************

'******************** Slide5b - First Destination ********************
Set Slide5 = ThisWorkbook.Worksheets("Slide 5")
Slide5.Activate
Range("G1:J5").Copy
Set wb = PPT.ActivePresentation.Slides(5).Shapes("Chart 35").Chart.ChartData.Workbook
wb.Worksheets(1).Cells(1, 1).PasteSpecial ppPasteText



'******************** Slide6a - Overall Satisfaction ********************
Set Slide6 = ThisWorkbook.Worksheets("Slide 6")
Slide6.Activate
Range("A1:C5").Copy
Set wb = PPT.ActivePresentation.Slides(6).Shapes("Chart 35").Chart.ChartData.Workbook
wb.Worksheets("Sheet1").Cells(1, 1).PasteSpecial ppPasteText
'******************** Slide6b - Purchase Intent ********************
Set Slide6 = ThisWorkbook.Worksheets("Slide 6")
Slide6.Activate
Range("G1:I5").Copy
Set wb = PPT.ActivePresentation.Slides(6).Shapes("Chart 74").Chart.ChartData.Workbook
wb.Worksheets("Sheet1").Cells(1, 1).PasteSpecial ppPasteText



'******************** Slide7a - Horizontal ********************
Set Slide7 = ThisWorkbook.Worksheets("Slide 7")
Slide7.Activate
Range("B3:F4").Copy
Set wb = PPT.ActivePresentation.Slides(7).Shapes("Chart 77").Chart.ChartData.Workbook
wb.Worksheets("Awareness").Cells(4, 3).PasteSpecial ppPasteText
'******************** Slide7b - Fashion ********************
Set Slide7 = ThisWorkbook.Worksheets("Slide 7")
Slide7.Activate
Range("B6:F7").Copy
Set wb = PPT.ActivePresentation.Slides(7).Shapes("Chart 125").Chart.ChartData.Workbook
wb.Worksheets("Awareness").Cells(4, 3).PasteSpecial ppPasteText
'******************** Slide7c - Electronics ********************
Set Slide7 = ThisWorkbook.Worksheets("Slide 7")
Slide7.Activate
Range("B9:F10").Copy
Set wb = PPT.ActivePresentation.Slides(7).Shapes("Chart 140").Chart.ChartData.Workbook
wb.Worksheets("Awareness").Cells(4, 3).PasteSpecial ppPasteText
'******************** Slide7d - Parts ********************
Set Slide7 = ThisWorkbook.Worksheets("Slide 7")
Slide7.Activate
Range("B12:F13").Copy
Set wb = PPT.ActivePresentation.Slides(7).Shapes("Chart 157").Chart.ChartData.Workbook
wb.Worksheets("Awareness").Cells(4, 3).PasteSpecial ppPasteText




'******************** Slide8a - Horizontal ********************
Set Slide8 = ThisWorkbook.Worksheets("Slide 8")
Slide8.Activate
Range("B3:F4").Copy
Set wb = PPT.ActivePresentation.Slides(8).Shapes("Chart 77").Chart.ChartData.Workbook
wb.Worksheets("Awareness").Cells(4, 3).PasteSpecial ppPasteText
'******************** Slide8b - Fashion ********************
Set Slide8 = ThisWorkbook.Worksheets("Slide 8")
Slide8.Activate
Range("B6:F7").Copy
Set wb = PPT.ActivePresentation.Slides(8).Shapes("Chart 125").Chart.ChartData.Workbook
wb.Worksheets("Awareness").Cells(4, 3).PasteSpecial ppPasteText
'******************** Slide8c - Electronics ********************
Set Slide8 = ThisWorkbook.Worksheets("Slide 8")
Slide8.Activate
Range("B9:F10").Copy
Set wb = PPT.ActivePresentation.Slides(8).Shapes("Chart 140").Chart.ChartData.Workbook
wb.Worksheets("Awareness").Cells(4, 3).PasteSpecial ppPasteText
'******************** Slide8d - Parts ********************
Set Slide8 = ThisWorkbook.Worksheets("Slide 8")
Slide8.Activate
Range("B12:F13").Copy
Set wb = PPT.ActivePresentation.Slides(8).Shapes("Chart 157").Chart.ChartData.Workbook
wb.Worksheets("Awareness").Cells(4, 3).PasteSpecial ppPasteText




'******************** Slide9a - Horizontal ********************
Set Slide9 = ThisWorkbook.Worksheets("Slide 9")
Slide9.Activate
Range("B3:F4").Copy
Set wb = PPT.ActivePresentation.Slides(9).Shapes("Chart 77").Chart.ChartData.Workbook
wb.Worksheets("Awareness").Cells(4, 3).PasteSpecial ppPasteText
'******************** Slide9b - Fashion ********************
Set Slide9 = ThisWorkbook.Worksheets("Slide 9")
Slide9.Activate
Range("B6:F7").Copy
Set wb = PPT.ActivePresentation.Slides(9).Shapes("Chart 125").Chart.ChartData.Workbook
wb.Worksheets("Awareness").Cells(4, 3).PasteSpecial ppPasteText
'******************** Slide9c - Electronics ********************
Set Slide9 = ThisWorkbook.Worksheets("Slide 9")
Slide9.Activate
Range("B9:F10").Copy
Set wb = PPT.ActivePresentation.Slides(9).Shapes("Chart 140").Chart.ChartData.Workbook
wb.Worksheets("Awareness").Cells(4, 3).PasteSpecial ppPasteText
'******************** Slide9d - Parts ********************
Set Slide9 = ThisWorkbook.Worksheets("Slide 9")
Slide9.Activate
Range("B12:F13").Copy
Set wb = PPT.ActivePresentation.Slides(9).Shapes("Chart 157").Chart.ChartData.Workbook
wb.Worksheets("Awareness").Cells(4, 3).PasteSpecial ppPasteText



'******************** Slide11a ********************
Set Slide11 = ThisWorkbook.Worksheets("Slide 11")
Slide11.Activate
Range("A1:D6").Copy
Set wb = PPT.ActivePresentation.Slides(11).Shapes("Chart 73").Chart.ChartData.Workbook
wb.Worksheets("Sheet1").Cells(1, 1).PasteSpecial ppPasteText
'******************** Slide11b - First Destination ********************
Set Slide11 = ThisWorkbook.Worksheets("Slide 11")
Slide11.Activate
Range("G1:H6").Copy
Set wb = PPT.ActivePresentation.Slides(11).Shapes("Chart 49").Chart.ChartData.Workbook
wb.Worksheets("Sheet1").Cells(1, 1).PasteSpecial ppPasteText
'******************** Slide11c - First Destination ********************
Set Slide11 = ThisWorkbook.Worksheets("Slide 11")
Slide11.Activate
Range("K1:L6").Copy
Set wb = PPT.ActivePresentation.Slides(11).Shapes("Chart 61").Chart.ChartData.Workbook
wb.Worksheets("Sheet1").Cells(1, 1).PasteSpecial ppPasteText


'******************** Slide12a - eBay ********************
Set Slide12 = ThisWorkbook.Worksheets("Slide 12")
Slide12.Activate
Range("B2:F2").Copy
Set wb = PPT.ActivePresentation.Slides(12).Shapes("Chart 77").Chart.ChartData.Workbook
wb.Worksheets("Awareness").Cells(4, 3).PasteSpecial ppPasteText
'******************** Slide12b - Amazon ********************
Set Slide12 = ThisWorkbook.Worksheets("Slide 12")
Slide12.Activate
Range("B4:F4").Copy
Set wb = PPT.ActivePresentation.Slides(12).Shapes("Chart 125").Chart.ChartData.Workbook
wb.Worksheets("Awareness").Cells(4, 3).PasteSpecial ppPasteText
'******************** Slide12c - Facebook ********************
Set Slide12 = ThisWorkbook.Worksheets("Slide 12")
Slide12.Activate
Range("B6:D6").Copy
Set wb = PPT.ActivePresentation.Slides(12).Shapes("Chart 140").Chart.ChartData.Workbook
wb.Worksheets("Awareness").Cells(4, 3).PasteSpecial ppPasteText
'******************** Slide12d - Craigslist ********************
Set Slide12 = ThisWorkbook.Worksheets("Slide 12")
Slide12.Activate
Range("B8:D8").Copy
Set wb = PPT.ActivePresentation.Slides(12).Shapes("Chart 157").Chart.ChartData.Workbook
wb.Worksheets("Awareness").Cells(4, 3).PasteSpecial ppPasteText


'******************** Slide13a - eBay ********************
Set Slide13 = ThisWorkbook.Worksheets("Slide 13")
Slide13.Activate
Range("B2:F2").Copy
Set wb = PPT.ActivePresentation.Slides(13).Shapes("Chart 77").Chart.ChartData.Workbook
wb.Worksheets("Awareness").Cells(4, 3).PasteSpecial ppPasteText
'******************** Slide13b - Amazon ********************
Set Slide13 = ThisWorkbook.Worksheets("Slide 13")
Slide13.Activate
Range("B4:F4").Copy
Set wb = PPT.ActivePresentation.Slides(13).Shapes("Chart 125").Chart.ChartData.Workbook
wb.Worksheets("Awareness").Cells(4, 3).PasteSpecial ppPasteText
'******************** Slide13c - Facebook ********************
Set Slide13 = ThisWorkbook.Worksheets("Slide 13")
Slide13.Activate
Range("B6:D6").Copy
Set wb = PPT.ActivePresentation.Slides(13).Shapes("Chart 140").Chart.ChartData.Workbook
wb.Worksheets("Awareness").Cells(4, 3).PasteSpecial ppPasteText
'******************** Slide13d - Craigslist ********************
Set Slide13 = ThisWorkbook.Worksheets("Slide 13")
Slide13.Activate
Range("B8:D8").Copy
Set wb = PPT.ActivePresentation.Slides(13).Shapes("Chart 157").Chart.ChartData.Workbook
wb.Worksheets("Awareness").Cells(4, 3).PasteSpecial ppPasteText



'******************** Slide14a - eBay ********************
Set Slide14 = ThisWorkbook.Worksheets("Slide 14")
Slide14.Activate
Range("B2:E2").Copy
Set wb = PPT.ActivePresentation.Slides(14).Shapes("Chart 77").Chart.ChartData.Workbook
wb.Worksheets("Awareness").Cells(4, 3).PasteSpecial ppPasteText
'******************** Slide14b - Amazon ********************
Set Slide14 = ThisWorkbook.Worksheets("Slide 14")
Slide14.Activate
Range("B4:E4").Copy
Set wb = PPT.ActivePresentation.Slides(14).Shapes("Chart 125").Chart.ChartData.Workbook
wb.Worksheets("Awareness").Cells(4, 3).PasteSpecial ppPasteText
'******************** Slide14c - Facebook ********************
Set Slide14 = ThisWorkbook.Worksheets("Slide 14")
Slide14.Activate
Range("B6:C6").Copy
Set wb = PPT.ActivePresentation.Slides(14).Shapes("Chart 140").Chart.ChartData.Workbook
wb.Worksheets("Awareness").Cells(4, 3).PasteSpecial ppPasteText
'******************** Slide14d - Craigslist ********************
Set Slide14 = ThisWorkbook.Worksheets("Slide 14")
Slide14.Activate
Range("B8:C8").Copy
Set wb = PPT.ActivePresentation.Slides(14).Shapes("Chart 157").Chart.ChartData.Workbook
wb.Worksheets("Awareness").Cells(4, 3).PasteSpecial ppPasteText

'******************** Slide16 - Base for Slide 7 - 9 ********************
Set Slide16 = ThisWorkbook.Worksheets("Slide 16")
Slide16.Activate
Range("B2:J18").Copy
'PPT.ActiveWindow.ViewType = ppViewSlide

PPT.ActivePresentation.Slides(16).Shapes.Paste
MsgBox ("Save the Updated PPT")

sFileSaveName = Application.GetSaveAsFilename
With PPT.ActivePresentation
   .SaveAs sFileSaveName, ppSaveAsDefault
End With
PPT.ActivePresentation.Close

ActiveWorkbook.Worksheets("Introduction Page").Activate
End Sub