<%@ Language=VBScript %>
<%
set oUpload = Server.CreateObject("Persits.Upload")
oUpload.Save("c:\upload")

Set oPDF = Server.CreateObject("Persits.PDFManager")
Set oDoc = oPDF.CreateDocument

Set oPage = oDoc.Pages.Add

Set oDest = oDoc.CreateDest(oPage, "Fit=True")
oDoc.OpenAction = oDest

Set oHelvFont = oDoc.Fonts("Helvetica")

set oCanvas = oPage.Canvas

if oUpload.Files.Count = 1 then
	set oLogo = oUpload.Files(1)
	if oLogo.ImageType <> "UNKNOWN" then
		lImageScale = 1
		if oLogo.ImageHeight > 100 then
			lImageScale = 100 / oLogo.ImageHeight
		end if
		if oLogo.ImageWidth > 240 then
			lImageScale2 = 240 / oLogo.ImageWidth
			if lImageScale2 < lImageScale then
				lImageScale = lImageScale2
			end if
		end if
		lImageScale = round(lImageScale, 2)
		sImageParam = "x=54;y=660;scalex=" & lImageScale & ";scaley=" & lImageScale
		'Response.Write simageparam
		'Response.end
		oCanvas.DrawImage oDoc.OpenImage(oLogo.Path), sImageParam
	end if
else
	oCanvas.DrawText "Logo Goes Here", "x=54;y=730;size=20", oHelvFont
end if

sAddress = oupload.form("Company")
oCanvas.DrawText sAddress, "x=54; y=650; width=190; size=10; alignment=left; rendering=0", oHelvFont

oCanvas.DrawText "Invoice", "x=320, y=750; size=70; rendering=1", oHelvFont

sNumber = Session.SessionID

oCanvas.DrawText "Number:", "x=320, y=665, size=12; rendering=0", oHelvFont
oCanvas.DrawText "Date:", "x=320, y=645, size=12; rendering=0", oHelvFont
oCanvas.DrawText sNumber, "x=390, y=665, size=12; rendering=0", oHelvFont

oCanvas.DrawText Formatdatetime(Now, vbLongDate), "x=390, y=645, size=12; rendering=0", oHelvFont
	

'==============================================================Bill To and Ship To boxes

set oTable = oDoc.CreateTable("rows=2; cols=1; cellspacing=-1; cellpadding=2; height=122; width=220; border=0")
oTable.Font = oHelvFont

With oTable.Rows(1)
	.Height = 22
	.Cells(1).AddText "BILL TO", "alignment=left;color=black;size=12"
end with
with oTable.rows(2)
	.Cells(1).AddText oupload.form("BillTo"), "alignment=left;color=black;size=12"
end with
oCanvas.DrawTable oTable, "x=54;y=590"

sName = oupload.form("Name")
with oTable.Rows(1).Cells(1)
	.ClearText
	.AddText "SHIP TO", "alignment=left;color=black;size=12"
end with
with oTable.Rows(2).Cells(1)
	.ClearText
	.AddText oupload.form("ShipTo"), "alignment=left;color=black;size=12"
end with

oCanvas.DrawTable oTable, "x=320;y=590"

'================================================================Header Box
set oTable = oDoc.CreateTable("rows=2;cols=6;cellpadding=2;height=60;width=500;border=0;bordercolor=black")
oTable.Font = oHelvFont

with oTable.rows(1)
	.Height = 22
	.bgcolor = "lightgray"
	sParam = "alignment=2;color=black;size=12"
	.cells(1).addtext "P.O. #", sParam 
	.cells(2).addtext "Terms", sParam 
	.cells(3).addtext "Ship Date", sParam 
	.cells(4).addtext "Ship Via", sParam 
	.cells(5).colspan = 2
	.cells(5).addtext "Notes", sParam 
end with

with oTable.Rows(2)
	.Height = 28
	.bgcolor = "white"
	.cells(1).addtext oupload.form("PurchaseOrder"), sParam
	.cells(2).addtext "Net 30", sParam
	.cells(3).addtext FormatDateTime(Now, vbShortDate), sParam
	.cells(4).addtext "UPS Ground", sParam 
	.cells(5).colspan = 2
	.cells(5).addtext oupload.form("Notes"), "alignment=2;color=black;size=10"
end with

oCanvas.DrawTable oTable, "x=54; y=452"

'//=================================================================Main Box
set oTable = oDoc.CreateTable("rows=2;cols=4;cellpadding=2;height=300;width=500;border=0")
oTable.Font = oHelvFont

sParam = "alignment=left;color=black;size=12"
sParam2 = "alignment=right; color=black; size=12"

with oTable.Rows(1)
	.Height = 22
	.BgColor = "lightgray"
	.Cells(1).width = 310
	.cells(1).addtext "Product", sParam
	.Cells(2).width = 30
	.cells(2).addtext "Qty", sParam2
	.Cells(3).width = 80
	.cells(3).addtext "Price", sParam2
	.Cells(4).width = 80
	.cells(4).addtext "Amount", sParam2
end with
with oTable.rows(2)
	.Height = 278
	.BgColor = "white"
	.Cells(1).width = 310
	.Cells(2).width = 30
	.Cells(3).width = 80
	.Cells(4).width = 80
	
	lTotal = 0	
	for i = 1 to 2
		sProductBox = sProductBox & oupload.form("Product" & i) & vbcrlf
		sQtyBox = sQtyBox & oupload.form("Qty" & i) & vbcrlf
		Response.Write i
		on error resume next 'skip type conversion errors
		Price = oUpload.Form("Price" & i)
		Price = "$ " & oPDF.FormatNumber(Price, "precision=2, delimiter=true")
		sPriceBox = sPriceBox & Price & vbcrlf
		Amount = oUpload.Form("Amount" & i)
		Amount = "$ " & oPDF.FormatNumber(Amount, "precision=2, delimiter=true")
		sAmountBox = sAmountBox & Amount & vbcrlf
		on error goto 0
	next
	.cells(1).addtext sProductBox, sParam
	.cells(2).addtext sQtyBox, sParam2
	.cells(3).addtext sPriceBox, sParam2
	.cells(4).addtext sAmountBox, sParam2
end with
oCanvas.DrawTable oTable, "x=54; y=380"
oCanvas.DrawText "STATUS:  Please remit payment", "alignment=left; color=black; size=12; x=54; y=78", oHelvFont

' Send directly to browser, do not save a temporary copy on disk
oDoc.SaveHttp "attachment;filename=invoice.pdf"

%>