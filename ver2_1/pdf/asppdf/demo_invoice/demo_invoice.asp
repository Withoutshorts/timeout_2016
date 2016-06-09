<%@ Language=VBScript %>

<HTML>
<HEAD>
<title>AspPDF Document Creation Demo</title>
</HEAD>
<BODY BGCOLOR="#FFFFFF">
<FONT FACE="Arial" SIZE="2">
<h3>AspPDF Document Creation Demo</h3>
AspPDF can generate entire PDF files programmatically, including tables and images.<br>This page demonstrates many of AspPDF's capabilities by demonstrating automatic creation of an invoice.<br>&nbsp;

<TABLE WIDTH=720 CELLSPACING=0 CELLPADDING=0 BGCOLOR="#EBC01A" BORDER=0>
	<TD>
	
	<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=2 BGCOLOR="#FFFFFF" WIDTH="720">

	<FORM METHOD="POST" ACTION="demo_makeinvoice.asp" NAME="InvoiceForm" ENCTYPE="multipart/form-data">
	<INPUT TYPE=HIDDEN NAME="OrderNumber" VALUE="">
	<TR>
		<TD><FONT FACE="Arial" SIZE=2><B>Company Name/Address:</B></FONT></TD>
		<TD><TEXTAREA COLS=25 ROWS=4 NAME="Company">My Company
140 Broadway
New York, NY 10005</TEXTAREA></TD>
		<TD WIDTH=200></TD>
	</TR>
	<TR>
		<TD><FONT FACE="Arial" SIZE=2><B>Upload a Logo Image:<br></B></FONT>
		    <FONT FACE="Arial" SIZE=1><i>This function uses <a href="http://www.aspupload.com">AspUpload</a></i></FONT>
		</TD>
		<TD><INPUT TYPE="FILE" NAME="File1"></TD>
	</TR>
	<TR>
		<TD><FONT FACE="Arial" SIZE=2><B>Purchase Order Number:</B></FONT></TD>
		<TD><INPUT TYPE=TEXT SIZE=30 NAME="PurchaseOrder" VALUE="1234-5678"></TD>
	</TR>
	<TR>
		<TD><FONT FACE="Arial" SIZE=2><B>Bill To:</B></FONT></TD>
		<TD><TEXTAREA COLS=25 ROWS=4 NAME="BillTo">My Customer
Accounts Payable
500 Main St.
Beverly Hills, CA 90210</TEXTAREA></TD>
	</TR>
	<TR>
		<TD><FONT FACE="Arial" SIZE=2><B>Ship To:</B></FONT></TD>
		<TD><TEXTAREA COLS=25 ROWS=4 NAME="ShipTo">
My Customer
100 Rodeo Dr.
Beverly Hills, CA 90210
		</TEXTAREA></TD>
	</TR>
	<TR>
		<TD VALIGN=TOP><FONT FACE="Arial" SIZE=2><B>Notes:</B></FONT></TD>
		<TD><TEXTAREA COLS=25 NAME="Notes"></TEXTAREA></TD>
	</TR>
	<TR>
		<TD COLSPAN="3">
			<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=0>
				<TR><FONT FACE="Arial" SIZE=2><TD>Product:</TD>
					<TD>Quantity:</TD>
					<TD>Price:</TD>
					<TD>Amount:</TD></FONT></TR>
			<TR><TD><INPUT NAME="Product1" VALUE="Dodge Viper GTS"></TD>
			    <TD><INPUT NAME="Qty1" VALUE="1"></TD>
			    <TD><INPUT NAME="Price1" VALUE="88000"></TD>
			    <TD><INPUT NAME="Amount1" VALUE="88000"></TD></TR>
			<TR><TD><INPUT NAME="Product2" VALUE=""></TD>
			    <TD><INPUT NAME="Qty2" VALUE=""></TD>
			    <TD><INPUT NAME="Price2" VALUE=""></TD>
			    <TD><INPUT NAME="Amount2" VALUE=""></TD></TR>
			<TR><TD COLSPAN="2"></TD>
				<TD align="right">Total:&nbsp;&nbsp;</TD>
				<TD><INPUT NAME="Total" SIZE="10" VALUE="88000"></td></tr>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD COLSPAN=3 ALIGN="left">
			To view the invoice, click this button:
			<INPUT TYPE=SUBMIT NAME="Submit" VALUE="Generate PDF"><BR>
		</TD>
	</TR>

	</FORM>
	</TABLE>
	
	</TD>
</TABLE>

<P>
<B><A HREF="demo_invoice.zip">Download source code (ASP) for this demo</A></B>
</font>
</HTML>
