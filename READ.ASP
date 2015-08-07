<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer=False%>
<!--#INCLUDE FILE="aspboardfunctions.asp"-->
<HTML>
<HEAD>
<meta name="keywords" content="">
<meta name="description" content="">
<meta name="author" content="">
<meta name="robots" content="index,follow">
<TITLE>Aulia Tradex</TITLE>
</HEAD>
<body text="#E6E9EC" bgcolor="#000000">

<h3><font face="arial" size="2" color="#ECEEF0">New Messages</h3>
</font><font face="arial" size="2" color="#e6e9ec">
<%
	If abBGImage="" Then
		If abBGColor <> "" Then
			Response.Write "<BODY BGCOLOR=""" & abBGColor & """>"
		End If
	Else
		Response.Write "<BODY BACKGROUND=""" & abBGImage & """>"
	End If			
%> </font><font face="arial" size="2" color="#ECEEF0">
<br></font><font face="arial" size="2" color="#ECEEF0">
<%	
	ListItems(0)

%>
<br><br>
<br>
 </font>
<p>&nbsp;</p>
</BODY>
</HTML>