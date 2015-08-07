<%@ Language=VBScript %>
<%Option Explicit%>
<%response.buffer=false%>
<!--#include file="aspboardfunctions.asp"-->
<html>
<head>
<meta name="keywords" content="">
<meta name="description" content="">
<meta name="author" content="">
<meta name="robots" content="index,follow">
<TITLE>Aulia Tradex</TITLE>
</head>

<%
	if abbgimage="" then
		if abbgcolor <> "" then
			response.write "<body bgcolor=""" & abbgcolor & """>"
		end if
	else
		response.write "<body background=""" & abbgimage & """>"
	end if			
%>
<body bgcolor="#000000">

<h3>
<font size="2" color="#edeef1" face="arial"><span style="font-weight: 400">
Specials<br>
</span><i><br>
</i></font>
<font face="arial" color="#edeef1"><font style="font-size: 14pt">
Be the first to know of our special offers, up dates and give-away
<i><font style="font-size: 9pt">.<br>
</font><span style="font-weight: 400"><font style="font-size: 8pt">
<i><br>
</i></font>
We value your privacy and our integrity, and your information will only be 
kept for this purpose and not sold or given away.</font></i></span><i><font face="arial" size="2" style="font-size: 8pt"><span style="font-weight: 400">
</span>
 </font>
</i>
 </font>
<font face="arial" size="2"><font face="arial" style="font-size: 8pt"></h3>
 </font>
<font color="#edeef1">
<br>
 </font>
 </font>
<font color="#edeef1">
</body>
</font>
<p><img border="0" src="images/msg.gif"><font face="arial" size="2"><font color="#ffcc00"><span style="font-weight: 400">&nbsp;</span></font><span style="font-weight: 400"><font color="#5f7aad"><span style="font-style: 12pt">
</span> </font> 
</span>
 </font>
<font color="#5f7aad">
<span style="font-weight: 400; font-style:12pt"><sup><font face="arial" style="font-size: 12">&nbsp;</font></sup></span></font><font color="#ffcc00"><sup><font face="arial" style="font-size: 11pt"><i style="font-style: 12pt"><u><a target="_self" href="aspboardpost.asp?id=0"><font color="#edeef1">Enter</font></a><font color="#edeef1">
</font> 
</u>
 </i>
 </font>
</sup> </font> <sup><a href="aspboardpost.asp?id=0">
<font face="arial" style="font-size: 11pt" color="#edeef1">Info</font></a></sup></p>

</html>