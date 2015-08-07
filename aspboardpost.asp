<%@ language=vbscript %>
<%option explicit%>
<%response.buffer=true%>
<!--#include file="aspboardfunctions.asp"-->
<html>
<head>
<meta name="keywords" content="">
<meta name="description" content="">
<meta name="author" content="">
<meta name="robots" content="index,follow">
<%
	dim lngnewid
	dim strmsglabel	
	dim strtitlelabel
	dim strheaderstring
	dim strdetailstring	
	
	if request("author_namestr")<> "" then
		with response
			.cookies("postname") = request("author_namestr")
			.cookies("postname").expires = dateadd("yyyy",1,date)
			.cookies("postemail") = request("author_emailstr")
			.cookies("postemail").expires = dateadd("yyyy",1,date)
			.cookies("posturl") = request("author_urlstr")
			.cookies("posturl").expires = dateadd("yyyy",1,date)
		end with
		strdetailstring = request("detailstr")	
		if request("incorig") = "yes" then
			strdetailstring = strdetailstring & aspcrlf & " in response to: " & aspcrlf & request("origmsgstr")	
		end if
		lngnewid = addresponse(request("msgid"), request("headerstr"), strdetailstring, request("author_namestr"), request("author_emailstr") & "", request("author_urlstr") & "")		
		response.redirect "aspboarddetail.asp?id=" & lngnewid
	end if	
	
	if request.querystring("id") > 0 then
		strmsglabel = "your response"
		strtitlelabel = "post a follow-up to:"
	else
		strmsglabel = "your message"		
		strtitlelabel = "post a new message"
	end if
%>
<TITLE>Aulia Tradex&nbsp;<%=strTitleLabel%></title>

</head>

<%
	if abbgimage="" then
		if abbgcolor <> "" then
			response.write "<body bgcolor=""" & abbgcolor & """>"
		end if
	else
		response.write "<body background=""" & abbgimage & """>"
	end if
	
	strheaderstring = getheaderstring(request.querystring("id"), 0)				
%>


<body bgcolor="#000000" link="#eef1f2" vlink="#e3e9ea" alink="#aeb6c1">


<h4><font face="arial" color="#e6e9ec"><%=strtitlelabel%><br>
<a href="aspboarddetail.asp?id=<%=request.querystring("id")%>">
<font color="#e6e9ec"><%=strheaderstring%></font></a></h4>
			
<%	
	dim adors
		
	strSql="SELECT * FROM msgDetail WHERE msgId = " & Request.QueryString("Id")		
	Set adoRs=Server.CreateObject("ADODB.Recordset")
	adoRS.Open strSql, adoConn, 1	
	
	If strHeaderString <> abDefaultHeader Then
		strHeaderString = "RE: " & strHeaderString
	End If
		
%>


<table width="80%" cellpadding="2" cellspacing="0" border="0" bgcolor="#000000" style="border-collapse: collapse">
	<form method="post" action="aspboardpost.asp" id="postForm" name="postForm" LANGUAGE="javascript" onsubmit="return Submit_onclick()">
		<tr>
			<td><font face="Arial"  COLOR="#e6e9ec" style="font-size: 8pt"><b>Your Name:</b></font></td>
			</font>
			<td><font COLOR="#e6e9ec" face="Arial" style="font-size: 8pt"><input id="author_nameStr" name="author_nameStr" type="text" size="24" Value="<%=Request.Cookies("postName")%>"></font></td>
		</tr>
		<font face="arial"  COLOR="#e6e9ec">
		<tr>
			<td><font face="Arial" COLOR="#e6e9ec" style="font-size: 8pt"><b>Your EMail:</b></font></td>
			<td><font COLOR="#e6e9ec" face="Arial" style="font-size: 8pt"><input id="author_emailStr" name="author_emailStr" type="text" size="24" Value="<%=Request.Cookies("postEmail")%>"></font></td>
		</tr>
		<tr>
			<td><font face="Arial" COLOR="#e6e9ec" style="font-size: 8pt"><b>Your URL:</b></font></td>
			<td><font COLOR="#e6e9ec" face="Arial" style="font-size: 8pt"><input id="author_urlStr" name="author_urlStr" type="text" size="24" Value="<%=Request.Cookies("postURL")%>"></font></td>
		</tr>
		<tr>
			<td><font face="Arial" COLOR="#e6e9ec" style="font-size: 8pt"><b>Subject:</b></font></td>
			<td><font COLOR="#e6e9ec" face="Arial" style="font-size: 8pt"><input id="headerStr" name="headerStr" type="text" size="48" value="<%=strHeaderString%>"></font></td>
		</tr>
		<tr>
			<td><b><font face="arial" size="-1"  COLOR="#e6e9ec"><%=strMsgLabel%></font><font color="#EEF1F2" face="Arial" style="font-size: 8pt">:</font></b></td>
			<td><font COLOR="#e6e9ec" face="Arial" style="font-size: 8pt"><textarea cols="48" rows="5" name="detailStr" id="detailStr"></textarea></font></td>
		</tr>
		<%If Request.QueryString("Id") > 0 Then%><font COLOR="#e6e9ec">
			<%If abIncOrigMsg = True Then%> </font>
				<tr><td colspan="2">&nbsp;</td></tr>
				<tr valign="top">
					<td>&nbsp;</td>
					<font face="Arial"  COLOR="#e6e9ec">
					<td></font><font COLOR="#e6e9ec">
                    <font face="Arial" COLOR="#e6e9ec" style="font-size: 8pt"><input type="checkbox" name="incOrig" Value="yes" CHECKED></font><font face="Arial" COLOR="#e6e9ec" style="font-size: 8pt"> <b>Include original message in response?<br></td>
				</tr>	
			<%End If%> </font></font>			
			<tr><td colspan="2">&nbsp;</td></tr>			
			<tr>
				<td>&nbsp;</td>
				<td><font face="Arial" COLOR="#e6e9ec" style="font-size: 8pt">Posted by <%=adoRs.Fields("author_nameStr").Value%>&nbsp;on&nbsp;<%=FormatDateTime(adoRs.Fields("msgTime").Value, 2)%>&nbsp;at&nbsp;<%=FormatDateTime(adoRs.Fields("msgTime").Value, 3)%></font></td>
			</tr>						
			<tr>			
				<td valign="top">
                <font face="Arial" COLOR="#e6e9ec" style="font-size: 8pt"><b>Original Message:</b></font></td><td>
                <font COLOR="#e6e9ec" face="Arial" style="font-size: 8pt"><textarea cols="48" rows="<%=GetRows(adoRs.Fields("detailStr").Value)%>" id="origMsg1" name="origMsg1" disabled><%=ReplaceQuotes(adoRs.Fields("detailStr").Value, 0)%></textarea></font></td>
			</tr>				
			<font COLOR="#e6e9ec">				
			<input type="hidden" id="origMsgStr" name="origMsgStr" value="<%=adoRs.Fields("detailStr").Value%>">
		</font>
		<%End If%><font COLOR="#e6e9ec">	
		<input type="hidden" id="msgId" name="msgId" value="<%=Request.QueryString("Id")%>">				
		</font><font face="Arial" COLOR="#e6e9ec">				
		<tr>			
			<td colspan="2" align="center">
				</font>
                <font face="Arial" COLOR="#e6e9ec" style="font-size: 8pt" COLOR="#e6e9ec">
				<input type="submit" value="Post Message" Name="Submit">
                <font face="Arial" COLOR="#e6e9ec">				
			</td>	
		</tr>
	</form>
</table>
</font><font face="arial" COLOR="#e6e9ec">
<br>
</center>
</BODY>
</HTML>

<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function Submit_onclick() {
	//check for client side (form) validation
	
	if (Form_Validate() == true) {							
		return true;
	} else {
		return false;
	}	

}

//-->
</SCRIPT>

<SCRIPT Language="JavaScript">

function Form_Validate() {

	if (document.postForm.author_nameStr.value==""){
		alert("[Your Name] cannot be blank");
		return false;
	}
	
	if (document.postForm.author_emailStr.value==""){
		alert("[Your Email] cannot be blank");
		return false;
	}
	
	if (document.postForm.detailStr.value==""){
		alert("[Your Response] cannot be blank");
		return false;
	}
	
	return true;
}	
</SCRIPT></font>