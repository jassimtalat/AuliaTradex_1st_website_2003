<%@ language=vbscript %>
<%option explicit%>
<%response.buffer=true%>
<!--#include file="aspboardfunctions.asp"-->
<%	
	
	dim adors	
	
	'uncomment these lines if you want to add the ability to 
	'delete messages. make sure you set a value for abadminpwd (see aspboarddata.asp):
	'if request.querystring("delid")<>"" and request.querystring("pwd") = abadminpwd then
	'	strsql = "delete * from msgdetail where msgid = " & request.querystring("delid")
	'	call adoconn.execute(strsql)
	'	response.redirect "aspboard.asp"
	'end if	
	
	strsql="select * from msgdetail where msgid = " & request.querystring("id")
		
	set adors=server.createobject("adodb.recordset")
	adors.open strsql, adoconn, 1
	
	if adors.eof then
		response.redirect "aspboard.asp"
	end if	
%>

<html>
<head>
<meta name="keywords">
<meta name="description">
<meta name="author">
<meta name="robots" content="index,follow">
<title>aulia tradex (<%=adors.fields("headerstr").value%>)</title>
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


<body text="#eceef0" bgcolor="#000000">


<h3><font face="arial" color="#e6e9ec"><%=adors.fields("headerstr").value%></font></h3>
<table width="80%" cellpadding="2" cellspacing="0" border="0" bgcolor="#000000" style="border-collapse: collapse">
	<tr><td><font face="arial" size="-1" color="#eceef0">posted by: <%=adors.fields("author_namestr").value%></font></td></tr>	
	<tr><td><font face="arial" size="-1" color="#eceef0">when: <%=formatdatetime(adors.fields("msgtime").value, 2)%>&nbsp;<%=formatdatetime(adors.fields("msgtime").value, 3)%></font></td></tr>	
	<%if not trim(adors.fields("author_emailstr").value) = "" then%>
		<tr><td><font face="arial" color="#eceef0" size="-1">email: </font><font face="arial" color="#e6e9ec" size="-1"> <a href="mailto:<%=adors.fields("author_emailstr").value%>">
          <font color="#eceef0"><%=adors.fields("author_emailstr").value%></font></a></font></td></tr>
	<%end if%>
	<%if not trim(adors.fields("author_urlstr").value) = "" then%>
		<tr><td><font face="arial" size="-1"><font color="#eceef0">url: </font> <a href="<%=adors.fields("author_urlstr").value%>">
          <font color="#eceef0"><%=adors.fields("author_urlstr").value%></font></a></font></td></tr>
	<%end if%>

	<tr>
		<td><textarea cols="64" rows="<%=getrows(adors.fields("detailstr").value)%>"><%=replacequotes(adors.fields("detailstr").value, 0)%></textarea></td>
	</tr>
			
	<%if adors.fields("parentid").value <> 0 then%>
		<tr><td>	
		<h3><font face="arial" color="#eceef0">previous thread:</font></h3>
		<%=getheaderstring(adors.fields("parentid").value, -1)%><font color="#eceef0">
        </font>
		</td></tr>
	<%end if%>
	
	<%if existfollowups(request.querystring("id")) then%>
		<tr><td>
		<h3><font face="arial" color="#eceef0">follow-ups:</font></h3>
		<%listitems(request.querystring("id"))%><font color="#eceef0"> </font>
		</td></tr>
	<%end if%>

</table>

<%	
	set adors = nothing
	set adoconn = nothing
%>		

<br><font face="arial" size="2"><a target="_self" href="message.html"><font color="#eceef0">

Previous</font></a></font><br>
</font>
<font size="2" face="arial"><a href="tellafriend1.asp">reply a message</a></font><br>
</body>
</html>