

<body text="#ECEEF0" bgcolor="#000000">


<FORM ACTION="login.asp" METHOD="post">
<div align="center" style="width: 682; height: 215">
  <p align="left"><font face="Fixedsys" color="#C0C0C0">Administrator Login :</font><br>
&nbsp;</p>
  <center>
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="273" bordercolorlight="#ECEEF0" bordercolordark="#40494A" bordercolor="#C0C0C0" height="19">
    <tr>
      <td width="206" bgcolor="#FFCC00" bordercolorlight="#FFCC00" bordercolor="#FFCC00" height="19">
      <font color="#000000">
      <font face="verdana" style="font-size: 8pt; font-weight: 700">
      <img border="0" src="images/key2.gif"> </font><sup>
      <font face="verdana" style="font-weight: 700" size="2">AULIA TRADEX</font><font face="Arial" style="font-size: 7pt">®</font><font face="Verdana"><span style="font-size: 10pt">
      </span><span style="font-size: 8pt">Admin Login</span></font></sup></font></td>
      <td width="67" bgcolor="#FFCC00" align="right" height="19" valign="top">
      <font style="font-size: 8pt" color="#40494A" face="Webdings"><b>0</b>2</font><a title="Close" target="_self" href="statistics.html" style="text-decoration: none"><font color="#000000" style="font-size: 8pt" face="Webdings">r</font></a></td>
    </tr>
  </table>
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="272" bordercolorlight="#FDFDFD" bordercolordark="#C0C0C0" height="90" bgcolor="#000000">
    <tr>
		<TD ALIGN="right" height="22" bordercolor="#D6D3CE" bgcolor="#D6D3CE" width="80" valign="top">
        <font color="#000000" face="Verdana" style="font-size: 8pt"><br>
        Login: <br>
&nbsp;</font></TD>
		<TD height="22" bordercolor="#D6D3CE" width="189" bgcolor="#D6D3CE"></INPUT>
        <font face="Verdana" color="#D6D3CE">
        <INPUT TYPE="text" NAME="login" size="20"></font></TD>
	</tr>
    <tr>
		<TD ALIGN="right" height="22" bordercolor="#D6D3CE" bgcolor="#D6D3CE" width="80">
        <font color="#000000" face="Verdana" style="font-size: 8pt">Password:</font></TD>
		<TD height="22" bordercolor="#D6D3CE" width="189" bgcolor="#D6D3CE">
        <font color="#D6D3CE"><INPUT TYPE="password" NAME="password" size="20"></font></INPUT></TD>
	</tr>
    <tr>
		<TD ALIGN="right" height="63" bordercolor="#D6D3CE" bgcolor="#D6D3CE" width="80">&nbsp;</TD>
		<TD height="62" bordercolor="#D6D3CE" bgcolor="#D6D3CE" width="189"><INPUT TYPE="submit" VALUE="Login"></INPUT>
			<INPUT TYPE="reset" VALUE="Reset"></TD>
	</tr>
  </table>
  <p><font face="Arial" size="2"><br>
&nbsp;</font></p>
  </center>
</div>
</FORM>
<%

If lcase(Request.Form("login")) = "admin" AND Request.Form("password") = "goldpaper" Then
	Response.Redirect ("underconstruction.html") 
	
End If
%>