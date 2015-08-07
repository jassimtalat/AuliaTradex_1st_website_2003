<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<HTML><HEAD><TITLE></TITLE>

<META NAME="author" content="">
<script language="JavaScript">
<!--
function MM_findObj(n, d) { //v3.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document); return x;
}

function MM_validateForm() { //v3.0
  var i,p,q,nm,test,num,min,max,errors='',args=MM_validateForm.arguments;
  for (i=0; i<(args.length-2); i+=3) { test=args[i+2]; val=MM_findObj(args[i]);
    if (val) { nm=val.name; if ((val=val.value)!="") {
      if (test.indexOf('isEmail')!=-1) { p=val.indexOf('@');
        if (p<1 || p==(val.length-1)) errors+='- '+nm+' must contain an e-mail address.\n';
      } else if (test!='R') { num = parseFloat(val);
        if (val!=''+num) errors+='- '+nm+' must contain a number.\n';
        if (test.indexOf('inRange') != -1) { p=test.indexOf(':');
          min=test.substring(8,p); max=test.substring(p+1);
          if (num<min || max<num) errors+='- '+nm+' must contain a number between '+min+' and '+max+'.\n';
    } } } else if (test.charAt(0) == 'R') errors += '- '+nm+' is required.\n'; }
  } if (errors) alert('The following error(s) occurred:\n'+errors);
  document.MM_returnValue = (errors == '');
}
//-->
</script>
</HEAD>
<body bgcolor="000000" topmargin=0 leftmargin=0 marginheight=0 marginwidth=0 text="DEDBC3" link="ffffdf" vlink="#FBF6C4">


<%
On Error Resume Next
'------------------Modify this section to customize your message
strMsg="Dear " & request.form("Yourfriendname") & "," & vbcrlf & vbcrlf 
strMsg=strMsg & request.form("yourname") 
strMsg=strMsg & " wants you to check out VisualBasicScript.Com's page at" & vbcrlf & vbcrlf
strMsg=strMsg & "http://www.VisualBasicScript.Com" & vbcrlf & vbcrlf
strMsg=strMsg & request.form("yourname") & " also says" & vbcrlf & vbcrlf
strMsg=strMsg & request.form("message") & vbcrlf & vbcrlf
strMsg=strMsg & "If you do not know " & request.form("yourname") 
strMsg=strMsg & ", please ignore this message or report it to info@rapidsolutionz.com"  
strMsg=strMsg & vbcrlf & vbcrlf
strMsg=strMsg & "Warmest Regards," & vbcrlf
strMsg=strMsg & "Rapid Solutionz Staff" & vbcrlf
strMsg=Cstr(strMsg)

Set objMail = CreateObject("CDONTS.NewMail")
      objMail.From= request.form("youremail") 'Specify sender's address
      objMail.To=request.form("Yourfriendemail")
      objMail.Subject="" ' Subject of the message
      objMail.Body=strMsg
      objMail.Send
Set objMail = nothing

%> 
<p>
<font color="#ECEEF0"><font face="Arial" size="2">
<br>
</font>
<p>
<b><font face="Arial" size="2">Reply message? </font></b></font>
<center>
<p>
&nbsp;<p>

    
<div align="center">
  <center>

    
<TABLE CELLPADDING=0 CELLSPACING=0 BORDER=0 width="75%" bgcolor="#000000" style="border-collapse: collapse">
<FORM METHOD="POST" ACTION="tellafriend1.asp"> 
  <TR> 
        
    <TD ALIGN=RIGHT width="187"><font color="#ECEEF0" face="Arial" size="2">&nbsp;*Your name:
    </font> </TD>
        
    <TD width="115"> 
          <font face="Arial" color="#ECEEF0"> 
          <INPUT TYPE="text" NAME="yourname" VALUE="<%=request.form("yourname")%>" SIZE=40 MAXLENGTH=70><font size="2">
          </font></font>
        </TD>
      </TR>
      <TR> 
        
    <TD ALIGN=RIGHT width="187"><font color="#ECEEF0" face="Arial" size="2">*Your e-mail address:
    </font> </TD>
        
    <TD width="115"> 
          <font face="Arial" color="#ECEEF0"> 
          <INPUT TYPE="text" NAME="youremail" VALUE="<%=request.form("youremail")%>" SIZE=40 MAXLENGTH=70><font size="2">
          </font></font>
        </TD>
      </TR>
      <TR> 
        
    <TD ALIGN=RIGHT width="187"><font color="#ECEEF0" face="Arial" size="2">*Your friend's name:
    </font> </TD>
        
    <TD width="115"> 
          <font face="Arial" color="#ECEEF0"> 
          <INPUT TYPE="text" NAME="Yourfriendname" VALUE="" SIZE=40 MAXLENGTH=70><font size="2">
          </font></font>
        </TD>
      </TR>
      <TR> 
        
    <TD ALIGN=RIGHT width="187"><font color="#ECEEF0" face="Arial" size="2">*Your friend's e-mail address:
    </font> </TD>
        
    <TD width="115"> 
          <font face="Arial" color="#ECEEF0"> 
          <INPUT TYPE="text" NAME="Yourfriendemail" VALUE="" SIZE=40 MAXLENGTH=70><font size="2">
          </font></font>
        </TD>
      </TR>
      <TR> 
        
    <TD ALIGN=RIGHT width="187"><font color="#ECEEF0" face="Arial" size="2">Your message:
    </font> </TD>
        
    <TD width="115"> 
          <font face="Arial" color="#ECEEF0"> 
          <TEXTAREA NAME='message' ROWS=5 COLS=40 WRAP=SOFT><%=request.form("message")%></TEXTAREA><font size="2">
          </font></font>
        </TD>
      </TR>
      <TR> 
        
    <TD ALIGN=RIGHT width="187"> </TD>
        
    <TD width="115"> 
        <font face="Arial" color="#ECEEF0"> 
        <INPUT TYPE="submit" NAME="Send Comments" VALUE="Send Comments" onClick="MM_validateForm('yourname','','R','youremail','','RisEmail','Yourfriendname','','R','Yourfriendemail','','RisEmail');return document.MM_returnValue"><input type="submit" name="Close" value="Close" onClick="javascript:self.close()"><font size="2">
        <a target="_self" href="read.asp">Back</a>&nbsp; </font></font>
      
      <td width="187"></FORM>
  
    </TABLE>


  </center>
</div>


</center> 
</body>
</html>