<HTML xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns="http://www.w3.org/TR/REC-html40">
<HEAD>
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">

<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=windows-1252">
<meta http-equiv="Content-Language" content="en-us">

<TITLE>registration form</TITLE>
<style fprolloverstyle>A:hover {color: #800080; font-weight: bold}
</style>
</HEAD>
<BODY bgcolor="#C0C0C0">
<H1><font size="4" color="#FFFFFF">Certificate Information Form&nbsp;&nbsp;
</font>
<span style="font-family: Times New Roman">
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span style="font-size: 12.0pt; font-family: Times New Roman">&nbsp;<b> </b></span>
<font color="#ffffff">&nbsp; </font></H1>
<P>
<img border="0" src="Optioncare/images/Windows%20line.jpg" width="100%" height="13"></P>
<P>
&nbsp;</P>
<!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Language="JavaScript" Type="text/javascript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.Contact_FirstName.value == "")
  {
    alert("Please enter a value for the \"First Name\" field.");
    theForm.Contact_FirstName.focus();
    return (false);
  }

  if (theForm.Contact_FirstName.value.length < 1)
  {
    alert("Please enter at least 1 characters in the \"First Name\" field.");
    theForm.Contact_FirstName.focus();
    return (false);
  }

  var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz����������������������������������������������������������������������- \t\r\n\f";
  var checkStr = theForm.Contact_FirstName.value;
  var allValid = true;
  var validGroups = true;
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
  }
  if (!allValid)
  {
    alert("Please enter only letter, whitespace and \"-\" characters in the \"First Name\" field.");
    theForm.Contact_FirstName.focus();
    return (false);
  }

  if (theForm.Contact_LastName.value == "")
  {
    alert("Please enter a value for the \"Last Name\" field.");
    theForm.Contact_LastName.focus();
    return (false);
  }

  if (theForm.Contact_LastName.value.length < 2)
  {
    alert("Please enter at least 2 characters in the \"Last Name\" field.");
    theForm.Contact_LastName.focus();
    return (false);
  }

  var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz����������������������������������������������������������������������- \t\r\n\f";
  var checkStr = theForm.Contact_LastName.value;
  var allValid = true;
  var validGroups = true;
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
  }
  if (!allValid)
  {
    alert("Please enter only letter, whitespace and \"-\" characters in the \"Last Name\" field.");
    theForm.Contact_LastName.focus();
    return (false);
  }

  if (theForm.Contact_Credentials.value == "")
  {
    alert("Please enter a value for the \"Contact_Credentials\" field.");
    theForm.Contact_Credentials.focus();
    return (false);
  }

  var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz����������������������������������������������������������������������,. \t\r\n\f";
  var checkStr = theForm.Contact_Credentials.value;
  var allValid = true;
  var validGroups = true;
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
  }
  if (!allValid)
  {
    alert("Please enter only letter, whitespace and \",.\" characters in the \"Contact_Credentials\" field.");
    theForm.Contact_Credentials.focus();
    return (false);
  }

  if (theForm.Contact_LicenseNo.value == "")
  {
    alert("Please enter a value for the \"Contact_LicenseNo\" field.");
    theForm.Contact_LicenseNo.focus();
    return (false);
  }

  if (theForm.Contact_LicenseNo.value.length < 4)
  {
    alert("Please enter at least 4 characters in the \"Contact_LicenseNo\" field.");
    theForm.Contact_LicenseNo.focus();
    return (false);
  }

  if (theForm.Contact_LicenseNo.value.length > 20)
  {
    alert("Please enter at most 20 characters in the \"Contact_LicenseNo\" field.");
    theForm.Contact_LicenseNo.focus();
    return (false);
  }

  var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz����������������������������������������������������������������������0123456789--,/ \t\r\n\f";
  var checkStr = theForm.Contact_LicenseNo.value;
  var allValid = true;
  var validGroups = true;
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
  }
  if (!allValid)
  {
    alert("Please enter only letter, digit, whitespace and \"-,/\" characters in the \"Contact_LicenseNo\" field.");
    theForm.Contact_LicenseNo.focus();
    return (false);
  }

  if (theForm.Contact_Email.value == "")
  {
    alert("Please enter a value for the \"Contact_Email\" field.");
    theForm.Contact_Email.focus();
    return (false);
  }

  var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz����������������������������������������������������������������������0123456789-.,/@_";
  var checkStr = theForm.Contact_Email.value;
  var allValid = true;
  var validGroups = true;
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
  }
  if (!allValid)
  {
    alert("Please enter only letter, digit and \".,/@_\" characters in the \"Contact_Email\" field.");
    theForm.Contact_Email.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><FORM METHOD="POST" ACTION="CertInfoForm-CCM.asp" name="FrontPage_Form1" onSubmit="return FrontPage_Form1_Validator(this)" language="JavaScript" webbot-action="--WEBBOT-SELF--">
<!--webbot bot="SaveResults" s-builtin-fields="REMOTE_NAME REMOTE_USER HTTP_USER_AGENT Date Time" startspan U-Confirmation-Url="Webinar032515/CCM_Cert.asp" U-File="Webinar032515/CertInfoForm-CCM1.mdb" S-Format="HTML/BR" S-Label-Fields="TRUE" B-Reverse-Chronology="FALSE" S-Date-Format="%m/%d/%Y" S-Time-Format="%I:%M %p" --><input TYPE="hidden" NAME="VTI-GROUP" VALUE="0"><!--webbot bot="SaveResults" endspan i-checksum="43374" --><P>
<b>
<font color="#FFFFFF" face="Calibri">Please enter the following information as you would like 
it to appear on your certificate of completion.</font></b></P>
<BLOCKQUOTE>
<TABLE style="border-collapse: collapse" bordercolor="#111111" cellpadding="0" cellspacing="0" bgcolor="#527AB8" width="475">
<TR>
<TD ALIGN="right" width="180" height="23">
<b>
<font color="#FFFFFF">
<EM><a name="First Name">First Name</a> &nbsp; </EM></font></b></TD>
<TD width="295" height="23">
<!--webbot bot="Validation" s-display-name="First Name" s-data-type="String" b-allow-letters="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="-" b-value-required="TRUE" i-minimum-length="1" --><INPUT NAME="Contact_FirstName" SIZE=25 tabindex="1">
</TD>
</TR>
<TR>
<TD ALIGN="right" width="180" height="23">
<b>
<font color="#FFFFFF">
<EM>Last Name&nbsp;&nbsp; </EM></font></b></TD>
<TD width="295" height="23">
<!--webbot bot="Validation" s-display-name="Last Name" s-data-type="String" b-allow-letters="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="-" b-value-required="TRUE" i-minimum-length="2" --><INPUT NAME="Contact_LastName" SIZE=25 tabindex="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</TD>
</TR>
<TR>
<TD ALIGN="right" width="180" height="23">
<b>
<font color="#FFFFFF">
<em>Credentials&nbsp;&nbsp; </em></font></b></TD>
<TD width="295" height="23">
<!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars=",." b-value-required="TRUE" --><INPUT NAME="Contact_Credentials" SIZE=25 tabindex="3">
</TD>
</TR>
<TR>
<TD ALIGN="right" width="180" height="23">
<b>
<font color="#FFFFFF">
<em>Credential/License No&nbsp;&nbsp; </em></font></b></TD>
<TD width="295" height="23">
<!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="-,/" b-value-required="TRUE" i-minimum-length="4" i-maximum-length="20" --><INPUT NAME="Contact_LicenseNo" SIZE=25 tabindex="4" maxlength="20">
</TD>
</TR>
<tr>
<TD ALIGN="right" width="180" height="23">
<i><b><font color="#FFFFFF">State of Issue&nbsp;&nbsp; </font></b></i></TD>
<TD width="295" height="23">
<input type="text" name="T1" size="25" tabindex="5"></TD>
</tr>
<tr>
<TD ALIGN="right" width="180" height="23" valign="top">
<b>
<font color="#FFFFFF">
<EM>E-mail&nbsp;&nbsp; </EM></font></b></TD>
<TD width="295" height="23">
<!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" s-allow-other-chars=".,/@_" b-value-required="TRUE" --><INPUT NAME="Contact_Email" SIZE=25 tabindex="7">
<p><b><font color="#FF0000" size="2">email must be correct to receive CE 
certificate</font></b></TD>
</tr>
</TABLE>
</BLOCKQUOTE>
<INPUT TYPE=submit VALUE="Submit Form" tabindex="20">&nbsp;
</FORM>
<HR color="#527AB8">
<H5>
<span style="font-weight: 400"><font face="Calibri" size="4">If you have 
difficulty submitting this form, wait a few minutes and try again.</font></span></H5>
<H5>
<img border="0" src="Optioncare/logo3(icon).jpg" width="43" height="23"> <span style="font-weight: 400">
<font size="1">RFG Hi-<i>R</i>-Ed2.<br>Copyright � 2014 Hi-<i>R</i>-Ed.org. All rights reserved. <BR>
Revised: 
<!--WEBBOT BOT=TimeStamp
    S-Type="EDITED"
    S-Format="%m/%d/%y" startspan
-->03/21/15<!--webbot bot="TimeStamp" i-checksum="12946" endspan --></font></span></H5>
</BODY>
</HTML>