<% @language="vbscript" %>
<% Option Explicit %>  
<!--#include virtual="/_private/login.inc"-->

<%
  
	' Define your variables.
	Dim mbjFSO,mInStream,mRows,msArrRows  
	Dim mFileName, s
	Dim mReturnValue
	Dim contact_firstname, contact_lastname, contact_credential, contact_credential_no
	Dim dUser, dPwd, tUser, tPwd
	
	mFileName = "/_private/registration_results.csv"	
	
	'*** Create Object ***'  
	Set mbjFSO = CreateObject("Scripting.FileSystemObject")
	
	'*** Check Exist Files ***'  
	
	If Not mbjFSO.FileExists(Server.MapPath(mFileName)) Then  
		Response.write("File not found.")  

	Else 
		 '*** Open Files ***'  
		Set mInStream = mbjFSO.OpenTextFile(Server.MapPath(mFileName),1,False) 
		
		Do Until mInStream.AtEndOfStream  
			mRows = mInStream.readLine  
			msArrRows = Split(mRows,",")
			dUser = msArrRows(14)
			dPwd = msArrRows(15)
			tUser = """" + Session("USER") + """"
			tPwd = """" + Session("PWD") + """"
			If (dUser = tUser) And (dPwd = tPwd) Then
				contact_firstname = Replace(msArrRows(0), Chr(34), "")
				contact_lastname = Replace(msArrRows(1), Chr(34), "")
				contact_credential = Replace(msArrRows(11), Chr(34), "")
				contact_credential_no = Replace(msArrRows(12), Chr(34), "")
			End If

		Loop
			
		mInStream.Close()  
		Set mInStream = Nothing 

	End IF 
	
%>

<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns="http://www.w3.org/TR/REC-html40">

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>MSW Cert</title>
<style>
<!--
h2
	{margin-bottom:.0001pt;
	text-align:center;
	page-break-after:avoid;
	font-size:45.0pt;
	font-family:"Times New Roman";
	letter-spacing:3.0pt;
	font-weight:normal; margin-left:0in; margin-right:0in; margin-top:0in}
-->
</style>
<!--[if !mso]>
<style> table.MsoNormalTable
	{mso-style-parent:"";
	font-size:10.0pt;
	font-family:"Times New Roman",serif}
</style>
<![endif]-->
</head>

<body>

<table border="10" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" bordercolorlight="#000000" bordercolordark="#000000">
  <tr>
    <td width="100%">
    <h2><img border="0" src="/Web%20Folder%20Sample/images/logo%20complete(large).jpg" width="110" height="62"></h2>
	<h2><span style="font-size:36.0pt">Certificate of Credit</span></h2>
	<p class="MsoNormal" align="center" style="text-align: center">
	<span style="font-size:13.5pt;letter-spacing:1.0pt">This is to certify that:</span></p>
	
    <p class="MsoNormal" align="center"><font size="5">
    <%=contact_firstname%>&nbsp;&nbsp;
    <%=contact_lastname%>,&nbsp;&nbsp;
    <%=contact_credential%> &nbsp;&nbsp; License No.
    <%=contact_credential_no%></font></p>

    
    <p class="MsoNormal" align="center" style="text-align: center">
	<span style="font-size:18.0pt">&nbsp;</span><span style="font-size:13.5pt;letter-spacing:1.0pt">&nbsp;&nbsp;</span><span style="letter-spacing:1.0pt">Has 
	successfully completed the following self-paced, reading-based, </span><span style="letter-spacing:1.0pt">
	self-study </span><span style="letter-spacing:1.0pt">online course</span></p>
	<p class="MsoNormal" align="center" style="text-align: center; line-height: 150%">
	<span style="font-size:20.0pt;
line-height:150%">&nbsp;&nbsp;<span style="letter-spacing:1.0pt">“</span><b>Achieving 
	Positive Outcomes in the Non-Adherent Diabetes Patient</b>”</span><span style="line-height: 150%; color: black; letter-spacing: 1.0pt"><b> </b></span></p>
	<p class="MsoNormal" align="center" style="text-align: center; line-height: 150%">
	<span style="letter-spacing: 1pt"><b><font size="5">1.0 CE Credits</font></b></span><span style="line-height: 150%; color: black; letter-spacing: 1.0pt"><b><font size="5">&nbsp;
	</font> </b></span></p>
	<p class="MsoNormal" align="center" style="text-align: center">
	<span style="font-size:13.5pt;letter-spacing:1.0pt">&nbsp;&nbsp;</span><span style="font-size:16.0pt;letter-spacing:1.0pt">on 
	this day
	<!--webbot bot="Timestamp" S-Type="REGENERATED" S-Format="%B %d, %Y" startspan -->October 15, 2019<!--webbot bot="Timestamp" i-checksum="31176" endspan --></span></p>
	<font SIZE="3">
	<p align="center">Hi-R-Ed Online University, provider #1091, is approved to 
	offer social work continuing education by the Association of Social Work 
	Boards (ASWB) Approved Continuing Education (ACE) program. Organizations, 
	not individual courses, are approved as ACE providers. State and provincial 
	regulatory boards have the final authority to determine whether an 
	individual course may be accepted for continuing education credit. Hi-R-Ed 
	Online University maintains responsibility for this course. ACE provider 
	approval period:
	<span style="font-size: 12.0pt; font-family: 'Times New Roman',serif">
	2/28/19 - 2/28/22</span>. Social workers completing this course receive 1.0 
	continuing education credits. </p>
	</font>
	<p class="MsoNormal" align="center" style="text-align: center; margin-top: 0; margin-bottom: 0">
	<span style="font-size:18.0pt;
  font-family:&quot;Script MT Bold&quot;">&nbsp;&nbsp; <u>
			Sheila Miles, RN</u></span><span style="font-size:10.0pt">&nbsp; 
			Continuing Education Director&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; </span></p>
	<table class="MsoNormalTable" border="0" cellspacing="0" cellpadding="0" width="962" style="width:721.55pt;border-collapse:collapse">
		<tr style="height: 13.95pt">
			<td width="100%" style="width:100.0%;padding:.75pt .75pt .75pt .75pt;
  height:13.95pt">
			<p align="center" style="text-align:center">
			&nbsp;</td>
		</tr>
	</table>
    <p class="MsoNormal">&nbsp;</td>
  </tr>
</table>

<p class="MsoNormal">&nbsp;</p>

</body>

</html>