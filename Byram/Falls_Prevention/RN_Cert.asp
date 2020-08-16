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
<link rel="File-List" href="../CAUTI/RN_Cert_files/filelist.xml">
<title>RN Cert</title>
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
<style>
v\:*         { behavior: url(#default#VML) }
o\:*         { behavior: url(#default#VML) }
.shape       { behavior: url(#default#VML) }
</style>
<![endif]--><!--[if gte mso 9]>
<xml><o:shapedefaults v:ext="edit" spidmax="1027"/>
</xml><![endif]-->
</head>

<body>

<table border="10" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" bordercolorlight="#C0C0C0" bordercolordark="#9999FF">
  <tr>
    <td width="100%">
    <p class="MsoNormal">&nbsp;</p>
    <p class="MsoNormal" align="center">&nbsp;<img border="0" src="../IVAbx/images/Optioncare%20logo(small).gif" width="273" height="76">&nbsp;</p>
    <h2><font size="7">Certificate of Completion</font></h2>
    <p class="MsoNormal" align="center" style="text-align:center"><!--[if gte vml 1]><v:shapetype id="_x0000_t75"
 coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe"
 filled="f" stroked="f">
 <v:stroke joinstyle="miter"/>
 <v:formulas>
  <v:f eqn="if lineDrawn pixelLineWidth 0"/>
  <v:f eqn="sum @0 1 0"/>
  <v:f eqn="sum 0 0 @1"/>
  <v:f eqn="prod @2 1 2"/>
  <v:f eqn="prod @3 21600 pixelWidth"/>
  <v:f eqn="prod @3 21600 pixelHeight"/>
  <v:f eqn="sum @0 0 1"/>
  <v:f eqn="prod @6 1 2"/>
  <v:f eqn="prod @7 21600 pixelWidth"/>
  <v:f eqn="sum @8 21600 0"/>
  <v:f eqn="prod @7 21600 pixelHeight"/>
  <v:f eqn="sum @10 21600 0"/>
 </v:formulas>
 <v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/>
 <o:lock v:ext="edit" aspectratio="t"/>
</v:shapetype><v:shape id="_x0000_s1025" type="#_x0000_t75" style='position:absolute;
 margin-left:504.75pt;margin-top:6pt;width:115.5pt;height:189pt;z-index:-2;
 mso-wrap-edited:f' wrapcoords="-57 0 -57 21565 21600 21565 21600 0 -57 0"
 filled="t" fillcolor="yellow" stroked="t" strokecolor="#cff">
 <v:imagedata src="RN_Cert.asp" o:title="CDR 96-15" gain="55050f" blacklevel="2621f"/>
</v:shape><![endif]--></p>
    <p class="MsoNormal" align="center" style="text-align:center">
    <span style="font-family: Times New Roman; letter-spacing: 1.0pt">
    <font size="4">This certificate is awarded to</font></span></p>
    
    <p class="MsoNormal" align="center" style="text-align:center"><u>
    <%=contact_firstname%>&nbsp;&nbsp;
    <%=contact_lastname%>,&nbsp;&nbsp;
    <%=contact_credential%></u></p>
    
    <p class="MsoNormal" align="center" style="text-align:center">License No.
    <%=contact_credential_no%></p>
    
    <p class="MsoNormal" align="center">
    <span style="font-size: 18.0pt; font-family: Times New Roman">&nbsp;</span><span style="font-family: Times New Roman; letter-spacing: 1.0pt"><font size="4">For completion of the </font></span>
    <span style="font-family: Times New Roman; letter-spacing: 1pt">
    <font size="4">course</font></span></p>
    <p class="MsoNormal" align="center" style="text-align:center;line-height:150%">
    <b>
    <span style="font-size: 16.0pt; font-family: Times New Roman; letter-spacing: 1.0pt">
    “IV Antibiotics: Focus on Cellulitis” </span></b></p>
    <p class="MsoNormal" align="center" style="text-align:center;line-height:150%">
    <span style="font-size: 16.0pt; font-family: Times New Roman; letter-spacing: 1.0pt">
    Provided by <img border="0" src="../IVAbx/images/logo%203cert.jpg" width="83" height="42"></span></p>
    <p class="MsoNormal" align="center" style="text-align:center;line-height:150%">
    <span style="font-size: 16.0pt; font-family: Times New Roman; letter-spacing: 1.0pt">
    and is awarded </span><span style="font-size: 16pt; letter-spacing: 1pt"><b>
    <i><span style="font-family: Times New Roman">One</span></i></b></span><b><i><span style="font-size: 16.0pt; font-family: Times New Roman; letter-spacing: 1.0pt">
    Hour<span style="color:red"> </span>(1.0)</span></i></b><span style="font-size: 16.0pt; font-family: Times New Roman; letter-spacing: 1.0pt">of 
    continuing professional education.</span></p>
    <!--
    Document.Write Date()
    
    -->
    <div align="center">
      <center>
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="90%">
        <tr>
          <td width="100%"><font size="1">All courses, artwork, designs and 
          content is the property of Option Care and/or Hi-R-Ed Online University and is protected 
          by intellectual property laws of the United States and the state of 
          Illinois. Any duplication, modification or transfer of intellectual 
          property is prohibited and must be authorized by the above-named 
          parties.</font></td>
        </tr>
      </table>
      </center>
    </div>
    <p class="MsoNormal">&nbsp;</td>
  </tr>
</table>

<p class="MsoNormal">click here to return to <a href="../default.htm">Hi-R-Ed 
home</a></p>

</body>

</html>