<% @language="vbscript" %>
<!--#include virtual="_private/login.inc"-->
<%
  ' Was this page posted to?
  If UCase(Request.ServerVariables("HTTP_METHOD")) = "POST" Then
	
    ' If so, check the username/password that was entered.
    If ComparePassword(Request("USER"),Request("PWD")) Then
    
      'If comparison was good, store the user name...
      Session("USER")= Request("USER")
	  Session("PWD")= Request("PWD")
      
      ' ...and redirect back to the original page.
      Response.Redirect "/sponsors(new).asp"
    End If
  End If
%>
<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns="http://www.w3.org/TR/REC-html40">

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta http-equiv="Content-Language" content="en-us">
<link rel="File-List" href="Login_files/filelist.xml">

<title>Registration Form</title>
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

<body bgcolor="#FFFFFF">

<p>
<!--webbot bot="PurpleText" preview="Feedback Form - Customize the form below to collect the information you need. By default, the form data is saved to a text file on the web server using the FrontPage Save Results component. Edit the Form Properties to change this behavior." -->
</p>

<font color="#FFFFFF"><b><i>
<a name="New Users:4">
<div align="center">
			<table border="0" width="847" style="border-collapse: collapse">
				<tr>
					<td bgcolor="#FFFFFF">&nbsp;</td>
					<td width="186" bgcolor="#FFFFFF">
					<p align="center"><b><font face="Calibri" color="#333333">
					Hi-R-Ed Online</font></b></td>
				</tr>
				<tr>
					<td bgcolor="#044996" height="5">&nbsp;</td>
					<td width="186" bgcolor="#044996" height="5">&nbsp;</td>
				</tr>
				<tr>
					<td bgcolor="#FFFFFF" height="71">
					<div align="center">
					<table border="0" width="100%" style="border-collapse: collapse">
						<tr>
							<td>
							<p align="center"><b>
							<font face="Calibri" color="#333333" style="font-size: 16pt">
							Your Source for Professional Continuing Education 
							and Training</font></b></td>
						</tr>
						<tr>
							<td>
							<p align="center"><i><b>
							<font face="Calibri" color="#333333">Where All 
							Courses Are Fully Accredited and Free!</font></b></i></td>
						</tr>
					</table>
					</div>
					</td>
					<td width="186" align="center" bgcolor="#FFFFFF" height="71">
					<img border="0" src="Web%20Folder%20Sample/images/logo%20complete(large).jpg" width="110" height="62"></td>
				</tr>
			</table>
			</div>

</i></b></font>

<a name="New Users:"><b>
<div align="center">
	<table border="0" width="850" style="border-collapse: collapse">
		<tr>
			<td>
<a name="New Users:3"><b>
			<font size="4" color="#FFFFFF">
			&nbsp;</font><font face="Calibri" size="5" color="#044996">Sign In Form</font></td>
		</tr>
		</table>
</div>
<div align="center">
<table border="0" width="850" id="table2" style="border-collapse: collapse" bgcolor="#EFEFEF">
	<tr>
		<td width="72"><!--[if gte vml 1]><v:line id="_x0000_s1027"
 style='position:absolute;left:0;text-align:left;top:0;flip:x;z-index:1'
 from="52.5pt,129.75pt" to="54.75pt,531pt" strokecolor="white"/><![endif]--><![if !vml]><span
style='mso-ignore:vglayout;position:absolute;z-index:1;left:69px;top:172px;
width:5px;height:537px'><img width=5 height=537 src="Login_files/image001.gif"
v:shapes="_x0000_s1027"></span><![endif]>&nbsp;&nbsp; <!--[if gte vml 1]><v:shapetype id="_x0000_t136"
 coordsize="21600,21600" o:spt="136" adj="10800" path="m@7,l@8,m@5,21600l@6,21600e">
 <v:formulas>
  <v:f eqn="sum #0 0 10800"/>
  <v:f eqn="prod #0 2 1"/>
  <v:f eqn="sum 21600 0 @1"/>
  <v:f eqn="sum 0 0 @2"/>
  <v:f eqn="sum 21600 0 @3"/>
  <v:f eqn="if @0 @3 0"/>
  <v:f eqn="if @0 21600 @1"/>
  <v:f eqn="if @0 0 @2"/>
  <v:f eqn="if @0 @4 21600"/>
  <v:f eqn="mid @5 @6"/>
  <v:f eqn="mid @8 @5"/>
  <v:f eqn="mid @7 @8"/>
  <v:f eqn="mid @6 @7"/>
  <v:f eqn="sum @6 0 @5"/>
 </v:formulas>
 <v:path textpathok="t" o:connecttype="custom" o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800"
  o:connectangles="270,180,90,0"/>
 <v:textpath on="t" fitshape="t"/>
 <v:handles>
  <v:h position="#0,bottomRight" xrange="6629,14971"/>
 </v:handles>
 <o:lock v:ext="edit" text="t" shapetype="t"/>
</v:shapetype><v:shape id="_x0000_s1028" type="#_x0000_t136" alt="registration"
 style='width:122.25pt;height:42pt;rotation:270' filled="f" strokecolor="#333">
 <v:shadow color="#868686"/>
 <v:textpath style='font-family:"Arial";font-weight:bold;v-text-kern:t' trim="t"
  fitpath="t" string="sign in"/>
</v:shape><![endif]--><![if !vml]><img border=0 width=58 height=165
src="Login_files/image002.gif" alt=registration v:shapes="_x0000_s1028"><![endif]></td>
		<td>
<form method="POST" action="<%=LOGON_PAGE%>">
	<dl>
		<dd>&nbsp;</dd>
		<dd><font face="Calibri">
		<a name="New Users:5"><b>
		<font color="#FF0000">In order to take a course on this site--and 
		receive CE credit, </font> <i>
<a href="#First Name" style="text-decoration: none"><font color="#FF0000">you must 
first sign in</font></a><font color="#FF0000">.</font></i><font color="#FFFFFF">&nbsp;</font></b></a></font><a name="New Users:1"><b><p>&nbsp;</p>
		<TABLE style="border-collapse: collapse" bordercolor="#111111" cellpadding="0" bgcolor="#044996" id="table3" width="341">
			<TR>
				<TD ALIGN="right"><em><b><font face="Calibri" color="#FFFFFF">
				Email</font></b></em></TD>
				<TD><input type="text" name="USER" size="25"></TD>
			</TR>
			<TR>
				<TD ALIGN="right"><em><b><font face="Calibri" color="#FFFFFF">Password</font></b></em></TD>
				<TD><input type="password" name="PWD" size="25"></TD>
			</TR>
		</TABLE>
		<p><input type="submit" value="SignIn"></dd>
	</dl>
	</a>

</a><font face="Calibri">
		<font color="#FF0000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </font>
	<a name="New Users:1" href="Registration%20Info%20Form.htm"><i>
	Don't have an account? Create one here.</i></a></font><a name="New Users:1"><p>&nbsp;</p>
</form>
		</b>

<a name="New Users:"><a name="New Users:1">
		<p align="center"><font size="1" face="Calibri">Hi-R-Ed Online will not 
		sell or share your personal information with any third party for any 
		purpose other than to facilitate the delivery of your CE Certificates of 
		Completion. You will never be asked for any billing information on this 
		site, for any purpose. Information collected in the registration form 
		and other forms on this site is that required by the professional 
		organizations Hi-R-Ed Online is accredited with as a provider of 
		continuing education. </font></td>
	</tr>
	<tr>
		<td width="72">&nbsp;</td>
		<td>
&nbsp;</td>
	</tr>
</table>
</div>
<hr width="850">
<div align="center">
	<table border="0" width="850" style="border-collapse: collapse">
		<tr>
			<td>
<a name="New Users:2"><span style="font-weight: 400">
			<b>
			<font size="1">
<img border="0" src="logo3(icon).jpg" width="43" height="23"></font><br>
			</b>
<font size="1">Copyright � 2017 Hi-R-Ed Online. All rights reserved.</font></span></a></a><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</b><a name="New Users:"><a name="New Users:2"><span style="font-weight: 400">
			<font size="1"><br>
Revised: <!--webbot bot="TimeStamp" s-type="EDITED" s-format="%m/%d/%y" startspan -->09/28/19<!--webbot bot="TimeStamp" i-checksum="13582" endspan --></font></span></a></a><span style="font-weight: 400"><font size="1"> </font></span>
<font face="Calibri" size="1" color="#FFFFFF">Rick Fields-Gardner</font></td>
		</tr>
	</table>
</div>

</a></b>
<p><b>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;

</p>

</body>

</html>