<%Option Explicit%>  
<html>  
<head>  
<title>ShotDev.Com Tutorial</title>  
</head>  
<body>  
<%  
Dim objFSO,oInStream,sRows,arrRows  
Dim sFileName  
  
sFileName = "_private/registration form_results.csv"  
  
'*** Create Object ***'  
Set objFSO = CreateObject("Scripting.FileSystemObject")  
  
'*** Check Exist Files ***'  
If Not objFSO.FileExists(Server.MapPath(sFileName)) Then  
Response.write("File not found.")  
Else  
  
'*** Open Files ***'  
Set oInStream = objFSO.OpenTextFile(Server.MapPath(sFileName),1,False)  
  
%>  
<table width="600" border="1">  
<tr>  
<th width="91"> <div align="center">CustomerID </div></th>  
<th width="98"> <div align="center">Name </div></th>  
<th width="198"> <div align="center">Email </div></th>  
<th width="97"> <div align="center">CountryCode </div></th>  
<th width="59"> <div align="center">Budget </div></th>  
<th width="71"> <div align="center">Used </div></th>  
</tr>  
<%  
Do Until oInStream.AtEndOfStream  
sRows = oInStream.readLine  
arrRows = Split(sRows,",")  
%>  
<tr>  
<td><div align="center"><%=arrRows(0)%></div></td>  
<td><%=arrRows(1)%></td>  
<td><%=arrRows(2)%></td>  
<td><div align="center"><%=arrRows(3)%></div></td>  
<td align="right"><%=arrRows(4)%></td>  
<td align="right"><%=arrRows(5)%></td>  
</tr>  
<%  
Loop  
  
oInStream.Close()  
Set oInStream = Nothing  
  
End IF  
%>  
</table>  
</body>  
</html>  