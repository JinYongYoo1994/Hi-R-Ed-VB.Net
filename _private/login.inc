<%
  ' Do not cache this page.
  Response.CacheControl = "no-cache"

  ' Define the name of the users table.
  Const USERS_TABLE  = "tblUsers"
  ' Define the path to the logon page.
  Const LOGON_PAGE   = "/Login.asp"


  ' Check to see whether you have a current user name.
  If Len(Session("USER")) = 0 Then
    ' Are you currently on the logon page?
    If LCase(LOGON_PAGE) <> LCase(Request.ServerVariables("URL")) Then
      ' If not, set a session variable for the page that made the request...
      Session("REFERRER") = "/sponsors(new).asp"
      ' ...and redirect to the logon page.
      Response.Redirect LOGON_PAGE
    End If
  End If
  
  ' This function checks for a username/password combination.
  Function ComparePassword(USER,PWD)
    ' Define your variables.
    Dim objFSO,oInStream,sRows,arrRows  
	Dim sFileName
	Dim returnValue
	
	sFileName = "_private/registration_results.csv"
	returnValue = False
	
	
    '*** Create Object ***'  
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	'*** Check Exist Files ***'  
	
	If Not objFSO.FileExists(Server.MapPath(sFileName)) Then  
	Response.write("File not found.")  
	
	Else 

		'*** Open Files ***'  
		Set oInStream = objFSO.OpenTextFile(Server.MapPath(sFileName),1,False) 
		
		Do Until oInStream.AtEndOfStream  
			sRows = oInStream.readLine  
			arrRows = Split(sRows,",")
			dUser = arrRows(14)
			dPwd = arrRows(15)
			tUser = """" + USER + """"
			tPwd = """" + PWD + """"
			If (dUser = tUser) And (dPwd = tPwd) Then
				returnValue = True
			End If
		Loop
		
		ComparePassword = returnValue
		
		oInStream.Close()  
		Set oInStream = Nothing  
  
	End IF 
  End Function

%>