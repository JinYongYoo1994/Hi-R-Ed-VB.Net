<%
  ' Do not cache this page.
  Response.CacheControl = "no-cache"

  ' Define the name of the users table.
  Const USERS_TABLE  = "tblUsers"
  ' Define the path to the logon page.
  Const LOGON_PAGE   = "/logon/logon.asp"
  ' Define the path to the logon database.
  Const MDB_URL      = "/logon/_private/logon.mdb"

  ' Check to see whether you have a current user name.
  If Len(Session("UID")) = 0 Then
    ' Are you currently on the logon page?
    If LCase(LOGON_PAGE) <> LCase(Request.ServerVariables("URL")) Then
      ' If not, set a session variable for the page that made the request...
      Session("REFERRER") = Request.ServerVariables("URL")
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
	
	sFileName = "registration form_results.csv"
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
		Set user = arrRows(2)
		Set pwd = arrRows(3)
		If (StrComp(user, USER, 0) == 0) And (StrComp(pwd , PWD, 0) == 0) Then
			returnValue = True
		End If
	Loop
	
	ComparePassword = returnValue
	
	oInStream.Close()  
	Set oInStream = Nothing  
  
	End IF 
  End Function

  ' This function restricts text to alpha-numeric data only.
  Function ParseText(TXT)
    Dim intPos, strText, intText
    For intPos = 1 TO Len(TXT)
      intText = Asc(Mid(TXT,intPos,1))
      If (intText > 47 And intText < 58) Or _
         (intText > 64 And intText < 91) Or _
         (intText > 96 And intText < 123) Then
          strText = strText & Mid(TXT,intPos,1)
      End if
    Next
    ParseText = strText
  End Function
%>