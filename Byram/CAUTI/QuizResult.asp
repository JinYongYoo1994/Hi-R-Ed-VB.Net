<%@ Language = VBScript %>
<%Option Explicit
response.buffer = true %>
<html>
<head>
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Quiz Results Page</title>
</head>
<body>
<%
'response.clear
'response.expires = 0
Dim strTitle
strTitle = request.form ("hidTitle")
Response.write "<font face='verdana'>"
If request.cookies("styleQuiz") = "Quiz done ! " & strTitle then
    Response.write "<font face='verdana'> Sorry you have already taken the quiz !!</font>"
Else
    response.cookies("styleQuiz") = "Quiz done ! " & strTitle
    response.cookies("styleQuiz").expires = Date
	Dim Qnum,intAns,intCMark,intNMark,intQCount,intMark,intCAns,intPercent
	intQCount = request.form ("hidCount")
	intCMark = request.form ("hidCMark")
	intNMark = request.form ("hidNMark")
	intMark = 0
	for Qnum = 1 to intQCount
		intAns = request.form ("OptQ" & Qnum)
		intCAns = request.form ("hidCA" & Qnum)
		if intAns = intCAns then
			intMark = intMark + intCMark
		elseif intAns <> "" then
			intMark = intMark - intNMark
		else
			intMark = intMark	
		end if
	next
	response.write ("<p align='center'><b><font color='maroon' size=4>Thank you </font></b></p>" & _
    "<p align='left'>Thank you for taking the post-test. Here is your result</p>")
	intPercent = intMark * 100 / (intQCount * intCMark)
	if (intPercent >= 80) then
		Response.write "<font color='green' size='2'><b>Congratulations, you passed !! <br><br>All grading is Pass/Fail. <br><br>Click on the link below to view and print your CEU certificate. <br><br>"
		Response.write "<br><br><font color='blue' size='4'><a href='credentials%20choice.htm'>CEU Certificate</a></font></font><br><br>"
	elseif (intPercent > 60) and (intPercent < 80) then
		Response.write "<font color='blue' size='2'><b>You did not pass--please try again ! Click on your browser's Back button to retake the test. <br><br>All grading is Pass/Fail. You must score at least 75% correct in order to pass. </b></font><br>"
	elseif (intPercent > 40) and (intPercent < 60)then
		Response.write " <font color='navy' size='2'><b>You did not pass--please try again ! Click on your browser's Back button to retake the test. <br><br>All grading is Pass/Fail.  You must score at least 75% correct in order to pass. </b></font><br>"
	else
		Response.write "<font color='red' size='2'><b>You did not pass--please try again ! Click on your browser's Back button to retake the test. <br><br>All grading is Pass/Fail. You must score at least 75% correct in order to pass. </font><br>"
	end if
	Response.write " Your Score is <font color='#008200'<b>" & intPercent & "%</b></font><br>"
End if
Response.write "</font>"
%>
</body>

</html>