<%option explicit%>
<!DOCTYPE html>
<%	
	'Generate Table
	Function GenerateTable(obj, table)
		Dim var, varItem, htmTable		
		htmTable = "<table id='info"& table &"'>"	
		htmTable = htmTable & "<tr><th colspan='2'>"& table &" Variables ("& obj.Count  &")</th></tr>"	
		For Each var in obj
			If IsArray(obj(var)) Then
				For varItem = LBound(obj(var)) to UBound(obj(var))
					htmTable = htmTable & "<tr>"
					htmTable = htmTable & "<th>" & var & " " & varItem & "</th>"
					htmTable = htmTable & "<td>" & obj(var)(varItem) & "</td>"
					htmTable = htmTable & "</tr>"
				Next
	  		Else
				htmTable = htmTable & "<tr>"
				htmTable = htmTable & "<th>" & var & "</th>"
				htmTable = htmTable & "<td>" & obj(var) & "</td>"
				htmTable = htmTable & "</tr>"
	  		End If
		Next
		htmTable = htmTable & "</table>"

		GenerateTable = htmTable
	End Function

	'Get Error
	Dim objASPError
	Set objASPError = Server.GetLastError

	'Reset Response
	If Response.Buffer Then
	    Response.Clear
	    Response.Status = "500 Internal Server Error"
	    Response.ContentType = "text/html"
	    Response.Expires = 0
	End If

	'CSS
	dim style : style = "<style type='text/css'>"
	style = style & "table { width: 800px;} "
	style = style & "#debugInfo th{ text-align:left; background-color:palegoldenrod; }"
	style = style & "#debugInfo td{ background-color:lightgoldenrodyellow; }"
	style = style & "#infoSession tr:first-child th{ text-align:center;bpadding:2px;background-color:dodgerblue; }"
	style = style & "#infoSession th{ text-align:left; background-color:cyan;}"
	style = style & "#infoSession td{ background-color:lightcyan; }"
	style = style & "#infoApplication tr:first-child th{ text-align:center;padding:2px;background-color:tomato; }"
	style = style & "#infoApplication th{ text-align:left; background-color:lightcoral; }"
	style = style & "#infoApplication td{ background-color:pink; }"
	style = style & "#infoServer tr:first-child th{ text-align:center;padding:2px;background-color:mediumseagreen; }"
	style = style & "#infoServer th{ text-align:left; background-color:limegreen;}"
	style = style & "#infoServer td{ background-color:palegreen; }"
	style = style & "</style>"

	' Error Message
	Dim errMsg : errMsg = ""	
	errMsg = errMsg &"<p>"
	errMsg = errMsg & objASPError.Category & "(0x" & hex(objASPError.Number) & ")<br />"
	errMsg = errMsg & objASPError.Description & "<br />"
	errMsg = errMsg & objASPError.File & ", line " & objASPError.Line & "<br />"
	errMsg = errMsg & objASPError.Source & "<br />"
	errMsg = errMsg & "</p>"

	'Debug Info
	Dim debugInfo : debugInfo = ""
	debugInfo = debugInfo & "<table id='debugInfo'>"	
	debugInfo = debugInfo & "<tr> <th colspan='2' style='text-align:center;background-color:gold;padding:2px;'>Debug Information</th> </tr>"	
	debugInfo = debugInfo & "<tr> <th>User</th>		<td>"& Request.serverVariables("AUTH_USER") &"</td> </tr>"
	debugInfo = debugInfo & "<tr> <th>Time</th>		<td>"& Now() &"</td> </tr>"		
	debugInfo = debugInfo & "<tr> <th>Page</th>		<td>"& Request.ServerVariables("SCRIPT_NAME") &"</td> </tr>"
	debugInfo = debugInfo & "<tr> <th>File</th>		<td>"& Session("Tablelist-tablelabel") & " ("& Session("thefile") & ")</td> </tr>"
	debugInfo = debugInfo & "<tr> <th>Report</th>	<td>"& Session("rname") & " ("& Session("seq_no") & ")</td> </tr>"
	debugInfo = debugInfo & "<tr> <th>User IP</th>	<td>"& Request.ServerVariables("REMOTE_HOST") & " (" & Request.ServerVariables("REMOTE_ADDR") &")</td> </tr>"
	debugInfo = debugInfo & "<tr> <th>Browser</th>	<td>"& Request.ServerVariables("HTTP_USER_AGENT") &"</td> </tr>"
	debugInfo = debugInfo & "<tr> <th>Server</th>	<td>"& Request.ServerVariables("SERVER_NAME") & " (" & Request.ServerVariables("LOCAL_ADDR") &")</td> </tr>"
	debugInfo = debugInfo & "<tr> <th>Referer</th>	<td>"& Request.ServerVariables("HTTP_REFERER") & "</td> </tr>"	
	debugInfo = debugInfo & "<tr> <th>POST</th>		<td>"& Request.Form &"</td> </tr>"
	debugInfo = debugInfo & "</table>"	
	
	'Session Variables
	dim sessionTable
	sessionTable = GenerateTable(Session.Contents, "Session")

	'Application Variables
	dim appTable
	appTable = GenerateTable(Application.Contents, "Application")

	'Server Variables
	dim serverTable
	serverTable = GenerateTable(Request.ServerVariables, "Server")

	' Compile Error Information
	dim errorDump
	errorDump = errMsg & debugInfo & sessionTable & appTable & serverTable

	'Send Email
	dim tech : tech = CreateObject("WScript.Network").UserName = "ITGuy"
	If NOT tech Then
		Dim oMessage : Set oMessage = CreateObject("CDO.Message") 
		oMessage.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") 	 = 2
		oMessage.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") 	 = "smtp.webapp.com"
		oMessage.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		oMessage.Configuration.Fields.Update

		oMessage.To = "ITGuy@webapp.com"
		oMessage.From = "webapp@webapp.com"
		oMessage.Subject = Application("Name") & " Error - " & objASPError.Category
		oMessage.htmlBody = style & errorDump		
		oMessage.Send		
	End If
%>
<html lang="en">
	<head>
		<meta charset="UTF-8">
		<title><%=Application("Name")%> Error</title>
		<% If tech Then Response.Write style%>
	</head>
	<body>
		Sorry! The system encountered an error. An email has been sent to IT. If you have any concerns or questions please send an email to <a href="mailto:help@webapp.com">help@webapp.com</a> or call x5555
		<%If tech Then Response.Write errorDump %>
	</body>
</html>
