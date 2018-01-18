<%option explicit%>
<!DOCTYPE html>
<%	
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
	style = style & "#sessionInfo th{ text-align:left; background-color:cyan;}"
	style = style & "#sessionInfo td{ background-color:lightcyan; }"
	style = style & "#appInfo th{ text-align:left; background-color:tomato; }"
	style = style & "#appInfo td{ background-color:pink; }"
	style = style & "#serverInfo th{ text-align:left; background-color:limegreen;}"
	style = style & "#serverInfo td{ background-color:palegreen; }"
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
	debugInfo = debugInfo & "<tr> <th>User IP</th>	<td>"& Request.ServerVariables("REMOTE_HOST") & " (" & Request.ServerVariables("REMOTE_ADDR") &")</td> </tr>"
	debugInfo = debugInfo & "<tr> <th>Browser</th>	<td>"& Request.ServerVariables("HTTP_USER_AGENT") &"</td> </tr>"
	debugInfo = debugInfo & "<tr> <th>Server</th>	<td>"& Request.ServerVariables("SERVER_NAME") & " (" & Request.ServerVariables("LOCAL_ADDR") &")</td> </tr>"
	debugInfo = debugInfo & "<tr> <th>Referer</th>	<td>"& Request.ServerVariables("HTTP_REFERER") & "</td> </tr>"	
	debugInfo = debugInfo & "<tr> <th>POST</th>		<td>"& Request.Form &"</td> </tr>"
	debugInfo = debugInfo & "</table>"	
	
	'Session Variables
	Dim sessionVar, sessionVarItem, sessionTable
	sessionTable = "<table id='sessionInfo'>"	
	sessionTable = sessionTable & "<tr><th colspan='2' style='text-align:center;background-color:dodgerblue;padding:2px;'>Session Variables ("& Session.Contents.Count  &")</th></tr>"	
	For Each sessionVar in Session.Contents
		If IsArray(Session(sessionVar)) Then
			For appVarItem = LBound(Session(sessionVar)) to UBound(Session(sessionVar))
				sessionTable = sessionTable & "<tr>"
				sessionTable = sessionTable & "<th>" & sessionVar & " " & sessionVarItem & "</th>"
				sessionTable = sessionTable & "<td>" & Session(sessionVar)(sessionVarItem) & "</td>"
				sessionTable = sessionTable & "</tr>"
			Next
  		Else
			sessionTable = sessionTable & "<tr>"
			sessionTable = sessionTable & "<th>" & sessionVar & "</th>"
			sessionTable = sessionTable & "<td>" & Session(sessionVar) & "</td>"
			sessionTable = sessionTable & "</tr>"
  		End If
	Next
	sessionTable = sessionTable & "</table>"

	'Application Variables
	Dim appVar, appVarItem, appVarTable
	appVarTable = "<table id='appInfo'>"	
	appVarTable = appVarTable & "<tr><th colspan='2' style='text-align:center;background-color:lightcoral;padding:2px;'>Application Variables ("& Application.Contents.Count  &")</th></tr>"	
	For Each appVar in Application.Contents
		If IsArray(Application(appVar)) Then
			For appVarItem = LBound(Application(appVar)) to UBound(Application(appVar))
				appVarTable = appVarTable & "<tr>"
				appVarTable = appVarTable & "<th>" & appVar & " " & appVarItem & "</th>"
				appVarTable = appVarTable & "<td>" & Application(item)(appVarItem) & "</td>"
				appVarTable = appVarTable & "</tr>"
			Next
  		Else
			appVarTable = appVarTable & "<tr>"
			appVarTable = appVarTable & "<th>" & appVar & "</th>"
			appVarTable = appVarTable & "<td>" & Application(appVar) & "</td>"
			appVarTable = appVarTable & "</tr>"
  		End If
	Next
	appVarTable = appVarTable & "</table>"

	'Server Variables
	Dim serverVar, serverVarItem, serverTable
	serverTable = "<table id='serverInfo'>"	
	serverTable = serverTable & "<tr><th colspan='2' style='text-align:center;background-color:mediumseagreen;padding:2px;'>Server Variables ("& Request.ServerVariables.Count  &")</th></tr>"	
	For Each serverVar in Request.ServerVariables
		If IsArray(Request.ServerVariables(serverVar)) Then
			For serverVarItem = LBound(Request.ServerVariables(serverVar)) to UBound(Request.ServerVariables(serverVar))
				serverTable = serverTable & "<tr>"
				serverTable = serverTable & "<th>" & serverVar & " " & serverVarItem & "</th>"
				serverTable = serverTable & "<td>" & Request.ServerVariables(item)(serverVarItem) & "</td>"
				serverTable = serverTable & "</tr>"
			Next
  		Else
			serverTable = serverTable & "<tr>"
			serverTable = serverTable & "<th>" & serverVar & "</th>"
			serverTable = serverTable & "<td>" & Request.ServerVariables(serverVar) & "</td>"
			serverTable = serverTable & "</tr>"
  		End If
	Next
	serverTable = serverTable & "</table>"

	' Compile Error Information
	dim errorDump
	errorDump = errMsg & debugInfo & sessionTable & appVarTable & serverTable

	'Send Email
	dim tech : tech = CreateObject("WScript.Network").UserName = "ITGuy"
	If NOT tech Then
		Dim oMessage : Set oMessage = CreateObject("CDO.Message") 
		oMessage.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") 	 = 2
		oMessage.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") 	 = "smtp.webapp.com"
		oMessage.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		oMessage.Configuration.Fields.Update

		oMessage.To = "it@webapp.com"
		oMessage.From = "WebApp@webapp.com"
		oMessage.Subject = "WebApp Error - " & objASPError.Category
		oMessage.htmlBody = style & errorDump		
		oMessage.Send		
	End If
%>
<html lang="en">
	<head>
		<meta charset="UTF-8">
		<title>WebApp Error</title>
		<% If tech Then Response.Write style%>
	</head>
	<body>
		Sorry! The system encountered an error. An email has been sent to IT. If you have any concerns or questions please send an email to <a href="mailto:help@webapp.com">help@webapp.com</a> or call x555.
		<%If tech Then Response.Write errorDump %>
	</body>
</html>
