# ClassicASP500

A simple page in VBScript for Classic ASP 500 error reporting in IIS 10.

## 500.asp

Compiles error information using objASPError, generates Session and Application table variable dumps, and sends error information to a specificed email address. Displays simple error message for end-user. Added admin check to disable email sending and displaying error information on page itself.

## web.config

Settings required for 500.asp to work. IIS no longer supports/reads customErrors and using the error tag causes an empty `Server.GetLastError`. You must set `500.asp` to defaultPath in httpErrors tag. If you are running the Web Applicaiton as a subfolder, ensure permission access on the parent/main folder.
