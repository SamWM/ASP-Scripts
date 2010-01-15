<% 

call downloadFile(replace(replace(Request("file"),"\",""),"/",""))

function downloadFile(strFile)
' make sure you are on the latest MDAC version for this to work
' -------------------------------------------------------------


' get full path of specified file
strFilename = server.MapPath(strFile)


' clear the buffer
Response.Buffer = True
Response.Clear

' create stream
Set s = Server.CreateObject("ADODB.Stream")
s.Open

' set as binary
s.Type = 1

' load in the file
on error resume next


' check the file exists
Set fso = Server.CreateObject("Scripting.FileSystemObject")
if not fso.FileExists(strFilename) then
	Response.Write("<h1>Error:</h1>" & strFilename & " does not exist<p>")
	Response.End
end if


' get length of file
Set f = fso.GetFile(strFilename)
intFilelength = f.size

 
s.LoadFromFile(strFilename)
if err then
	Response.Write("<h1>Error: </h1>" & err.Description & "<p>")
	Response.End
end if

' send the headers to the users browser
Response.AddHeader "Content-Disposition", "attachment; filename=" & f.name
Response.AddHeader "Content-Length", intFilelength
Response.Charset = "UTF-8"
Response.ContentType = "application/octet-stream"

' output the file to the browser
Response.BinaryWrite s.Read
Response.Flush


' tidy up
s.Close
Set s = Nothing


end function

%>
