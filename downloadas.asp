<%@LANGUAGE="JAVASCRIPT" CODEPAGE="1252"%>
<%
// clear buffer
Response.Buffer = true;
Response.Clear();
// path to file
var sFilePath = Server.MapPath("AnnualReport.pdf");
// attachment name (without extension), this will be what the name of the downloaded file will be
var sAttachName = "attachmentname";
// remove spaces, else file may not download
sAttachName = sAttachName.replace(/\s/g,"");
// create filesystem object
var oFSO = Server.CreateObject("Scripting.FileSystemObject");
// check if file exists
if(!oFSO.FileExists(sFilePath)){
	Response.Write('<strong>File does not exist. <a href="'+Request.ServerVariables("HTTP_REFERER")+'">Go back</s>.</strong>');
	Response.End();
}
// get file
var oFile = oFSO.GetFile(sFilePath);
// get file size
var nFileSize = oFile.Size;
// get extension
var sExt = sFilePath.substring(sFilePath.lastIndexOf("."));
// download as attachment
Response.AddHeader("Content-Disposition","attachment;filename="+sAttachName+sExt);
// add content length header (for showing progress of download)
Response.AddHeader("Content-Length", nFileSize); 
// set to always download by setting content type
Response.ContentType = "application/octet-stream";
// create stream
var oStream = Server.CreateObject("ADODB.Stream");
// open stream
oStream.Open;
// set type of stream to binary
oStream.Type = 1;
// load data from file into stream
oStream.LoadFromFile(sFilePath);
// send data to client
Response.BinaryWrite(oStream.Read);
// close stream and end output to client
oStream.Close;
Response.End();
%>