<%@ Language=JavaScript %>
<!--#include virtual="scripts/filefolder.asp" -->
<!--#include virtual="scripts/datefunctions.asp" -->
<%
var arLinks = new Array()
arLinks = [
	// ["Text to show as link","URL to file/folder (folder must end with /)"]
	["Home","home.asp"],
	["Folder 1","/folder1/"],
	["Folder 2","/folder2/"],
	["Folder 3","/folder3/"]
]

// error message
var sError = false

// how many days since last modification should a file be marked as new? Also setting for cookies and menu
var nNewDays = 7

// get this URL
var sURL = Request.ServerVariables("URL")

// check if cookies have been set
var ckFolders = String(Request.Cookies("folders"))

// if "folders" cookie is not set ckFolders will be undefined
// also check if this page has already been loaded with session "active"
if(ckFolders!="undefined" && !Session("active")){
	Session("active")=true // set session "active" to 1 as we don't want to redirect on every page load
	Response.Redirect(sURL+"?"+ckFolders) // reload page with folder settings from cookie
}

// get folders from querystring
var qsFolder = (Request.QueryString)?String(Request.QueryString):""
// set "folders" cookie
Response.Cookies("folders")=qsFolder
// set expiry on cookie
var dtExpires = new Date()
dtExpires.setTime(dtExpires.getTime() + ( nNewDays*24*60*60*1000 ) ) // expire in nNewDays days
Response.Cookies("folders").Expires = dtExpires.toLocaleString()

// this is for toggling between showing and hiding the folder list
function toggleFolder(folder){
	// if folder already exists in querystring, remove it
	if(inQS(folder)) return qsFolder.replace(folder+";","")
	// if not, append to querystring
	return qsFolder+folder+";"
}

// check if folder is in querystring
function inQS(folder){
	if(qsFolder=="") return false
	qsFolderArray = qsFolder.split(";")
	return inArray(folder, qsFolderArray)
}
// check if item (itm) is in array (ar)
function inArray(itm,ar){
	for(i in ar){
		if(ar[i]==itm) return true
	} return false
}

// build a list of files in a folder
function buildFileList(folder,target){
	// todays date, for date comparison
	var dtToday = new Date().toLocaleString();
	// return zero length string if folder is not part of the querystring
	if(!inQS(folder)) return ""
	// load list from application variable if it is set and if the folder date has changed (files removed or added) and if the cookie expiry date has passed
	if(Application(folder) && (String(getFolderLastModified(folder))==Application(folder+"date")) && dtExpires < dtToday) return Application(folder)
	// create enumerator of files
	var eFiles = getFiles(folder)
	// allowed file types
	var arAllow = new Array("rtf","doc","pdf","htm","html","xls","ppt")
	var s = "<ul>" // what we want to return
	// set default target to blank if undefined
	var target = (!target)?"blank":target
	// no files found
	if(eFiles.atEnd()){
		s+="<li>No files found</li>"
	}else{ // loop through the enumerator of files
		for(;!eFiles.atEnd();eFiles.moveNext()){
			// path to file
			var sFilePath = String(eFiles.item().Path)
			// file extension
			var sFileExt = getFileExt(sFilePath);
			// check if file is an allowed type
			if(inArray(sFileExt,arAllow)){
				// file name for URL
				// remove extension
				var sFNUrl = rmFileExt(eFiles.item().Name);
				// encode URL and replace +'s with %20
				sFNUrl = Server.URLEncode(sFNUrl).replace(/\+/g,"%20");
				// virtual path to file
				var vPathURL = folder+sFNUrl+"."+sFileExt;
				// file name, HTML encoded, no extension
				var sFileName = Server.HTMLEncode(rmFileExt(eFiles.item().Name));
				// set class name to file extension
				var sClassName = sFileExt;
				s+="<li class=\""+sClassName+"\"><a href=\""+vPathURL+"\"";
				// last modified
				var dtLastModified = getLastModified(sFilePath);
				// day difference
				nDayDifference = dayDifference(dtToday,dtLastModified);
				// mark as new if appropriate
				s+=(nDayDifference<nNewDays)?" class=\"new\"":"";
				s+=" target=\""+target+"\"";
				s+=" title=\""+sFileName;
				s+="\n\nfile size: "+fmtFileSize(eFiles.item().Size);
				s+="\nlast modified: "+dtLastModified;
				s+="\">";
				s+=sFileName;
				s+="</a></li>";
			}
		}
	}
	s+="</ul>"
	// add list to application variable
	Application(folder) = s
	// add folder date application variable
	Application(folder+"date") = String(getFolderLastModified(folder))
	return s
}
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
	"http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Files in Folder <%=qsFolder%></title>
</head>
<body>
<%if(sError){Response.Write(sError)}else{ // an error did not occur%>
<a class="new">This colour text</a> indicates a new document (less than <%=nNewDays%> days old).<br /><br />
<a href="<%=sURL%>" title="Collapse all"><img src="images/minimize.gif" border="0" alt="Collapse all" /></a>
<%
// build list of folder links for Expand All button
// folders are detected in the array by checking if the 2nd array item ends with a forward slash
var sFolderList = ""
for(var l in arLinks){
	if(arLinks[l][1].lastIndexOf("/")==arLinks[l][1].length-1){
		sFolderList+=arLinks[l][1]+";"
	}
}
%>
<a href="<%=sURL%>?<%=sFolderList%>" title="Expand all"><img src="images/maximize.gif" border="0" alt="Expand All" /></a>
<ul>
<%
// now build the menu itself, doing same check for folder as above
for(var l in arLinks){
	if(arLinks[l][1].lastIndexOf("/")!=arLinks[l][1].length-1){
	// it is a file (no forward slash at end)
%>
<li class="document"><a href="<%=arLinks[l][1]%>" target="rbottom"><%=arLinks[l][0]%></a></li>
<%
	} // file
	else{
	// it is a folder
	liclass = inQS(arLinks[l][1])?"open":"closed"
%>
<li class="<%=liclass%>"><a name="itm<%=l%>" href="<%=sURL%>?<%=toggleFolder(arLinks[l][1])%>#itm<%=l%>"><%=Server.HTMLEncode(arLinks[l][0])%></a><%=buildFileList(arLinks[l][1])%></li>
<%
	}// folder
}//end for l in arLinkText
%>
</ul>
<%} // an error did not occur%>
A <img src="images/closed.gif" /> image indicates a folder containing documents. Click the text to show them.
</body>
</html>