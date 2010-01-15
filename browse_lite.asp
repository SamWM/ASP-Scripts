<%@ LANGUAGE="JSCRIPT" %>
<!--#include virtual="scripts/filefolder.asp" -->
<%
// error message
var sError = false

// get this URL
var sURL = Request.ServerVariables("URL")

// get folders from querystring
var qsFolder = (Request.QueryString)?String(Request.QueryString):""

// this is for toggling between showing and hiding the folder list
function toggleFolder(folder){
  // if folder already exists in array, remove it from querystring
  if(inQS(folder)) return qsFolder.replace(folder+";","")
  // if not, append to querystring
  return qsFolder+folder+";"
}
// check if folder is in querystring
function inQS(folder){
  if(qsFolder=="") return false
  qsFolderArray = qsFolder.split(";")
  return inArray(folder,qsFolderArray)
}
// check if item (itm) is in array (ar)
function inArray(itm,ar){
  for(i in ar){
    if(ar[i]==itm) return true
  } return false
}

function buildFileList(folder,target){
    // check if folder is in array
    var blnFound = inQS(folder)
    if(!blnFound) return ""
    // create enumerator of files
    var eFiles = getFiles(folder)
    // allowed file types
    var arAllow = new Array("rtf","doc","pdf","htm","html","xls","ppt")
    var s = "<ul>" // what we want to return
    // set default target to blank if undefined
    target = (!target)?"_blank":target
    // no files found
    if(eFiles.atEnd()){
        s+="<li>No files found</li>"
    }
    else{ // loop through the enumerator of files
        for(;!eFiles.atEnd();eFiles.moveNext()){
            // path to file
            filePath = String(eFiles.item().Path)
            // file extension
            fileExt = getFileExt(filePath.toLowerCase());
            // check if file is an allowed type
            if(inArray(fileExt,arAllow)){
                // virtual path to file
                vPath = folder+"/"+eFiles.item().Name;
                s+="<li class=\""+fileExt+"\"><a href=\""+vPath+"\"";
                s+=" target=\""+target+"\"";
                s+=" title=\"file size: "+getFileSize(filePath)+"\nlast modified: "+getLastModified(filePath)+"\">";
                s+=getFileName(filePath,false);
                s+="</a></li>";
            }
        }
    }
    return s+"</ul>"
}

var arLinkText = new Array()
var arLinkURL = new Array()

arLinkText = [
"Home",
"Folder 1",
"Folder 2",
"Folder 3"
]

arLinkURL = [
"home.asp",
"/folder1",
"/folder2",
"/folder3"
]
if(arLinkText.length!=arLinkURL.length) sError = "arLinkText and arLinkURL should have the same number of items"
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
    "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Files in Folder <%=qsFolder%></title>
</head>
<body>
<%if(sError){Response.Write(sError)}else{ // an error did not occur%>
<a href="<%=sURL%>" title="Collapse all"><img src="images/minimize.gif" border="0" alt="Collapse all" /></a>
<%
var sFolderList = ""
for(var l in arLinkText){
    if(arLinkURL[l].indexOf(".")==-1){
        sFolderList+=arLinkURL[l]+";"
    }
}
%>
<a href="<%=sURL%>?<%=sFolderList%>" title="Expand all"><img src="images/maximize.gif" border="0" alt="Expand All" /></a>
<ul>
<%
for(var l in arLinkURL){
    if(arLinkURL[l].indexOf(".")!=-1){
    // it is a file
%>
<li class="document"><a href="<%=arLinkURL[l]%>" target="rbottom"><%=arLinkText[l]%></a></li>
<%
    } // file
    else{
    // it is a folder
    liclass = inQS(arLinkURL[l])?"open":"closed"
%>
<li class="<%=liclass%>"><a name="<%=arLinkURL[l]%>" href="<%=sURL%>?<%=toggleFolder(arLinkURL[l])%>"><%=arLinkText[l]%></a><%=buildFileList(arLinkURL[l])%></li>
<%
    }// folder
}//end for l in arLinkText
%>
</ul>
<%} // an error did not occur%>
</body>
</html>