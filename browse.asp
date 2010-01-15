<%@ Language=JavaScript %>
<!--#include virtual="scripts/filefolder.asp" -->
<!--#include virtual="scripts/datefunctions.asp" -->
<%
var arLinks = new Array();
arLinks = [
	// ["Text to show as link","URL to file/folder (folder must end with /)"]
	["Home","home.asp"],
	["Folder 1","/folder1/"],
	["Folder 2","/folder2/"],
	["Folder 3","/folder3/"],
	["Subpage","/subpage"] // link to another page, without needing default.asp
];

// error message
var sError = false;

// how many days since last modification should a file be marked as new? Also setting for cookies and menu
var nNewDays = 7;

// get this URL
var sURL = Request.ServerVariables("URL");

// check if cookies have been set
var ckFolders = String(Request.Cookies("folders"));

// if "folders" cookie is not set ckFolders will be undefined
// also check if this page has already been loaded with session "active"
if(ckFolders!="undefined" && !Session("active")){
	Session("active")=1; // set session "active" to 1 as we don't want to redirect on every page load
	//Response.Redirect(sURL+"?"+ckFolders); // reload page with folder settings from cookie
}

var ignoreFolders = new Array("Hide Me", "Admin Only");

// get folders from querystring
var qsFolder = (Request.QueryString)?String(Request.QueryString):"";
// set "folders" cookie
Response.Cookies("folders")=qsFolder;
// set expiry on cookie
var dtExpires = new Date();
dtExpires.setTime(dtExpires.getTime() + ( nNewDays*24*60*60*1000 ) ); // expire in nNewDays days
Response.Cookies("folders").Expires = dtExpires.toLocaleString();

// this is for toggling between showing and hiding the folder list
function toggleFolder(folder){
	// if folder already exists in querystring, remove it
	if(inQS(escape(folder))) return qsFolder.replace(escape(folder)+";","")
	// if not, append to querystring
	return qsFolder+escape(folder)+";"
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
	/* todays date, for date comparison */
	var dtToday = new Date().toLocaleString();
	/* return zero length string if folder is not part of the querystring */
	if(!inQS(escape(folder))){return ""};
	/* load list from application variable if it is set and if the folder date has changed (files removed or added) and if the cookie expiry date has passed */
	if(Application(folder) && (getFolderLastModified(folder)==Application(folder+"date")) && dtExpires < dtToday){"\n<!-- cached: -->\n"+Application(folder)};
	/* create enumerator of files */
	var eFiles = getFiles(folder);
	/* allowed file types */
	var arAllow = new Array("rtf","doc","pdf","htm","html","xls","ppt","lnk", "pub");
	/* build list of subfolders */
	var s = buildSubFolderList(folder)+"<!-- list of files in "+folder+" -->\n<ul>";
	/* set default target to blank if undefined */
	var target = (!target)?"blank":target
	/* no files found */
	if(eFiles.atEnd()){
		s+=""/*"\n<li>No files found in folder "+folder.replace(/%20/g," ")+"</li>\n"*/
	}else{ /* loop through the enumerator of files */
		for(;!eFiles.atEnd();eFiles.moveNext()){
			/* path to file */
			var sFilePath = String(eFiles.item().Path);
			/* file extension */
			var sFileExt = getFileExt(sFilePath);
			/* check if file is an allowed type */
			if(inArray(sFileExt.toLowerCase(),arAllow)){
				/* skip if temp file (created by word) */
				if(String(eFiles.item().Name).indexOf("~$")==0){continue}
				/* skip if hidden file */
				if(eFiles.item().Attributes & 2){continue}
				/* remove extension */
				var sFNUrl = rmFileExt(eFiles.item().Name);
				/* escape filename (for spaces) */
				sFNUrl = escape(sFNUrl);
				/* virtual path to file */
				var vPathURL = folder+sFNUrl+"."+sFileExt;
				/* file name, HTML encoded, no extension */
				var sFileName = Server.HTMLEncode(rmFileExt(eFiles.item().Name));
				/* set class name to file extension */
				var sClassName = sFileExt.toLowerCase();
				if(sClassName == "lnk") sClassName = "document";
				s+="\n<li class=\""+sClassName+"\">\n<a onclick=\"return showDocument(this.href);\" href=\""+vPathURL.replace(/\s{1}/g,"%20")+"\"";
				/* last modified */
				var dtLastModified = getLastModified(sFilePath);
				/* day difference */
				nDayDifference = dayDifference(dtToday,dtLastModified);
				/* mark as new if appropriate */
				s+=(nDayDifference<nNewDays)?" class=\"new\"":"";
				/* set target window */
				s+=" target=\""+target+"\"";
				/* set tooltip */
				s+=" title=\""+sFileName;
				s+="\nfile size: "+fmtFileSize(eFiles.item().Size);
				s+="\nlast modified: "+dtLastModified;
				s+="\">";
				/* link text is file name */
				s+=sFileName;
				/* close link and list item */
				s+="</a>\n</li>\n";
			}
		}
	}
	s+="\n</ul>\n<!-- end list of files in "+folder+"-->\n";
	/* // debug
	 s+="<li>"+folder+"</li>" */
	// add list to application variable */
	Application(folder) = s;
	// add folder date application variable */
	Application(folder+"date") = getFolderLastModified(folder);
	return s.replace(/undefined/g,"");
}

function buildSubFolderList(folder){
	/* todays date, for date comparison */
	var dtToday = new Date().toLocaleString();
	/* return zero length string if folder is not part of the querystring */
	if(!inQS(escape(folder))){return ""};
	/* load list from application variable if it is set and if the folder date has changed (files removed or added) and if the cookie expiry date has passed */
	if(Application(folder) && (String(getFolderLastModified(folder))==Application(folder+"date")) && dtExpires < dtToday){return Application(folder)};
	/* create enumerator of folders */
	var eFolders = getFolders(folder);
	/* return folder list */
	if(!eFolders.atEnd()){
		var s = "\n<!-- list of folders in "+folder+"-->\n<ul>"
		for(;!eFolders.atEnd();eFolders.moveNext()){
			/* get folder name */
			sFolderName = eFolders.item().Name;
			// folders to skip
			if(inArray(sFolderName, ignoreFolders))
			{
				continue;
			}
			/* get virtual path to folder by appending folder name to parent */
			sFolderVPath = folder+sFolderName+"/";
			/* get when the folder was last modified */
			dtLastModified = new Date(eFolders.item().DateLastModified);
			/* folders ending with _files are generated when word saves a webpage - this folder contains images related to that web page */
			if(/_files$/.test(sFolderName)){continue};
			/* set class to open (if folder in querystring) or closed (if folder not in querystring) */
			liclass = inQS(escape(sFolderVPath))?"open":"closed";
			/* generate link, starting with the URL of the page with the list of folders to be opened appended. Link to anchor point so page doen't scroll to the top - unless when closing a folder list */
			sURLLink = sURL+"?"+toggleFolder(sFolderVPath)+"#itm"+((liclass=="closed")?dtLastModified.getTime():"TOP");
			/* initialise list item, set class and anchor for folder */
			s+="<li class=\""+liclass+"\"><a name=\"#itm"+dtLastModified.getTime()+"\" class=\""+liclass+"\"";
			/* set link and tooltip */
			s+=" href=\""+sURLLink+"\" title=\"Modified: "+dtLastModified.toLocaleString()+"\">"+sFolderName+"</a>";
			/* build list of files in folder */
			s+=buildFileList(sFolderVPath);
			/* close list item */
			s+="</li>";
			/* // debug
			s+="<li>"+sFolderVPath+"</li>" */
		}
		s+="</ul>\n<!-- end list of folders in "+folder+"-->\n";
		return s.replace(/undefined/g,"");
	}
}
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN"
    "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Browse</title>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_displayStatusMsg(msgStr) { //v1.0
  status=msgStr;
  document.MM_returnValue = true;
}
function showDocument(sDocURL){
	if(sDocURL.indexOf("file:///")!=-1) return true;
	var newwin = window.open(sDocURL,'newwin','scrollbars=yes,toolbar=no,menubar=yes,top=0,left=0,width='+(window.screen.availWidth)+',height='+(window.screen.availHeight)+'');
	newwin.window.moveTo(0,0);
	if(sDocURL.indexOf(".pdf")==-1){
		newwin.window.resizeTo(screen.availWidth,screen.availHeight);
	}
	newwin.focus();
	return false;
}
//-->
</script>
</head>
<body>
<div id="outer">
	<div id="header">
	<h1>Browse</h1>
	</div>
  <div id="contents"> 
    <p>
      Go to <strong> <a href="http://www.google.co.uk/" target="_blank">Google 
      Search Engine</a></p>
  <h2>What's New</h2>
  <p><a onclick="showDocument(this.href); return false" href="gallery/" target="_blank">Photo Gallery</a><br />
  
  <p>Remember to visit our public site: <a href="http://www.example.org" target="_blank" onmouseover="MM_displayStatusMsg('Go to Website');return document.MM_returnValue" onmouseout="MM_displayStatusMsg('');return document.MM_returnValue">www.example.org</a></p>
</div>
<div id="nav">
<a name="itmTOP"></a>
<%if(sError){Response.Write(sError)}else{ // an error did not occur%><a class="new">This colour text</a> indicates a new document (less than <%=nNewDays%> days old).<br /><br />
<%
// build list of folder links for Expand All button
// folders are detected in the array by checking if the 2nd array item ends with a forward slash
var sFolderList = ""
for(var l in arLinks){
	var sLink = arLinks[l][1]
	if(sLink.lastIndexOf("/")==sLink.length-1){
		sFolderList+=sLink+";"
	}
}
%><a href="<%=sURL%>?<%=sFolderList%>" title="Expand all"><img src="images/maximize.gif" border="0" alt="Expand All" /></a>
<!-- menu listing -->
<ul><%
// now build the menu itself, doing same check for folder as above
for(var l in arLinks){
	var sLink = arLinks[l][1]
	var sText = arLinks[l][0]
	if(sLink.lastIndexOf("/")!=sLink.length-1){
	// it is a file (no forward slash at end)
%><!-- file <%=sLink%> --><li class="document"><a onclick="showDocument(this.href); return false" href="<%=sLink%>"><%=sText%></a></li><!-- end file <%=sLink%> --><%
	} // file
	else{
	// it is a folder
	liclass = inQS(sLink)?"open":"closed"
%>
<!-- folder <%=sLink%> -->
<li class="<%=liclass%>"><a class="<%=liclass%>" name="itm<%=l%>" href="<%=sURL%>?<%=toggleFolder(sLink)%>#itm<%=l%>"><%=Server.HTMLEncode(sText)%></a></li><%=buildFileList(sLink)%>
<!-- end folder <%=sLink%> -->
<%
	}// folder
}//end for l in arLinkText
%></ul>
<!-- end menu listing -->
<%} // an error did not occur%>
  A <img src="images/closed.gif" alt="Folder" /> image indicates a folder containing 
  documents.<br />
  Click the text to show them. </div>  
<div style="clear: both"> </div>
</div>
</body>
</html>