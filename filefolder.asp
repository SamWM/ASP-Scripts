<script language="JScript" runat="server">
/*
sFolder = which folder (e.g. '/folder/subfolder/', 'c:\folder\subfolder\')
return enumerator of files
*/
function getFiles(sFolder){ 
    // create filesystemobject
    var oFSO = Server.CreateObject("Scripting.FileSystemObject")
    // get folder object
    try{
        var oFolder = oFSO.getFolder((sFolder.indexOf("/")==-1)?sFolder:Server.MapPath(sFolder))
    }catch(e){
        // if error occurs return blank enumerator
        return new Enumerator()
    }
    // create enumerator for files
    var eFiles = new Enumerator(oFolder.Files)
    // return files list
    return eFiles
}

/*
sFolder = which folder (e.g. '/folder/subfolder/', 'c:\folder\subfolder\') 
returns enumerator of folders
*/
function getFolders(sFolder){ 
    // create filesystemobject
    var oFSO = Server.CreateObject("Scripting.FileSystemObject")
    // get folder object
    try{
        var oFolder = oFSO.getFolder((sFolder.indexOf("/")==-1)?sFolder:Server.MapPath(sFolder))
    }catch(e){
        // if error occurs return blank enumerator
        return new Enumerator()
    }
    // create enumerator for folders
    var eFolders = new Enumerator(oFolder.SubFolders)
    // return files list
    return eFolders
}

/*
sFolder = which folder (e.g. '/folder/subfolder/', 'c:\folder\subfolder\') 
returns date last modified
*/
function getFolderLastModified(sFolder){ 
    // create filesystemobject
    var oFSO = Server.CreateObject("Scripting.FileSystemObject")
    // get folder object
    try{
        var oFolder = oFSO.getFolder((sFolder.indexOf("/")==-1)?sFolder:Server.MapPath(sFolder))
    }catch(e){
        // if error occurs return 0
        return 0
    }
    // return folder modified date
    return oFolder.DateLastModified
}

/*
sFolder = which folder (e.g. '/folder/subfolder/', 'c:\folder\subfolder\') 
returns folder size
*/
function getFolderSize(sFolder){ 
    // create filesystemobject
    var oFSO = Server.CreateObject("Scripting.FileSystemObject")
    // get folder object
    try{
        var oFolder = oFSO.getFolder((sFolder.indexOf("/")==-1)?sFolder:Server.MapPath(sFolder))
    }catch(e){
        // if error occurs return 0
        return 0
    }
    // return folder modified date
    return oFolder.Size
}

/*
sPTF = path to file, bExt = include extension, default yes
returns filename as string
*/
function getFileName(sPTF,bExt){ 
    // new method using filesystemobject
    if(bExt!=false) bExt = true // if bExt is not supplied, the default is to return the file with the extension
    // create filesystemobject
    var oFSO = Server.CreateObject("Scripting.FileSystemObject")
    /*  
    get file object
    ---------------
    there is a method to get the file name by using oFSO.GetFileName(sPTF)
    but it does not check to see if the file exists
    this method returns blank string if the file does not exist
    */
    try{
        var oFile = oFSO.GetFile((sPTF.indexOf("/")==-1)?sPTF:Server.MapPath(sPTF))
        
    }catch(e){
        // if error occurs return blank string
        return ""
    }
    // get file name
    sFileName = oFile.Name
    // return dot?
    sFileName = (bExt)?sFileName:sFileName.substr(0,sFileName.lastIndexOf('.'))
    return sFileName.replace("&","&amp;")
    /* // old method using string manipulation
    iFNB = sPTF.lastIndexOf("\\")+1 // file name begins here (the last \ in the file path), 1 is added to the result because javascript is zero based
    // if iFNB returns 0 then the supplied path is an absolute one, it is therefor compensated for
    iFNB = (iFNB==0)?sPTF.lastIndexOf("/")+1:iFNB
    if(bExt){ // if we want the extension
        // substring begins at iFNB and ends at the end of the string
        return sPTF.substr(iFNB) // return file (with extension)
    }
    else{ // if we don't want the extension
        // substring begins at iFNB and ends at the position of the last dot
        return sPTF.substr(iFNB,sPTF.lastIndexOf('.')-iFNB) // remove extension - including '.'
    } // end if(bExt)
    */
}

/*
sPTF = path to file, bDot = include dot, default false
returns file extension as string
*/
function getFileExt(sPTF,bDot){
    if(bDot!=true) bDot = false // if bDot is not supplied, the default is to return the extension with no dot
    // return substring beginning before or after the final dot according to bDot
    return sPTF.substr(sPTF.lastIndexOf('.')+((!bDot)?1:0))
}

/*
sFN = filename
return filename without extension
*/
function rmFileExt(sFN){
	return sFN.substr(0, sFN.lastIndexOf("."))
}

/*
sPTF = path to file
sFmt = size format (b, kb, mb) - automatic if undefined
nDP = number of decimal places - default 2 if undefined
returns file size as string
*/
function getFileSize(sPTF, sFmt, nDP){
    // create filesystemobject
    var oFSO = Server.CreateObject("Scripting.FileSystemObject")
    // get file object
    try{
        var oFile = oFSO.GetFile((sPTF.indexOf("/")==-1)?sPTF:Server.MapPath(sPTF))
    }catch(e){
        // if error occurs return blank string
        return ""
    }
    // get file size
    var nFileSize = parseInt(oFile.size)
    sFmt = String(sFmt).toLowerCase()
    return fmtFileSize(nFileSize, sFmt, nDP)
}

function fmtFileSize(nFileSize, sFmt, nDP){
	if(isNaN(nDP)) nDP = 2
	var sAppend = ""
    if(sFmt!="b"||sFmt!="kb"||sFmt!="mb"){
        if(nFileSize<1000){
            var sFmt = "b"; var sAppend = " bytes";
        }
        else if(nFileSize>999 && nFileSize<1000000){
            var sFmt = "kb"; var sAppend = " kb";
        }
        else{
            var sFmt = "mb"; var sAppend = " mb";
        }
    }else{var sAppend = " "+sFmt}
    if(sFmt=="b") return nFileSize + sAppend
    if(sFmt=="kb"){
        // divide by 1024 to get the size in kb
        nFileSize = nFileSize / 1024;
        // multiply by 10 to the power of nDP
        nFileSize = nFileSize * Math.pow(10,nDP);
        // round up/down the new file size, and divide to remove the power and return the correct value
        nFileSize = Math.round(nFileSize) / Math.pow(10,nDP);
        return nFileSize + sAppend;
    }
    if(sFmt=="mb"){
        // divide by 1024 twice to get the size in mb
        nFileSize = nFileSize / 1024 / 1024;
        // multiply by 10 to the power of nDP
        nFileSize = nFileSize * Math.pow(10,nDP);
        // round up/down the new file size, and divide to remove the power and return the correct value
        nFileSize = Math.round(nFileSize) / Math.pow(10,nDP);
        return nFileSize + sAppend;
    }
}

/*
sPTF = path to file
returns last modified date of file as date string
*/
function getLastModified(sPTF){
    // create filesystemobject
    var oFSO = Server.CreateObject("Scripting.FileSystemObject")
    // get file object
    try{
        var oFile = oFSO.GetFile((sPTF.indexOf("/")==-1)?sPTF:Server.MapPath(sPTF))
    }catch(e){
        // if error occurs return blank string
        return ""
    }
    // get file last modified
    dtLastModified = new Date(oFile.DateLastModified)
    return dtLastModified.toLocaleString()
}

/*
sPTF = path to file
returns file type as string
*/
function getFileType(sPTF){
    // create filesystemobject
    var oFSO = Server.CreateObject("Scripting.FileSystemObject")
    // get file object
    try{
        var oFile = oFSO.GetFile((sPTF.indexOf("/")==-1)?sPTF:Server.MapPath(sPTF))
    }catch(e){
        // if error occurs return blank string
        return ""
    }
    // get file type
    return oFile.Type
}

/*
sFolder = which folder (e.g. 'c:\inetpub\www\path\subfolder\') 
sVP = virtual parent
returns virtual path ('/path/subfolder/')
*/
function getVirtualPath(sFolder,sVP){
    try{
        var sVPD = String(sFolder).replace(Server.MapPath("/"),"") // remove root directory
    }catch(e){
        // if error occurs return blank string
        return ""
    }
    /*
        if the mapped directory is stored in a directory not under the site directory, it needs to be removed
        it is removed by changing the mapped path of the parent directory to the virtual path
        for it to work, the virtual parent has to be supplied as an argument (sVP)
    */
    sVPD = (String(sVP!="undefined"))?sVPD.replace(Server.MapPath(sVP),sVP):sVPD
    sVPD = sVPD.replace(/\\/g,"/") // replace \ with /
    sVPD = sVPD.replace(/\s/g,"%20") // replace spaces with %20
    sVPD = sVPD.replace("&","&amp;") // replace & with &amp;
    return sVPD
}

/*
sVPD = virtual path to directory
*/
function getParentDirectory(sVPD){
    /*
        get the position of the last /, compensate for extra slash at end
        lastIndexOf is zero based, so 1 has to be added to the result
        if the last slash is equal to the length of sVPD it is removed and sVPD is updated
    */
    try{
        var sVPD = (sVPD.lastIndexOf("/")+1==sVPD.length)?sVPD.substr(0,sVPD.length-1):sVPD
        var nLastSlash = parseInt(sVPD.lastIndexOf("/"))
    }catch(e){
        // if error occurs return blank string
        return ""
    }
    // substring begins at 0 and ends at the final / in the path
    sVPD = sVPD.substr(0,nLastSlash)
    if(sVPD=="") return "/"
    return sVPD
}

/*
http://www.aspfaq.com/show.asp?id=2296
http://msdn.microsoft.com/library/en-us/shellcc/platform/shell/reference/objects/folder/getdetailsof.asp
*/
function getComment(sPTF){
    sPTF = (sPTF.indexOf("/")==-1)?sPTF:Server.MapPath(sPTF)
    var objShell = Server.CreateObject("Shell.Application");
    objFolder = objShell.NameSpace(sPTF.substr(0,sPTF.lastIndexOf("\\")+1));
    if (objFolder != null){
        objFolderItem = objFolder.ParseName(getFileName(sPTF,true));
        if (objFolderItem != null){
            return objFolder.GetDetailsOf(objFolderItem, 5);
        }
    }
}
function getTitle(sPTF){
    sPTF = (sPTF.indexOf("/")==-1)?sPTF:Server.MapPath(sPTF)
    var objShell = Server.CreateObject("Shell.Application");
    objFolder = objShell.NameSpace(sPTF.substr(0,sPTF.lastIndexOf("\\")+1));
    if (objFolder != null){
        objFolderItem = objFolder.ParseName(getFileName(sPTF,true));
        if (objFolderItem != null){
            return objFolder.GetDetailsOf(objFolderItem, 11);
        }
    }
}
</script>