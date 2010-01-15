<script language="JScript" runat="server">
// sURL = URL to append to, sParam = parameter name, sParamVal = parameter value
function appendToURL(sURL,sParam,sParamVal){
	// remove existing parameters
	sURL = removeFromURL(sURL,sParam)
	// remove ? if it is the last character on the URL
	sURL = sURL.replace(/\?$/,'')
	// generate the querystring (add '?' if querystring is blank, '&' if there are other parameters)
	sURL += ((sURL.indexOf("?") == -1)?"?":"&") + sParam + "=" + sParamVal;
	return sURL;
}
// sURL = URL to remove from, sParam = parameter to remove from URL
function removeFromURL(sURL,sParam) {
	// where the querystring starts (i.e. where ? appears)
	var qStart = sURL.indexOf('?') + 1;
	// if ? is not in sURL (i.e. qStart = 0) or ? is at the end of sURL, then there are no parameters to remove
	if (qStart == 0 || qStart == sURL.length-1) return sURL;
	/* create regular expression, match sParam followed by '=', then any character (except '&') 1 or more times
	and optionally ending with a single '&' (optional as there may be a parameter after this one, which will need to be preserved) */
	var regexp = new RegExp("(" + sParam + "=[^&]+(&)?)");
	/* return the URL, first by getting the root url (which ends at qStart),
	and then replacing any text that matches the regular expression, as well as making sure the url does not end with '&'
	*/
	return sURL.substring(0,qStart) + sURL.substring(qStart).replace(regexp,'').replace(/&$/,'');
}
</script>