<script language="JScript" runat="server">
function LZ(x) {return(x<0||x>9?'':'0')+x} // source from http://www.merlyn.demon.co.uk/js-date1.htm#WDTF
function formatDate(dtSup, sFmt, sSep){ // dtSup = supplied valid date; sFmt = time format: uk, us, iso; sSep = seperator (default is '-' for iso '/' for uk and us)
	if (dtSup==null || dtSup==undefined) return null
	if(sSep==null){
		if(sFmt=='iso') sSep = '-'
		else sSep = '/'
	}
	var d = new Date(dtSup)
	var dd = LZ(d.getDate())
	var MM = LZ(d.getMonth()+1)
	var yyyy = d.getFullYear()
	if(sFmt=='uk' || sFmt==null) return(dd+sSep+MM+sSep+yyyy)
	if(sFmt=='iso') return(yyyy+sSep+MM+sSep+dd)
	if(sFmt=='us') return(MM+sSep+dd+sSep+yyyy)
}

function formatTime(dtSup){ // dtSup = supplied valid date; sFmt = time format: uk, us, iso; sSep = seperator (default is '-' for iso '/' for uk and us)
	sSep = '/'
	var d = new Date(dtSup)
	var hh = LZ(d.getHours())
	var mm = LZ(d.getMinutes())
	var ss = d.getSeconds()
	return(hh+":"+mm+":"+ss)
}
function formatDate(dtSup, sFmt, sSep){ // dtSup = supplied valid date; sFmt = time format: uk, us, iso; sSep = seperator (default is '-' for iso '/' for uk and us)
	if(sSep==null){
		if(sFmt=='iso') sSep = '-'
		else sSep = '/'
	}
	var d = new Date(dtSup)
	var dd = LZ(d.getDate())
	var MM = LZ(d.getMonth()+1)
	var yyyy = d.getFullYear()
	if(sFmt=='uk' || sFmt==null) return(dd+sSep+MM+sSep+yyyy)
	if(sFmt=='iso') return(yyyy+sSep+MM+sSep+dd)
	if(sFmt=='us') return(MM+sSep+dd+sSep+yyyy)
}

function formatTime(dtSup){ // dtSup = supplied valid date; sFmt = time format: uk, us, iso; sSep = seperator (default is '-' for iso '/' for uk and us)
	sSep = '/'
	var d = new Date(dtSup)
	var hh = LZ(d.getHours())
	var mm = LZ(d.getMinutes())
	var ss = d.getSeconds()
	return(hh+":"+mm+":"+ss)
}

function dayDifference(dt1,dt2){ // dt1 = date 1 dt2 = date 2
	var one_day=1000*60*60*24
	dt1 = new Date(dt1)
	dt2 = new Date(dt2)
	return Math.ceil((dt1.getTime()-dt2.getTime())/one_day)
}

function validUKDate(dt){
	/*
	// regular expression found at: http://www.3leaf.com/default/NetRegExpRepository.aspx
	var regExpDate = /^(?:(?:0?[1-9]|1\d|2[0-8])(\/|-)(?:0?[1-9]|1[0-2]))(\/|-)(?:[1-9]\d\d\d|\d[1-9]\d\d|\d\d[1-9]\d|\d\d\d[1-9])$|^(?:(?:31(\/|-)(?:0?[13578]|1[02]))|(?:(?:29|30)(\/|-)(?:0?[1,3-9]|1[0-2])))(\/|-)(?:[1-9]\d\d\d|\d[1-9]\d\d|\d\d[1-9]\d|\d\d\d[1-9])$|^(29(\/|-)0?2)(\/|-)(?:(?:0[48]00|[13579][26]00|[2468][048]00)|(?:\d\d)?(?:0[48]|[2468][048]|[13579][26]))$/
	return regExpDate.test(dt)
	*/
	// method using date object: http://www.experts-exchange.com/Web/Web_Languages/JavaScript/Q_20504455.html#7918454
	var regExpDate = /^\d{1,2}\/\d{1,2}\/\d{4}$/
	// check supplied date against regular expression
	if(regExpDate.test(dt)){
		var dArr = String(dt).split("/");
		// day is the first item in the array, so it is item 0
		var dd = parseInt(dArr[0]);
		// month is the second item in the array, so it is item 1. Take away one as months start at 0 (January)
		var mm = parseInt(dArr[1])-1;
		// day is the last (third) item in the array, so it is item 2
		var yyyy = parseInt(dArr[2]);
		// generate a new date object based on the supplied date (dt)
		var d = new Date(yyyy,mm,dd);
		/* 
			check if the returned values for day (getDate), month (getMonth), and year (getFullYear)
			are equal to the parsed items above	as a date will be generated even with an invalid month/day/year
			returns true if they are all equal, false if any are not
		*/
		return d.getDate() == dd && d.getMonth() == mm && d.getFullYear() == yyyy
	}
	else{
		// date is not in the format dd/mm/yyyy so return false
		return false
	}
}

// dt = supplied date
function getMonthName(dt){
    try{
        arMonths = new Array("January","February","March","April","May","June","July","August","September","October","November","December");
        if(typeof(dt) == "number") return arMonths[dt-1]
		dt = new Date(dt);
        return arMonths[dt.getMonth()];
    }catch(e){return null}
}

// dt supplied date, n = how many months to add/remove?
function getAdjacentMonthName(dt,n){
    if(isNaN(parseInt(n))) n = 1
    try{
        arMonths = new Array("January","February","March","April","May","June","July","August","September","October","November","December");
		dt = new Date(dt);
        dt.setMonth(dt.getMonth()+n);
        return arMonths[dt.getMonth()];
    }catch(e){return null}
}

// dt supplied date, n = how many months to add/remove?
// returns date in array (year,month,day)
function getAdjacentMonth(dt,n){
    if(isNaN(parseInt(n))) n = 1
    try{
        dt = new Date(dt);
        dt.setMonth(dt.getMonth()+n);
        return new Array(dt.getFullYear(),dt.getMonth()+1,dt.getDate());
    }catch(e){return null }
}

// dt supplied date
function getDayName(dt){
    try{
        arDays = new Array("Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday");
		if(!isNaN(dt)) return arDays[dt]
		dt = new Date(dt);
        return arDays[dt.getDay()];
    }catch(e){return null}
}

function validTime(tm){
	// [01]?\d matches 0 or 1 (zero or one time) followed by a digit (0-9)
	// 2[0-3] matches 20 - 23
	// enclosed in brackets to capture the match (would match 0 - 23)
	// :[0-5]\d matches the time (00-59)
	var regExpTime = /^(([01]?\d|2[0-3]):[0-5]\d)$/
	return regExpTime.test(String(tm))
}
</script>