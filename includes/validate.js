/* Javascript Functions to Validate String, Number, Email and String with specified Characters */

var whitespace = " \t\n\r";

function isEmpty(s)  //Checks whether string is Empty
	{  
		return ((s == null) || (s.length == 0))
	}

/* Function added by Rama Raju on 28 July 2006 to check time sheet value in project time sheet */
function isTimeSheetValue(ControlVal)  //Checks whether it is a time sheet value
{  
		var regObj;
		var boolRet;
		regObj=new RegExp(/(^\d{1,2}\.[05]$)|(^\d{1,2}$)/);
		boolRet=regObj.test(ControlVal);
	    return(boolRet);
}
/* Function added on 2 August 2006 By ramaraju for trimming*/
function trim(str)
{
   return str.replace(/^\s*|\s*$/g,"");
}

function isGreaterDate(frmDate,toDate)
{
    //The function checks if toDate is greater than or equal to from date
	//Breaking up the start and End date into month, day, year formats
	var fromDatec =frmDate.split('/');
	var toDatec = toDate.split('/');

	if(eval(toDatec[2] > fromDatec[2]))
	    return true;
	else if(eval(toDatec[2] < fromDatec[2]))
	    return false;	
	else if(eval(toDatec[2] == fromDatec[2]))
	{
	   if (eval(toDatec[0] - fromDatec[0])>=1)
	   {
		     return true;
	   }
	   else if(eval(toDatec[0] - fromDatec[0])== 0)
	   {
		  if(eval(toDatec[1] -fromDatec[1]) > 0)
		      return true;
	      else if(eval(toDatec[1] - fromDatec[1]) <= 1)
		    { 
			   return false;
			}	 
	   }	
	   else if(eval(toDatec[0] - fromDatec[0]) < 1)
	   {
//	 	    alert('in month2 - month1'+eval(fromDatec[0] > toDatec[0]));
			return false;
	   }   	   	   	   
	}	
}
/* End of function to find wheter a give date is greater than equal to a date or not */


function isGreaterOrEqualDate(frmDate,toDate)
{
    //The function checks if toDate is greater than or equal to from date
	//Breaking up the start and End date into month, day, year formats
	var fromDatec =frmDate.split('/');
	var toDatec = toDate.split('/');

	if(eval(toDatec[2] > fromDatec[2]))
	    return true;
	else if(eval(toDatec[2] < fromDatec[2]))
	    return false;	
	else if(eval(toDatec[2] == fromDatec[2]))
	{
	   if (eval(toDatec[0] - fromDatec[0])>=1)
	   {
		     return true;
	   }
	   else if(eval(toDatec[0] - fromDatec[0])== 0)
	   {
		  if(eval(toDatec[1] -fromDatec[1]) >= 0)
		      return true;
	      else if(eval(toDatec[1] - fromDatec[1]) <= 1)
		    { 
			   return false;
			}	 
	   }	
	   else if(eval(toDatec[0] - fromDatec[0]) < 1)
	   {
//	 	    alert('in month2 - month1'+eval(fromDatec[0] > toDatec[0]));
			return false;
	   }   	   	   	   
	}	
}
/* End of function to find wheter a give date is greater than equal to a date or not */


	
function isWhitespace (s) //Checks for White spaces in the string 
	{  
		var i;
	 	if (isEmpty(s)) return true;
	 	   for (i = 0; i < s.length; i++)
	 	   {   
	 	       var c = s.charAt(i);
	 	       if (whitespace.indexOf(c) == -1) return false;
	 	   }
		   return true;
	}

function noWhitespace (s) //no white space in between the string
	{  
		var i;
	 	if (isEmpty(s)) return true;
	 	   for (i = 0; i < s.length; i++)
	 	   {   
	 	       var c = s.charAt(i);
				if (whitespace.indexOf(c) == 0){
			   	   return true;
				}
	 	   }
		   return false;
	}
function isSingleQuote(s) //Checks whether the number is numeric
	{  
	    var i;
	    var bag="'\''"
	    for (i = 0; i < s.length; i++)
	    {   
	        var c = s.charAt(i);
	        if (bag.indexOf(c) >=0) return true;
	    }
	    return false;
	}
function isNumber(s) //Checks whether the number is numeric
	{  
	    var i;
		var bag="1234567890"
	    for (i = 0; i < s.length; i++)
	    {   
	        var c = s.charAt(i);
	        if (bag.indexOf(c) == -1) return false;
	    }
	    return true;
	}

function isCharsInBag (s, bag) /* Checks whether the string only contains the characters 
								specified in the parameter bag */
	{  
	    var i;
	    for (i = 0; i < s.length; i++)
	    {   
	        var c = s.charAt(i);
	        if (bag.indexOf(c) == -1) return false;
	    }
	    return true;
	}

function isHours (s) //Checks whether the string contains only numbers and .
	{  
	    var i;
		var bag="1234567890."
	    for (i = 0; i < s.length; i++)
	    {   
	        var c = s.charAt(i);
	        if (bag.indexOf(c) == -1) return false;
	    }
	    return true;
	}
function isValidFileOrFolderName(s) //Checks whether the string contains only alphabets, . and _
	{  
	    var i;
		var bag="/\:*?''<>|"
	    for (i = 0; i < s.length; i++)
	    {   
	        var c = s.charAt(i);
	        if (bag.indexOf(c) > 0) return false;
	    }
	    return true;
	}	
function isString (s) //Checks whether the string contains only alphabets, . and _
	{  
	    var i;
		var bag="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz _-."
	    for (i = 0; i < s.length; i++)
	    {   
	        var c = s.charAt(i);
	        if (bag.indexOf(c) == -1) return false;
	    }
	    return true;
	}
function isAlpha (s) //Checks whether the string contains only alphabets
	{  
	    var i;
		var bag="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz "
	    for (i = 0; i < s.length; i++)
	    {   
	        var c = s.charAt(i);
	        if (bag.indexOf(c) == -1) return false;
	    }
	    return true;
	}
function isAlphaNumeric (s) //Checks whether the string contains only alphabets and Numbers
	{  
	    var i;
		var bag="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_-/ ";
	    for (i = 0; i < s.length; i++)
	    {   var c = s.charAt(i);
			if (bag.indexOf(c) == -1) return false;
		}
		return true;
	}
	
	function isFileName (s) //Checks whether the string contains only alphabets and Numbers
	{  
	    var i;
		var bag="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_";
	    for (i = 0; i < s.length; i++)
	    {   var c = s.charAt(i);
			if (bag.indexOf(c) == -1) return false;
		}
		return true;
	}
function isTextNumeric (s) //Checks whether the string contains only alphabets and Numbers
	{  
	    var i;
		var bag="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
	    for (i = 0; i < s.length; i++)
	    {   var c = s.charAt(i);
			if (bag.indexOf(c) == -1) return false;
		}
		return true;
	}
		
function isEmail (emailStr) //Checks whether the given string is a valid EMail address
 	{
		var emailPat=/^(.+)@(.+)$/
		var specialChars="\\(\\)<>@,;:\\\\\\\"\\.\\[\\]"
		var validChars="\[^\\s" + specialChars + "\]"
		var quotedUser="(\"[^\"]*\")"
		var ipDomainPat=/^\[(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})\]$/
		var atom=validChars + '+'
		var word="(" + atom + "|" + quotedUser + ")"
		var userPat=new RegExp("^" + word + "(\\." + word + ")*$")
		var domainPat=new RegExp("^" + atom + "(\\." + atom +")*$")

		var matchArray=emailStr.match(emailPat)

		if (matchArray==null) 
		{
			return false
		}
		
		var user=matchArray[1]
		var domain=matchArray[2]
		
		
		if (user.match(userPat)==null) 
		{
		    return false
		}
		
		var IPArray=domain.match(ipDomainPat)
		
		if (IPArray!=null) 
		{
		  	for (var i=1;i<=4;i++) 
		  		{
			    if (IPArray[i]>255) 
					{
			    		return false
			    	}
		    	}
		    return true
		}
	
		var domainArray=domain.match(domainPat)
		
		if (domainArray==null) 
			{
		    	return false
			}
		
		var atomPat=new RegExp(atom,"g")
		var domArr=domain.match(atomPat)
		var len=domArr.length

		if (domArr[domArr.length-1].length<2 || 
		    domArr[domArr.length-1].length>3) 
			{
			   return false
			}
		
		if (len<2) 
			{
			   var errStr="This address is missing a hostname!"
	   		   return false
			}

		return true;
	}
function isNumberHyphen(s) //Checks whether the number is numeric
	{  
	    var i;
		var bag="1234567890-"
	    for (i = 0; i < s.length; i++)
	    {   
	        var c = s.charAt(i);
	        if (bag.indexOf(c) == -1) return false;
	    }
	    return true;
	}
function isNumberHyphencommabrac(s) //Checks whether the number is numeric
	{  
	    var i;
		var bag="1234567890-,() "
	    for (i = 0; i < s.length; i++)
	    {   
	        var c = s.charAt(i);
	        if (bag.indexOf(c) == -1) return false;
	    }
	    return true;
	}


function isValidDate(dateStr) 	// Checks for the following valid date formats:
	{							// MM/DD/YY   MM/DD/YYYY   MM-DD-YY   MM-DD-YYYY
									
	
		var datePat = /^(\d{1,2})(\/|-)(\d{1,2})\2(\d{4})$/; // requires 4 digit year
		
		var matchArray = dateStr.match(datePat); // is the format ok?
		if (matchArray == null) 
		{
			alert(dateStr + " Date is not in a valid format.")
			return false;
		}
		
		// parse date into variables
		
		month = matchArray[1]; 
		day = matchArray[3];
		year = matchArray[4];
		
		if (month < 1 || month > 12) 
		{ // check month range
		  //	alert("Month must be between 1 and 12.");
			return false;
		}
		
		if (day < 1 || day > 31) 
		{	//check day range
		//	alert("Day must be between 1 and 31.");
			return false;
		}
		
		if ((month == 4 || month == 6 || month == 9 || month == 11) && day == 31) 
		{ //check whether the month has 31 days
			alert("Month "+month+" doesn't have 31 days!")
			return false;
		}
		if (month == 2) 
		{ // check for february 29th
			var isleap = (year % 4 == 0 && (year % 100 != 0 || year % 400 == 0)); //leap year check
			if (day > 29 || (day == 29 && !isleap)) 
			{
				alert("February " + year + " doesn't have " + day + " days!");
				return false;
		    }
		}
//included by saravanan for checking year
		if(year < 1900 || year > 2100)
		{
			alert('year is  not valid')
			return false;
		}		
		return true; 
	}	
/* End of Validation Java Scripts */
/* Function added on 25/03/2002 By lingaraju */
		
/* End of function to find wheter a give date is greater than equal to a date or not */	


/* fucntion added by saravanan to check whether end time is earlier than start date 
given the start noon and end noon by saravanan on 19/11/2004 */
function chkStartEndTime(startTime,startNoon,endTime,endNoon)
{
/*alert('startTime' + startTime)
alert('startNoon' + startNoon)
alert('endTime' + endTime) 
alert('endNoon' + endNoon)*/

	// start time is the time in integer 
	// startNoon is AM/PM
	
	// end time is the time in integer 
	// end Noon is AM/PM
	
	if(startNoon == 'AM')
	{
		var start = 1;
	}
	else
	{
		var start = 2;
	}
	
	if(endNoon == 'AM')
	{
		var end = 1;
	}
	else
	{
		var end = 2;
	}
	
	if(start==1)
	{
		if(end==2)
			{
				return true;
			}
		else
		{
			if(start==end)
			{
				if(parseInt(endTime)>=parseInt(startTime))
				{
					return true;
				}
				else
				{
					return false;
				}
			}
		}
	}
	else
	{
		if(end==1)
		{
			return false;
		}
		else
		{
			if(parseInt(endTime)>=parseInt(startTime))
			{
				return true;
			}
			else
			{
				return false;
			}
		}
	}
	return true;
			
	
	
	/*if(end == start)
	{
		if(parseInt(endTime) < parseInt(startTime))
		{
			return false;
		}
	}
	else
	{
		if(parseInt(endTime) >= parseInt(startTime))
		{
			return true;
		}	 
		else
		{
			return false;
		}
		
	}*/
}
/* end of function check end time */
/* popup function will open the url included by saravanan on 23/11/2004
function popup(URL) 
{
window.open(URL) 
//window.open(URL, toolbar=0,scrollbars=0,location=0,statusbar=0,menubar=0,resizable=0,width=400,height=300);
}
end of popup function*/

function compareDates(date1,dateformat1,date2,dateformat2) {
	var d1=getDateFromFormat(date1,dateformat1);
	var d2=getDateFromFormat(date2,dateformat2);
	if (d1==0 || d2==0) {
		return -1;
		}
	else if (d1 > d2) {
		return 1;
		}
	else if (d1==d2){
		return 0;	
		}
	return 0;
	}
	

// ------------------------------------------------------------------
// formatDate (date_object, format)
// Returns a date in the output format specified.
// The format string uses the same abbreviations as in getDateFromFormat()
// ------------------------------------------------------------------
function formatDate(date,format) {
	format=format+"";
	var result="";
	var i_format=0;
	var c="";
	var token="";
	var y=date.getYear()+"";
	var M=date.getMonth()+1;
	var d=date.getDate();
	var E=date.getDay();
	var H=date.getHours();
	var m=date.getMinutes();
	var s=date.getSeconds();
	var yyyy,yy,MMM,MM,dd,hh,h,mm,ss,ampm,HH,H,KK,K,kk,k;
	// Convert real date parts into formatted versions
	var value=new Object();
	if (y.length < 4) {y=""+(y-0+1900);}
	value["y"]=""+y;
	value["yyyy"]=y;
	value["yy"]=y.substring(2,4);
	value["M"]=M;
	value["MM"]=LZ(M);
	value["MMM"]=MONTH_NAMES[M-1];
	value["NNN"]=MONTH_NAMES[M+11];
	value["d"]=d;
	value["dd"]=LZ(d);
	value["E"]=DAY_NAMES[E+7];
	value["EE"]=DAY_NAMES[E];
	value["H"]=H;
	value["HH"]=LZ(H);
	if (H==0){value["h"]=12;}
	else if (H>12){value["h"]=H-12;}
	else {value["h"]=H;}
	value["hh"]=LZ(value["h"]);
	if (H>11){value["K"]=H-12;} else {value["K"]=H;}
	value["k"]=H+1;
	value["KK"]=LZ(value["K"]);
	value["kk"]=LZ(value["k"]);
	if (H > 11) { value["a"]="PM"; }
	else { value["a"]="AM"; }
	value["m"]=m;
	value["mm"]=LZ(m);
	value["s"]=s;
	value["ss"]=LZ(s);
	while (i_format < format.length) {
		c=format.charAt(i_format);
		token="";
		while ((format.charAt(i_format)==c) && (i_format < format.length)) {
			token += format.charAt(i_format++);
			}
		if (value[token] != null) { result=result + value[token]; }
		else { result=result + token; }
		}
	return result;
	}
	
// ------------------------------------------------------------------
// Utility functions for parsing in getDateFromFormat()
// ------------------------------------------------------------------
function _isInteger(val) {
	var digits="1234567890";
	for (var i=0; i < val.length; i++) {
		if (digits.indexOf(val.charAt(i))==-1) { return false; }
		}
	return true;
	}
function _getInt(str,i,minlength,maxlength) {
	for (var x=maxlength; x>=minlength; x--) {
		var token=str.substring(i,i+x);
		if (token.length < minlength) { return null; }
		if (_isInteger(token)) { return token; }
		}
	return null;
	}
	
// ------------------------------------------------------------------
// getDateFromFormat( date_string , format_string )
//
// This function takes a date string and a format string. It matches
// If the date string matches the format string, it returns the 
// getTime() of the date. If it does not match, it returns 0.
// ------------------------------------------------------------------
function getDateFromFormat(val,format) {
	val=val+"";
	format=format+"";
	var i_val=0;
	var i_format=0;
	var c="";
	var token="";
	var token2="";
	var x,y;
	var now=new Date();
	var year=now.getYear();
	var month=now.getMonth()+1;
	var date=1;
	var hh=now.getHours();
	var mm=now.getMinutes();
	var ss=now.getSeconds();
	var ampm="";
	
	while (i_format < format.length) {
		// Get next token from format string
		c=format.charAt(i_format);
		token="";
		while ((format.charAt(i_format)==c) && (i_format < format.length)) {
			token += format.charAt(i_format++);
			}
		// Extract contents of value based on format token
		if (token=="yyyy" || token=="yy" || token=="y") {
			if (token=="yyyy") { x=4;y=4; }
			if (token=="yy")   { x=2;y=2; }
			if (token=="y")    { x=2;y=4; }
			year=_getInt(val,i_val,x,y);
			if (year==null) { return 0; }
			i_val += year.length;
			if (year.length==2) {
				if (year > 70) { year=1900+(year-0); }
				else { year=2000+(year-0); }
				}
			}
		else if (token=="MMM"||token=="NNN"){
			month=0;
			for (var i=0; i<MONTH_NAMES.length; i++) {
				var month_name=MONTH_NAMES[i];
				if (val.substring(i_val,i_val+month_name.length).toLowerCase()==month_name.toLowerCase()) {
					if (token=="MMM"||(token=="NNN"&&i>11)) {
						month=i+1;
						if (month>12) { month -= 12; }
						i_val += month_name.length;
						break;
						}
					}
				}
			if ((month < 1)||(month>12)){return 0;}
			}
		else if (token=="EE"||token=="E"){
			for (var i=0; i<DAY_NAMES.length; i++) {
				var day_name=DAY_NAMES[i];
				if (val.substring(i_val,i_val+day_name.length).toLowerCase()==day_name.toLowerCase()) {
					i_val += day_name.length;
					break;
					}
				}
			}
		else if (token=="MM"||token=="M") {
			month=_getInt(val,i_val,token.length,2);
			if(month==null||(month<1)||(month>12)){return 0;}
			i_val+=month.length;}
		else if (token=="dd"||token=="d") {
			date=_getInt(val,i_val,token.length,2);
			if(date==null||(date<1)||(date>31)){return 0;}
			i_val+=date.length;}
		else if (token=="hh"||token=="h") {
			hh=_getInt(val,i_val,token.length,2);
			if(hh==null||(hh<1)||(hh>12)){return 0;}
			i_val+=hh.length;}
		else if (token=="HH"||token=="H") {
			hh=_getInt(val,i_val,token.length,2);
			if(hh==null||(hh<0)||(hh>23)){return 0;}
			i_val+=hh.length;}
		else if (token=="KK"||token=="K") {
			hh=_getInt(val,i_val,token.length,2);
			if(hh==null||(hh<0)||(hh>11)){return 0;}
			i_val+=hh.length;}
		else if (token=="kk"||token=="k") {
			hh=_getInt(val,i_val,token.length,2);
			if(hh==null||(hh<1)||(hh>24)){return 0;}
			i_val+=hh.length;hh--;}
		else if (token=="mm"||token=="m") {
			mm=_getInt(val,i_val,token.length,2);
			if(mm==null||(mm<0)||(mm>59)){return 0;}
			i_val+=mm.length;}
		else if (token=="ss"||token=="s") {
			ss=_getInt(val,i_val,token.length,2);
			if(ss==null||(ss<0)||(ss>59)){return 0;}
			i_val+=ss.length;}
		else if (token=="a") {
			if (val.substring(i_val,i_val+2).toLowerCase()=="am") {ampm="AM";}
			else if (val.substring(i_val,i_val+2).toLowerCase()=="pm") {ampm="PM";}
			else {return 0;}
			i_val+=2;}
		else {
			if (val.substring(i_val,i_val+token.length)!=token) {return 0;}
			else {i_val+=token.length;}
			}
		}
	// If there are any trailing characters left in the value, it doesn't match
	if (i_val != val.length) { return 0; }
	// Is date valid for month?
	if (month==2) {
		// Check for leap year
		if ( ( (year%4==0)&&(year%100 != 0) ) || (year%400==0) ) { // leap year
			if (date > 29){ return 0; }
			}
		else { if (date > 28) { return 0; } }
		}
	if ((month==4)||(month==6)||(month==9)||(month==11)) {
		if (date > 30) { return 0; }
		}
	// Correct hours value
	if (hh<12 && ampm=="PM") { hh=hh-0+12; }
	else if (hh>11 && ampm=="AM") { hh-=12; }
	var newdate=new Date(year,month-1,date,hh,mm,ss);
	return newdate.getTime();
	}

// ------------------------------------------------------------------
// parseDate( date_string [, prefer_euro_format] )
//
// This function takes a date string and tries to match it to a
// number of possible date formats to get the value. It will try to
// match against the following international formats, in this order:
// y-M-d   MMM d, y   MMM d,y   y-MMM-d   d-MMM-y  MMM d
// M/d/y   M-d-y      M.d.y     MMM-d     M/d      M-d
// d/M/y   d-M-y      d.M.y     d-MMM     d/M      d-M
// A second argument may be passed to instruct the method to search
// for formats like d/M/y (european format) before M/d/y (American).
// Returns a Date object or null if no patterns match.
// ------------------------------------------------------------------
function parseDate(val) {
	var preferEuro=(arguments.length==2)?arguments[1]:false;
	generalFormats=new Array('y-M-d','MMM d, y','MMM d,y','y-MMM-d','d-MMM-y','MMM d');
	monthFirst=new Array('M/d/y','M-d-y','M.d.y','MMM-d','M/d','M-d');
	dateFirst =new Array('d/M/y','d-M-y','d.M.y','d-MMM','d/M','d-M');
	var checkList=new Array('generalFormats',preferEuro?'dateFirst':'monthFirst',preferEuro?'monthFirst':'dateFirst');
	var d=null;
	for (var i=0; i<checkList.length; i++) {
		var l=window[checkList[i]];
		for (var j=0; j<l.length; j++) {
			d=getDateFromFormat(val,l[j]);
			if (d!=0) { return new Date(d); }
			}
		}
	return null;
	}
