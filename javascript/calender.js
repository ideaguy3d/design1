		
var weekend = [0,6];var weekendColor = "#e0e0e0";var sShowWeekend = 1;var gNow = new Date();var ggWinCal;var DAYS_OF_WEEK = 7;gNow = new Date(gNow.getFullYear(), gNow.getMonth(), gNow.getDate());
var DayOfWeek = new Array('S','M','T','W','T','F','S');
Calendar.Months = ["January", "February", "March", "April", "May", "June","July", "August", "September", "October", "November", "December"];
Calendar.DOMonth = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
Calendar.lDOMonth = [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];

function Calendar(p_item, p_WinCal, p_month, p_year, p_format) {
	if ((p_month == null) && (p_year == null))	return;
	if (p_WinCal == null)
		this.gWinCal = ggWinCal;
	else
		this.gWinCal = p_WinCal;

	if (p_month == null) {
		this.gMonthName = null;
		this.gMonth = null;
	} else {
		this.gMonthName = Calendar.get_month(p_month);
		this.gMonth = new Number(p_month);
	}

	var vUserDate = eval("window.document." + p_item + ".value");
	this.gYear = p_year;
	this.gFormat = p_format;
	this.gReturnItem = p_item;
	this.SelectedDate = vUserDate;
}

Calendar.get_month = Calendar_get_month;
Calendar.get_daysofmonth = Calendar_get_daysofmonth;
Calendar.calc_month_year = Calendar_calc_month_year;

function Calendar_get_month(monthNo) {
	return Calendar.Months[monthNo];
}

function Calendar_get_daysofmonth(monthNo, p_year) {
	if ((p_year % 4) == 0) {
		if ((p_year % 100) == 0 && (p_year % 400) != 0)	return Calendar.DOMonth[monthNo];
		return Calendar.lDOMonth[monthNo];
	} else
		return Calendar.DOMonth[monthNo];
}

function Calendar_calc_month_year(p_Month, p_Year, incr) {
var ret_arr = new Array();
if (incr == -1) {if (p_Month == 0) {ret_arr[0] = 11;	ret_arr[1] = parseInt(p_Year) - 1;}else {ret_arr[0] = parseInt(p_Month) - 1; ret_arr[1] = parseInt(p_Year);}}
else if (incr == 1) {if (p_Month == 11) {ret_arr[0] = 0;ret_arr[1] = parseInt(p_Year) + 1;}else {ret_arr[0] = parseInt(p_Month) + 1; ret_arr[1] = parseInt(p_Year);}}
return ret_arr;
}

new Calendar();

Calendar.prototype.getMonthlyCalendarCode = function() {
	var vCode = "";	var vHeader_Code = "";	var vData_Code = "";
	vHeader_Code = this.cal_header();
	vData_Code = this.cal_data();
	vCode = vCode + vHeader_Code + vData_Code + "</TABLE>";
	return vCode;
}
/*
function include_dom(script_filename) {

  
    // var script_filename= 'createAjaxObj.js' ;
    var html_doc = document.getElementsByTagName('head').item(0);
    var js = document.createElement('script');
    js.setAttribute('language', 'javascript');
    js.setAttribute('type', 'text/javascript');
    js.setAttribute('src', script_filename);
    html_doc.appendChild(js);
    return false;
}
*/

Calendar.prototype.show = function() {
	var vCode = "";	var a = p_item;
	this.gWinCal.document.open();
	this.wwrite("<html><head><META HTTP-EQUIV='PRAGMA' CONTENT='NO-CACHE'><title>Calendar</title>");
	this.wwrite("<style type='text/css'>" +
				"TABLE{ BACKGROUND-COLOR: White}" +
				"TD{ text-align: center; FONT-SIZE: 9pt;COLOR: #333333; font-family:arial;}" +
				".backclr {BACKGROUND-COLOR: #08416C}" +
				".month { font-family:arial; font-size: 10pt;FONT-WEIGHT: bold; color:#FFFFFF}" +
				".daysofweek {font-family:arial; font-size: 9pt; color:#000000}" +
				".otherday {FONT-SIZE: 9pt;COLOR: #CCCCCC;FONT-FAMILY: Arial, Helvetica, sans-serif;}" +
				"A.arrow {font-family:arial; font-size: 10pt;  color:#FFFFFF; text-decoration: none;}" +
				"A.day {FONT-SIZE: 9pt;COLOR: #000000;FONT-WEIGHT: bold; FONT-FAMILY: Arial, Helvetica, sans-serif;text-decoration: none;}" +				
				"A.day:hover {FONT-SIZE: 9pt;COLOR: #000000;FONT-WEIGHT: bold; BACKGROUND-COLOR: #CCCCCC; FONT-FAMILY: Arial, Helvetica, sans-serif;text-decoration: none;}" +
				"</style>" );   
	if (need_submit == true )
	{ 
	this.wwrite("<"+"SCR" + "IPT   src='createAjaxObj.js' type='text/javascript' defer><\/sc"+"ript"+">");
	}
	this.wwrite("</head>");
	this.wwrite("<body text='black' topmargin='12' leftmargin='5' marginwidth='5' marginheight='8'>");
	var prevMMYYYY = Calendar.calc_month_year(this.gMonth, this.gYear, -1);
	var prevMM = prevMMYYYY[0];
	var prevYYYY = prevMMYYYY[1];
	var vSpacer = "<img src='images/clear.gif' border='0' width='25' height='1'>";
	var nextMMYYYY = Calendar.calc_month_year(this.gMonth, this.gYear, 1);
	var nextMM = nextMMYYYY[0];
	var nextYYYY = nextMMYYYY[1];
	this.wwrite("<TABLE border='0' CELLSPACING='0' CELLPADDING='1'>");
	this.wwrite("<TR class='backclr'><td>" + vSpacer + "</td><td>" + vSpacer + "</td><td>" + vSpacer + "</td><td>" + vSpacer + "</td><td>" + vSpacer + "</td><td>" + vSpacer + "</td><td>" + vSpacer + "</td></tr>");
	this.wwrite("<TR class='backclr'>");
	( (this.gMonth <= gNow.getMonth() && this.gYear <= gNow.getFullYear()) ? this.wwrite("<TD><img src='images/clear.gif' border='0'></TD>") :
		this.wwrite("<TD><A class='arrow' HREF=\"" + "javascript:window.opener.Build('" + this.gReturnItem + "', '" + prevMM + "', '" + prevYYYY + "', '" + this.gFormat + "');" +
			"\"><<<\/A></TD>") );
	this.wwrite("<TD class='month' colspan='5'>" + this.gMonthName +"&nbsp;"+ this.gYear + "</TD>");
	this.wwrite("<TD><A class='arrow' HREF=\"" + "javascript:window.opener.Build('" + this.gReturnItem + "', '" + nextMM + "', '" + nextYYYY + "', '" + this.gFormat + "');" +
		"\">>><\/A></TD>");
	vCode = this.getMonthlyCalendarCode();
	this.wwrite(vCode);
	this.wwrite("<br>");
	this.wwrite("<center>");
	this.wwrite("<a class='daysofweek' href=\"javascript:self.close();\">Close<\/a>");
	this.wwrite("</center>");
	this.wwrite("</body></html>");
	this.gWinCal.document.close();
}

Calendar.prototype.wwrite = function(wtext) {
	this.gWinCal.document.writeln(wtext);
}

Calendar.prototype.wwriteA = function(wtext) {
	this.gWinCal.document.write(wtext);
}

Calendar.prototype.cal_header = function() {
	var vCode = "";
	for(index=0; index < DAYS_OF_WEEK ; index++)
	{
		vCode = vCode + "<TD class='daysofweek' height='10'>" + DayOfWeek[index] + "</TD>";
	}
	vCode = "<TR>"+ vCode + "</TR>";	
	vCode = vCode + "<TR><TD colspan='7' BGCOLOR='#CCCCCCC'><img src='images/clear.gif' height='1' border='0'></TD></TR>";
	return vCode;
}

Calendar.prototype.cal_data = function() {
	var vDate = new Date();
	vDate.setDate(1);
	vDate.setMonth(this.gMonth);
	vDate.setFullYear(this.gYear);
	var vFirstDay=vDate.getDay();
	var vDay=1;
	var vLastDay=Calendar.get_daysofmonth(this.gMonth, this.gYear);
	var vOnLastDay=0;
	var vCode = "";
	var vFlowDate = new Date();

	vCode = vCode + "<TR>";
	for (i=0; i<vFirstDay; i++) {
		vCode = vCode + "<TD ";	if (sShowWeekend != 1) vCode = vCode + this.write_weekend_string(i);
		vCode = vCode + ">&nbsp;</TD>";
	}

	for (j=vFirstDay; j<7; j++) {
		vFlowDate = new Date(this.gYear, this.gMonth, vDay);

		if (sShowWeekend == 1)
			{
			vCode = vCode + "<TD  "+ this.format_cell(vDay) + ">" ;
			if (vFlowDate < gNow){vCode = vCode + "<FONT COLOR='#808080'>" + vDay + "</FONT>";}
			else{vCode = vCode + "<A class='day' HREF='#' " +
					"onClick=\"self.opener.document." + this.gReturnItem + ".value='" +
					//this.format_data(vDay) + "';window.close(); "+((need_submit==true) ? "self.opener.document.basketform.submit();" : " ")+ " \">" + vDay +
					this.format_data(vDay) + "';window.close(); "+((need_submit==true) ? "javascript:Changeshipdate("+p_item_id+",'"+this.format_data(vDay)+"');" : " ")+ "\">" + vDay +
				"</A></TD>";}
			}
		else
			{

			vCode = vCode + "<TD "+ this.format_cell(vDay) + " " + this.write_weekend_string(j) + ">" + (this.isWeekend(j) ? (vDay +"</TD>") :
				( ( (vFlowDate < gNow) ? "<FONT COLOR='#808080'>" + vDay + "</FONT>" :
				"<A  class='day' HREF='#' " +
					"onClick=\"self.opener.document." + this.gReturnItem + ".value='" +
					//this.format_data(vDay) + "';window.close(); "+((need_submit==true) ? "self.opener.document.basketform.submit();" : " ")+ "\">" + vDay +
					this.format_data(vDay) + "';window.close(); "+((need_submit==true) ? "javascript:Changeshipdate("+p_item_id+",'"+this.format_data(vDay)+"');" : " ")+ "\">" +  vDay +
				"</A>") +
				"</TD>"));
		}
		vDay=vDay + 1;
	}
	vCode = vCode + "</TR>";

	for (k=2; k<7; k++) {
		vCode = vCode + "<TR>";

		for (j=0; j<7; j++) {
			vFlowDate = new Date(this.gYear, this.gMonth, vDay);
			if (sShowWeekend == 1)
				{
					vCode = vCode + "<TD "+ this.format_cell(vDay) +">";
					if(vFlowDate < gNow)
						{
							vCode = vCode + "<FONT COLOR='#808080'>" + vDay + "</FONT>";
						}
					else
						{
						vCode = vCode + "<A  class='day'  HREF='#' onClick=\"self.opener.document." + this.gReturnItem + ".value='" +
							//this.format_data(vDay) + "';window.close(); "+((need_submit==true) ? "self.opener.document.basketform.submit();" : " ")+"\">" + vDay + "</A>";
							this.format_data(vDay) + "';window.close(); "+((need_submit==true) ? "javascript:Changeshipdate("+p_item_id+",'"+this.format_data(vDay)+"');" : " ")+ "\">" + vDay + "</A>";
						}
					vCode = vCode + "</TD>";
				}

			else
				{
				vCode = vCode + "<TD " + this.format_cell(vDay) +" " + this.write_weekend_string(j) + ">" + (this.isWeekend(j) ? (vDay) :
				//( (vFlowDate < gNow) ? "<FONT COLOR='#808080'>" + vDay + "</FONT>" : "<A  class='day'  HREF='#' onClick=\"self.opener.document." + this.gReturnItem + ".value='" +	this.format_data(vDay) + "';window.close(); "+((need_submit==true) ? "self.opener.document.basketform.submit();" : " ")+"\">" + vDay +	"</A>"));
				( (vFlowDate < gNow) ? "<FONT COLOR='#808080'>" + vDay + "</FONT>" : "<A  class='day'  HREF='#' onClick=\"self.opener.document." + this.gReturnItem + ".value='" +	this.format_data(vDay) + "';window.close(); "+((need_submit==true) ? "javascript:Changeshipdate("+p_item_id+",'"+this.format_data(vDay)+"');" : " ")+ "\">" + vDay +	"</A>"));
				vCode = vCode + "</TD>";

				}
			vDay=vDay + 1;

			if (vDay > vLastDay)
			{
				vOnLastDay = 1;
				break;
			}
		}
		if (j == 6)
			vCode = vCode + "</TR>";
		if (vOnLastDay == 1)
			break;
	}
	for (m=1; m<(7-j); m++) {vCode = vCode + "<TD class='otherday'><img src='images/clear.gif' width='1' border='0'></TD>";}
	return vCode;
}

Calendar.prototype.format_cell = function(vday) {
	if ( (this.SelectedDate.length == 0) || (!validateUSDate(this.SelectedDate)) ) return "''" ;
	var vUSDate = new Date(this.SelectedDate);
	var vUSDay = vUSDate.getDate();
	var vUSMonth = vUSDate.getMonth();
	var vUSYear = vUSDate.getFullYear();
	if (vday == vUSDay && this.gMonth == vUSMonth && this.gYear == vUSYear)	return " BGCOLOR='#9D9D9D' ";
	else return ' ';
}
Calendar.prototype.write_weekend_string = function(vday) {
	var i;
	for (i=0; i<weekend.length; i++) {
		if (vday == weekend[i])
			return (" BGCOLOR=\"" + weekendColor + "\"");
	}
	return ' ';
}

Calendar.prototype.isWeekend = function(vday) {
	var i;
	for (i=0; i<weekend.length; i++) {
		if (vday == weekend[i])
			return true;
	}
	return false;
}


Calendar.prototype.format_data = function(p_day) {
	var vData;
	var vMonth = 1 + this.gMonth;
	vMonth = (vMonth.toString().length < 2) ? "0" + vMonth : vMonth;
	var vMon = Calendar.get_month(this.gMonth).substr(0,3).toUpperCase();
	var vFMon = Calendar.get_month(this.gMonth).toUpperCase();
	var vY4 = new String(this.gYear);
	var vY2 = new String(this.gYear.substr(2,2));
	var vDD = (p_day.toString().length < 2) ? "0" + p_day : p_day;

	vData = vMonth + "\/" + vDD + "\/" + vY4;
	return vData;
}

function Build(p_item, p_month, p_year, p_format) {
	var p_WinCal = ggWinCal;
	gCal = new Calendar(p_item, p_WinCal, p_month, p_year, p_format);
	gCal.show();
}

function show_calendar() {
/*
	sShowWeekend = arguments[0];
	p_item = arguments[1];
	if (arguments[2] == null)
		p_month = new String(gNow.getMonth());
	else
		p_month = arguments[3];
	if (arguments[3] == "" || arguments[3] == null)
		p_year = new String(gNow.getFullYear().toString());
	else
		p_year = arguments[3];
	if (arguments[3] == null)
		p_format = "MM/DD/YYYY";
	else
		p_format = arguments[4];
	*/
	sShowWeekend = 0 ;
	p_item = arguments[0];
	p_month = new String(gNow.getMonth());
	p_year = new String(gNow.getFullYear().toString());
	p_format = "MM/DD/YYYY";
	
	//p_string = new String(p_item) ;
	//p_basket_string = p_string.substr(0,10);
	
	//p_item_id =arguments[1]
	p_item_id =0 ;
	
	need_submit = false ;
	
	if (arguments[1] !=null)
	{
	    p_item_id =arguments[1];
	    need_submit = true ;
	}
	else{	    
	    need_submit = false ;
	}
	/*
	if (p_basket_string == "basketform")
	{
		need_submit= true ;
	}
	*/
	


	vWinCal = window.open("", "CalendarDemo", "width=200,height=195,status=no,resizable=no,top=200,left=200");
	//vWinCal = window.open("", "CalendarDemo", "width=500,height=500,status=no,resizable=no,top=200,left=200");
	vWinCal.opener = self;
	vWinCal.focus();
	ggWinCal = vWinCal;
	Build(p_item, p_month, p_year, p_format);
}

function validateUSDate( strValue ) {
var objRegExp = /^\d{1,2}(\-|\/|\.)\d{1,2}\1\d{4}$/ ;
if(!objRegExp.test(strValue))return false; else{
var strSeparator = strValue.substring(2,3); var arrayDate = strValue.split(strSeparator);
var arrayLookup = { '01' : 31,'03' : 31, '04' : 30,'05' : 31,'06' : 30,'07' : 31, '08' : 31,'09' : 30,'10' : 31,'11' : 30,'12' : 31}
var intDay = parseInt(arrayDate[1]); if(arrayLookup[arrayDate[0]] != null) {if(intDay <= arrayLookup[arrayDate[0]] && intDay != 0)return true;}var intYear = parseInt(arrayDate[2]); var intMonth = parseInt(arrayDate[0]); if( ((intYear % 4 == 0 && intDay <= 29) || (intYear % 4 != 0 && intDay <=28)) && intDay !=0) return true;}return false;}
