
// **********************************************************************************
// * Calling convention of these fucntions
// *  
// *  1.All the textControls must be in form of   formname.controlname,
// *   for compliance with netscape.     -- tree 10/25/2000
// **********************************************************************************

function ShowCalendar( TextControl, DatetoShow, YearControl, MonthControl, DayControl )
{
	if (DatetoShow == "")
	{
		var Today = new Date()
		DatetoShow = ( Today.getMonth() + 1 )  + "/" + Today.getDate() + "/" + Today.getYear();
	}
	var CalParam  = "/Common/calendar.asp?ctlDate=" + TextControl  + "&UserDate=" + escape(DatetoShow);
	CalParam +=  "&ctlYear=" + YearControl + "&ctlMonth=" + MonthControl + "&ctlDay="  + DayControl;
	popbox=window.open(CalParam,"Calendar","location=no,menubar=no,resizable=no,toolbar=no,scrollbars=no,directories=no,titlebar=no,status=no,width=240,height=240");			
		
}



function ShowCalendar1( TextControl, DatetoShow, rbooEditTime, rstrOnOkFunct )
{
	var lstrEditTimeFlag;

	if ( rbooEditTime == true )
		lstrEditTimeFlag = "&EditTimeFlag=true";
	else
		lstrEditTimeFlag = "&EditTimeFlag=false";

	if ( DatetoShow == "" )
	{
		var Today = new Date()
		DatetoShow = (Today.getMonth()+ 1)  + "/" + Today.getDate() + "/" + Today.getYear() + " 00:00:00";
	}

	var CalParam = "/Common/calendar.asp?ctlDate=" + TextControl  + 
		"&UserDate=" + escape(DatetoShow) + lstrEditTimeFlag;

	if ( rstrOnOkFunct != "" )
		CalParam = CalParam + "&OnOK=" + escape(rstrOnOkFunct);

	popbox = window.open( CalParam, "Calendar", "location=no, menubar=no, resizable=no, toolbar=no, " +
		"scrollbars=no, directories=no, titlebar=no, status=no, width=240, height=270" );
}



function ShowCalendar2( DatetoShow, YearControl, MonthControl, DayControl )
{
	if (DatetoShow == "")
	{
		var Today = new Date()
		DatetoShow = (Today.getMonth()+ 1)  + "/" + Today.getDate() + "/" + Today.getYear();
	}
	var CalParam  = "/Common/calendar.asp?UserDate=" + escape(DatetoShow);
	CalParam +=  "&ctlYear=" + YearControl + "&ctlMonth=" + MonthControl + "&ctlDay="  + DayControl;
	popbox = window.open( CalParam, "Calendar", "location=no, menubar=no, resizable=no, " + 
		"toolbar=no, scrollbars=no, directories=no, titlebar=no, status=no, width=240, height=240");			
		
}



function ShowCalendar3( DateControl, TimeControl, DateTimetoShow )
{
	var lstrEditTimeFlag;
	var DatetoShow;

	lstrEditTimeFlag = "&EditTimeFlag=true";

	if ( DatetoShow == "" )
	{
		var Today = new Date()
		DateTimetoShow = (Today.getMonth() + 1)  + "/" + Today.getDate() + "/" +
			Today.getYear() + " 00:00:00";
	}

	var CalParam  = "/Common/calendar.asp?ctlDate=" + DateControl  + 
		"&UserDate=" + escape(DateTimetoShow) + lstrEditTimeFlag + 
		"&ctlTime=" + TimeControl;

	popbox = window.open( CalParam, "Calendar", "location=no, menubar=no, resizable=no," +
		"toolbar=no, scrollbars=no, directories=no, titlebar=no, status=no, width=240, " +
		"height=270" );
}