'Author : 					Doug Wilson
'Email:						doug@ix.co.za
'ICQ:						173155300
'Date : 		 			12 January 2003
'Application :  			iX Enterprise
'Module : 					Controls.DatePicker
'Description : 	    		DatePicker Web Control
'Updated By:
'Upd On:
'Reason:
'Email:
'ICQ:

Namespace iX.Controls

    Public Class DatePicker
        Inherits Control
        Implements INamingContainer

#Region "Members"

        Dim m_closeGif As String = "close.gif"
        Dim m_dividerGif As String = "divider.gif"
        Dim m_drop1Gif As String = "drop1.gif"
        Dim m_drop2Gif As String = "drop2.gif"
        Dim m_left1Gif As String = "left1.gif"
        Dim m_left2Gif As String = "left2.gif"
        Dim m_right1Gif As String = "right1.gif"
        Dim m_right2Gif As String = "right2.gif"
        Dim m_imgDirectory As String = "/img/"
        Dim m_ControlCssClass As String = ""
        Dim m_text As String = ""
        Dim m_Css As String = ""
        Dim m_DateType As String = "dd mmm yyyy"


#End Region
#Region "Properties"

        Public Property Text() As String
            Get
                If (Me.Controls.Count = 0) Then
                    Return ""
                End If
                Return CType(Controls(0), System.Web.UI.WebControls.TextBox).Text
            End Get
            Set(ByVal Value As String)
                If Not (Me.Controls.Count = 0) Then
                    CType(Controls(0), System.Web.UI.WebControls.TextBox).Text = Value
                    m_text = ""
                Else
                    m_text = Value
                End If
            End Set
        End Property
        Public Property CssClass() As String
            Get
                If (Me.Controls.Count = 0) Then
                    Return ""
                End If
                Return CType(Controls(0), System.Web.UI.WebControls.TextBox).CssClass
            End Get
            Set(ByVal Value As String)
                If Not (Me.Controls.Count = 0) Then
                    CType(Controls(0), System.Web.UI.WebControls.TextBox).CssClass = Value
                    m_Css = ""
                Else
                    m_Css = Value
                End If
            End Set
        End Property
        Public Property imgClose() As String
            Get
                Return m_closeGif
            End Get
            Set(ByVal Value As String)
                m_closeGif = Value
            End Set
        End Property
        Public Property imgDivider() As String
            Get
                Return m_dividerGif
            End Get
            Set(ByVal Value As String)
                m_dividerGif = Value
            End Set
        End Property
        Public Property imgDrop1() As String
            Get
                Return m_drop1Gif
            End Get
            Set(ByVal Value As String)
                m_drop1Gif = Value
            End Set
        End Property
        Public Property imgDrop2() As String
            Get
                Return m_drop2Gif
            End Get
            Set(ByVal Value As String)
                m_drop2Gif = Value
            End Set
        End Property
        Public Property imgLeft1() As String
            Get
                Return m_left1Gif
            End Get
            Set(ByVal Value As String)
                m_left1Gif = Value
            End Set
        End Property
        Public Property imgLeft2() As String
            Get
                Return m_left2Gif
            End Get
            Set(ByVal Value As String)
                m_left2Gif = Value
            End Set
        End Property
        Public Property imgRight1() As String
            Get
                Return m_right1Gif
            End Get
            Set(ByVal Value As String)
                m_right1Gif = Value
            End Set
        End Property
        Public Property imgRight2() As String
            Get
                Return m_right2Gif
            End Get
            Set(ByVal Value As String)
                m_right2Gif = Value
            End Set
        End Property
        Public Property imgDirectory() As String
            Get
                Return m_imgDirectory
            End Get
            Set(ByVal Value As String)
                m_imgDirectory = Value
            End Set
        End Property
        Public Property ControlCssClass() As String
            Get
                Return m_ControlCssClass
            End Get
            Set(ByVal Value As String)
                m_ControlCssClass = Value
            End Set
        End Property
        Public Property DateType() As String
            Get
                Return m_DateType
            End Get
            Set(ByVal Value As String)
                m_DateType = Value
            End Set
        End Property

#End Region

        Private Sub placeJavascript()
            Dim strBuildUp As String = "<script language=JavaScript>"
            strBuildUp += "//	written	by Tan Ling	Wee	on 2 Dec 2001" & Chr(13)
            strBuildUp += "//	last updated 23 June 2002" & Chr(13)
            strBuildUp += "//	email :	fuushikaden@yahoo.com" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "var	fixedX = -1			// x position (-1 if to appear below control)" & Chr(13)
            strBuildUp += "var	fixedY = -1			// y position (-1 if to appear below control)" & Chr(13)
            strBuildUp += "var startAt = 1			// 0 - sunday ; 1 - monday" & Chr(13)
            strBuildUp += "var showWeekNumber = 1	// 0 - don't show; 1 - show" & Chr(13)
            strBuildUp += "var showToday = 1		// 0 - don't show; 1 - show" & Chr(13)
            strBuildUp += "var imgDir = """ & m_imgDirectory & """" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "var gotoString = ""Go To Current Month""" & Chr(13)
            strBuildUp += "var todayString = ""Today is""" & Chr(13)
            strBuildUp += "var weekString = ""Wk""" & Chr(13)
            strBuildUp += "var scrollLeftMessage = ""Click to scroll to previous month. Hold mouse button to scroll automatically.""" & Chr(13)
            strBuildUp += "var scrollRightMessage = ""Click to scroll to next month. Hold mouse button to scroll automatically.""" & Chr(13)
            strBuildUp += "var selectMonthMessage = ""Click to select a month.""" & Chr(13)
            strBuildUp += "var selectYearMessage = ""Click to select a year.""" & Chr(13)
            strBuildUp += "var selectDateMessage = ""Select [date] as date."" // do not replace [date], it will be replaced by date." & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "var	crossobj, crossMonthObj, crossYearObj, monthSelected, yearSelected, dateSelected, omonthSelected, oyearSelected, odateSelected, monthConstructed, yearConstructed, intervalID1, intervalID2, timeoutID1, timeoutID2, ctlToPlaceValue, ctlNow, dateFormat, nStartingYear" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "var	bPageLoaded=false" & Chr(13)
            strBuildUp += "var	ie=document.all" & Chr(13)
            strBuildUp += "var	dom=document.getElementById" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "var	ns4=document.layers" & Chr(13)
            strBuildUp += "var	today =	new	Date()" & Chr(13)
            strBuildUp += "var	dateNow	 = today.getDate()" & Chr(13)
            strBuildUp += "var	monthNow = today.getMonth()" & Chr(13)
            strBuildUp += "var	yearNow	 = today.getYear()" & Chr(13)
            strBuildUp += "var	imgsrc = new Array(""" & m_drop1Gif & """,""" & m_drop2Gif & """,""" & m_left1Gif & """,""" & m_left2Gif & """,""" & m_right1Gif & """,""" & m_right2Gif & """)" & Chr(13)
            strBuildUp += "var	img	= new Array()" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "var bShow = false;" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "/* hides <select> and <applet> objects (for IE only) */" & Chr(13)
            strBuildUp += "function hideElement( elmID, overDiv )" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "if( ie )" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "for( i = 0; i < document.all.tags( elmID ).length; i++ )" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "obj = document.all.tags( elmID )[i];" & Chr(13)
            strBuildUp += "if( !obj || !obj.offsetParent )" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "continue;" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "// Find the element's offsetTop and offsetLeft relative to the BODY tag." & Chr(13)
            strBuildUp += "objLeft   = obj.offsetLeft;" & Chr(13)
            strBuildUp += "objTop    = obj.offsetTop;" & Chr(13)
            strBuildUp += "objParent = obj.offsetParent;" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "while( objParent.tagName.toUpperCase() != ""BODY"" )" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "objLeft  += objParent.offsetLeft;" & Chr(13)
            strBuildUp += "objTop   += objParent.offsetTop;" & Chr(13)
            strBuildUp += "objParent = objParent.offsetParent;" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "objHeight = obj.offsetHeight;" & Chr(13)
            strBuildUp += "objWidth = obj.offsetWidth;" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "if(( overDiv.offsetLeft + overDiv.offsetWidth ) <= objLeft );" & Chr(13)
            strBuildUp += "else if(( overDiv.offsetTop + overDiv.offsetHeight ) <= objTop );" & Chr(13)
            strBuildUp += "else if( overDiv.offsetTop >= ( objTop + objHeight ));" & Chr(13)
            strBuildUp += "else if( overDiv.offsetLeft >= ( objLeft + objWidth ));" & Chr(13)
            strBuildUp += "else" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "obj.style.visibility = ""hidden"";" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "/*" & Chr(13)
            strBuildUp += "* unhides <select> and <applet> objects (for IE only)" & Chr(13)
            strBuildUp += "*/" & Chr(13)
            strBuildUp += "function showElement( elmID )" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "if( ie )" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "for( i = 0; i < document.all.tags( elmID ).length; i++ )" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "obj = document.all.tags( elmID )[i];" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "if( !obj || !obj.offsetParent )" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "continue;" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "obj.style.visibility = """";" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "function HolidayRec (d, m, y, desc)" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "this.d = d" & Chr(13)
            strBuildUp += "this.m = m" & Chr(13)
            strBuildUp += "this.y = y" & Chr(13)
            strBuildUp += "this.desc = desc" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "var HolidaysCounter = 0" & Chr(13)
            strBuildUp += "var Holidays = new Array()" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "function addHoliday (d, m, y, desc)" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "Holidays[HolidaysCounter++] = new HolidayRec ( d, m, y, desc )" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "if (dom)" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "for	(i=0;i<imgsrc.length;i++)" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "img[i] = new Image" & Chr(13)
            strBuildUp += "img[i].src = imgDir + imgsrc[i]" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "document.write (""<div onclick='bShow=true' id='calendar'	style='z-index:+999;position:absolute;visibility:hidden;'><table	width=""+((showWeekNumber==1)?250:220)+"" style='font-family:arial;font-size:11px;border-width:1;border-style:solid;border-color:#a0a0a0;font-family:arial; font-size:11px}' bgcolor='#ffffff'><tr bgcolor='#0000aa'><td><table width='""+((showWeekNumber==1)?248:218)+""'><tr><td style='padding:2px;font-family:arial; font-size:11px;'><font color='#ffffff'><B><span id='caption'></span></B></font></td><td align=right><a href='javascript:hideCalendar()'><IMG SRC='""+imgDir+""" & m_closeGif & "' WIDTH='15' HEIGHT='13' BORDER='0' ALT='Close the Calendar'></a></td></tr></table></td></tr><tr><td style='padding:5px' bgcolor=#ffffff><span id='content'></span></td></tr>"")" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "if (showToday==1)" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "document.write (""<tr bgcolor=#f0f0f0><td style='padding:5px' align=center><span id='lblToday'></span></td></tr>"")" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "document.write (""</table></div><div id='selectMonth' style='z-index:+999;position:absolute;visibility:hidden;'></div><div id='selectYear' style='z-index:+999;position:absolute;visibility:hidden;'></div>"");" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "var	monthName =	new	Array(""January"",""February"",""March"",""April"",""May"",""June"",""July"",""August"",""September"",""October"",""November"",""December"")" & Chr(13)
            strBuildUp += "if (startAt==0)" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "dayName = new Array	(""Sun"",""Mon"",""Tue"",""Wed"",""Thu"",""Fri"",""Sat"")" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "else" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "dayName = new Array	(""Mon"",""Tue"",""Wed"",""Thu"",""Fri"",""Sat"",""Sun"")" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "var	styleAnchor=""text-decoration:none;color:black;""" & Chr(13)
            strBuildUp += "var	styleLightBorder=""border-style:solid;border-width:1px;border-color:#a0a0a0;""" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "function swapImage(srcImg, destImg){" & Chr(13)
            strBuildUp += "if (ie)	{ document.getElementById(srcImg).setAttribute(""src"",imgDir + destImg) }" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "function init()	{" & Chr(13)
            strBuildUp += "if (!ns4)" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "if (!ie) { yearNow += 1900	}" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "crossobj=(dom)?document.getElementById(""calendar"").style : ie? document.all.calendar : document.calendar" & Chr(13)
            strBuildUp += "hideCalendar()" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "crossMonthObj=(dom)?document.getElementById(""selectMonth"").style : ie? document.all.selectMonth	: document.selectMonth" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "crossYearObj=(dom)?document.getElementById(""selectYear"").style : ie? document.all.selectYear : document.selectYear" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "monthConstructed=false;" & Chr(13)
            strBuildUp += "yearConstructed=false;" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "if (showToday==1)" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "document.getElementById(""lblToday"").innerHTML =	todayString + "" <a onmousemove='window.status=\""""+gotoString+""\""' onmouseout='window.status=\""\""' title='""+gotoString+""' style='""+styleAnchor+""' href='javascript:monthSelected=monthNow;yearSelected=yearNow;constructCalendar();'>""+dayName[(today.getDay()-startAt==-1)?6:(today.getDay()-startAt)]+"", "" + dateNow + "" "" + monthName[monthNow].substring(0,3)	+ ""	"" +	yearNow	+ ""</a>""" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "sHTML1=""<span id='spanLeft'	style='border-style:solid;border-width:1;border-color:#3366FF;cursor:pointer' onmouseover='swapImage(\""changeLeft\"",\""left2.gif\"");this.style.borderColor=\""#88AAFF\"";window.status=\""""+scrollLeftMessage+""\""' onclick='javascript:decMonth()' onmouseout='clearInterval(intervalID1);swapImage(\""changeLeft\"",\""left1.gif\"");this.style.borderColor=\""#3366FF\"";window.status=\""\""' onmousedown='clearTimeout(timeoutID1);timeoutID1=setTimeout(\""StartDecMonth()\"",500)'	onmouseup='clearTimeout(timeoutID1);clearInterval(intervalID1)'>&nbsp<IMG id='changeLeft' SRC='""+imgDir+""left1.gif' width=10 height=11 BORDER=0>&nbsp</span>&nbsp;""" & Chr(13)
            strBuildUp += "sHTML1+=""<span id='spanRight' style='border-style:solid;border-width:1;border-color:#3366FF;cursor:pointer'	onmouseover='swapImage(\""changeRight\"",\""right2.gif\"");this.style.borderColor=\""#88AAFF\"";window.status=\""""+scrollRightMessage+""\""' onmouseout='clearInterval(intervalID1);swapImage(\""changeRight\"",\""right1.gif\"");this.style.borderColor=\""#3366FF\"";window.status=\""\""' onclick='incMonth()' onmousedown='clearTimeout(timeoutID1);timeoutID1=setTimeout(\""StartIncMonth()\"",500)'	onmouseup='clearTimeout(timeoutID1);clearInterval(intervalID1)'>&nbsp<IMG id='changeRight' SRC='""+imgDir+""right1.gif'	width=10 height=11 BORDER=0>&nbsp</span>&nbsp""" & Chr(13)
            strBuildUp += "sHTML1+=""<span id='spanMonth' style='border-style:solid;border-width:1;border-color:#3366FF;cursor:pointer'	onmouseover='swapImage(\""changeMonth\"",\""drop2.gif\"");this.style.borderColor=\""#88AAFF\"";window.status=\""""+selectMonthMessage+""\""' onmouseout='swapImage(\""changeMonth\"",\""drop1.gif\"");this.style.borderColor=\""#3366FF\"";window.status=\""\""' onclick='popUpMonth()'></span>&nbsp;""" & Chr(13)
            strBuildUp += "sHTML1+=""<span id='spanYear' style='border-style:solid;border-width:1;border-color:#3366FF;cursor:pointer' onmouseover='swapImage(\""changeYear\"",\""drop2.gif\"");this.style.borderColor=\""#88AAFF\"";window.status=\""""+selectYearMessage+""\""'	onmouseout='swapImage(\""changeYear\"",\""drop1.gif\"");this.style.borderColor=\""#3366FF\"";window.status=\""\""'	onclick='popUpYear()'></span>&nbsp;""" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "document.getElementById(""caption"").innerHTML  =	sHTML1" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "bPageLoaded=true" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "function hideCalendar()	{" & Chr(13)
            strBuildUp += "crossobj.visibility=""hidden""" & Chr(13)
            strBuildUp += "if (crossMonthObj != null){crossMonthObj.visibility=""hidden""}" & Chr(13)
            strBuildUp += "if (crossYearObj !=	null){crossYearObj.visibility=""hidden""}" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "showElement( 'SELECT' );" & Chr(13)
            strBuildUp += "showElement( 'APPLET' );" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "function padZero(num) {" & Chr(13)
            strBuildUp += "return (num	< 10)? '0' + num : num ;" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "function constructDate(d,m,y)" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "sTmp = dateFormat" & Chr(13)
            strBuildUp += "sTmp = sTmp.replace	(""dd"",""<e>"")" & Chr(13)
            strBuildUp += "sTmp = sTmp.replace	(""d"",""<d>"")" & Chr(13)
            strBuildUp += "sTmp = sTmp.replace	(""<e>"",padZero(d))" & Chr(13)
            strBuildUp += "sTmp = sTmp.replace	(""<d>"",d)" & Chr(13)
            strBuildUp += "sTmp = sTmp.replace	(""mmm"",""<o>"")" & Chr(13)
            strBuildUp += "sTmp = sTmp.replace	(""mm"",""<n>"")" & Chr(13)
            strBuildUp += "sTmp = sTmp.replace	(""m"",""<m>"")" & Chr(13)
            strBuildUp += "sTmp = sTmp.replace	(""<m>"",m+1)" & Chr(13)
            strBuildUp += "sTmp = sTmp.replace	(""<n>"",padZero(m+1))" & Chr(13)
            strBuildUp += "sTmp = sTmp.replace	(""<o>"",monthName[m])" & Chr(13)
            strBuildUp += "return sTmp.replace (""yyyy"",y)" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "function closeCalendar() {" & Chr(13)
            strBuildUp += "var	sTmp" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "hideCalendar();" & Chr(13)
            strBuildUp += "ctlToPlaceValue.value =	constructDate(dateSelected,monthSelected,yearSelected)" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "/*** Month Pulldown	***/" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "function StartDecMonth()" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "intervalID1=setInterval(""decMonth()"",80)" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "function StartIncMonth()" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "intervalID1=setInterval(""incMonth()"",80)" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "function incMonth () {" & Chr(13)
            strBuildUp += "monthSelected++" & Chr(13)
            strBuildUp += "if (monthSelected>11) {" & Chr(13)
            strBuildUp += "monthSelected=0" & Chr(13)
            strBuildUp += "yearSelected++" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "constructCalendar()" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "function decMonth () {" & Chr(13)
            strBuildUp += "monthSelected--" & Chr(13)
            strBuildUp += "if (monthSelected<0) {" & Chr(13)
            strBuildUp += "monthSelected=11" & Chr(13)
            strBuildUp += "yearSelected--" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "constructCalendar()" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "function constructMonth() {" & Chr(13)
            strBuildUp += "popDownYear()" & Chr(13)
            strBuildUp += "if (!monthConstructed) {" & Chr(13)
            strBuildUp += "sHTML =	""""" & Chr(13)
            strBuildUp += "for	(i=0; i<12;	i++) {" & Chr(13)
            strBuildUp += "sName =	monthName[i];" & Chr(13)
            strBuildUp += "if (i==monthSelected){" & Chr(13)
            strBuildUp += "sName =	""<B>"" +	sName +	""</B>""" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "sHTML += ""<tr><td id='m"" + i + ""' onmouseover='this.style.backgroundColor=\""#FFCC99\""' onmouseout='this.style.backgroundColor=\""\""' style='cursor:pointer' onclick='monthConstructed=false;monthSelected="" + i + "";constructCalendar();popDownMonth();event.cancelBubble=true'>&nbsp;"" + sName + ""&nbsp;</td></tr>""" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "document.getElementById(""selectMonth"").innerHTML = ""<table width=70	style='font-family:arial; font-size:11px; border-width:1; border-style:solid; border-color:#a0a0a0;' bgcolor='#FFFFDD' cellspacing=0 onmouseover='clearTimeout(timeoutID1)'	onmouseout='clearTimeout(timeoutID1);timeoutID1=setTimeout(\""popDownMonth()\"",100);event.cancelBubble=true'>"" +	sHTML +	""</table>""" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "monthConstructed=true" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "function popUpMonth() {" & Chr(13)
            strBuildUp += "constructMonth()" & Chr(13)
            strBuildUp += "crossMonthObj.visibility = (dom||ie)? ""visible""	: ""show""" & Chr(13)
            strBuildUp += "crossMonthObj.left = parseInt(crossobj.left) + 50" & Chr(13)
            strBuildUp += "crossMonthObj.top =	parseInt(crossobj.top) + 26" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "hideElement( 'SELECT', document.getElementById(""selectMonth"") );" & Chr(13)
            strBuildUp += "hideElement( 'APPLET', document.getElementById(""selectMonth"") );" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "function popDownMonth()	{" & Chr(13)
            strBuildUp += "crossMonthObj.visibility= ""hidden""" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "/*** Year Pulldown ***/" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "function incYear() {" & Chr(13)
            strBuildUp += "for	(i=0; i<7; i++){" & Chr(13)
            strBuildUp += "newYear	= (i+nStartingYear)+1" & Chr(13)
            strBuildUp += "if (newYear==yearSelected)" & Chr(13)
            strBuildUp += "{ txtYear =	""&nbsp;<B>""	+ newYear +	""</B>&nbsp;"" }" & Chr(13)
            strBuildUp += "else" & Chr(13)
            strBuildUp += "{ txtYear =	""&nbsp;"" + newYear + ""&nbsp;"" }" & Chr(13)
            strBuildUp += "document.getElementById(""y""+i).innerHTML = txtYear" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "nStartingYear ++;" & Chr(13)
            strBuildUp += "bShow=true" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "function decYear() {" & Chr(13)
            strBuildUp += "for	(i=0; i<7; i++){" & Chr(13)
            strBuildUp += "newYear	= (i+nStartingYear)-1" & Chr(13)
            strBuildUp += "if (newYear==yearSelected)" & Chr(13)
            strBuildUp += "{ txtYear =	""&nbsp;<B>""	+ newYear +	""</B>&nbsp;"" }" & Chr(13)
            strBuildUp += "else" & Chr(13)
            strBuildUp += "{ txtYear =	""&nbsp;"" + newYear + ""&nbsp;"" }" & Chr(13)
            strBuildUp += "document.getElementById(""y""+i).innerHTML = txtYear" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "nStartingYear --;" & Chr(13)
            strBuildUp += "bShow=true" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "function selectYear(nYear) {" & Chr(13)
            strBuildUp += "yearSelected=parseInt(nYear+nStartingYear);" & Chr(13)
            strBuildUp += "yearConstructed=false;" & Chr(13)
            strBuildUp += "constructCalendar();" & Chr(13)
            strBuildUp += "popDownYear();" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "function constructYear() {" & Chr(13)
            strBuildUp += "popDownMonth()" & Chr(13)
            strBuildUp += "sHTML =	""""" & Chr(13)
            strBuildUp += "if (!yearConstructed) {" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "sHTML =	""<tr><td align='center'	onmouseover='this.style.backgroundColor=\""#FFCC99\""' onmouseout='clearInterval(intervalID1);this.style.backgroundColor=\""\""' style='cursor:pointer'	onmousedown='clearInterval(intervalID1);intervalID1=setInterval(\""decYear()\"",30)' onmouseup='clearInterval(intervalID1)'>-</td></tr>""" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "j =	0" & Chr(13)
            strBuildUp += "nStartingYear =	yearSelected-3" & Chr(13)
            strBuildUp += "for	(i=(yearSelected-3); i<=(yearSelected+3); i++) {" & Chr(13)
            strBuildUp += "sName =	i;" & Chr(13)
            strBuildUp += "if (i==yearSelected){" & Chr(13)
            strBuildUp += "sName =	""<B>"" +	sName +	""</B>""" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "sHTML += ""<tr><td id='y"" + j + ""' onmouseover='this.style.backgroundColor=\""#FFCC99\""' onmouseout='this.style.backgroundColor=\""\""' style='cursor:pointer' onclick='selectYear(""+j+"");event.cancelBubble=true'>&nbsp;"" + sName + ""&nbsp;</td></tr>""" & Chr(13)
            strBuildUp += "j ++;" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "sHTML += ""<tr><td align='center' onmouseover='this.style.backgroundColor=\""#FFCC99\""' onmouseout='clearInterval(intervalID2);this.style.backgroundColor=\""\""' style='cursor:pointer' onmousedown='clearInterval(intervalID2);intervalID2=setInterval(\""incYear()\"",30)'	onmouseup='clearInterval(intervalID2)'>+</td></tr>""" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "document.getElementById(""selectYear"").innerHTML	= ""<table width=44 style='font-family:arial; font-size:11px; border-width:1; border-style:solid; border-color:#a0a0a0;'	bgcolor='#FFFFDD' onmouseover='clearTimeout(timeoutID2)' onmouseout='clearTimeout(timeoutID2);timeoutID2=setTimeout(\""popDownYear()\"",100)' cellspacing=0>""	+ sHTML	+ ""</table>""" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "yearConstructed	= true" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "function popDownYear() {" & Chr(13)
            strBuildUp += "clearInterval(intervalID1)" & Chr(13)
            strBuildUp += "clearTimeout(timeoutID1)" & Chr(13)
            strBuildUp += "clearInterval(intervalID2)" & Chr(13)
            strBuildUp += "clearTimeout(timeoutID2)" & Chr(13)
            strBuildUp += "crossYearObj.visibility= ""hidden""" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "function popUpYear() {" & Chr(13)
            strBuildUp += "var	leftOffset" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "constructYear()" & Chr(13)
            strBuildUp += "crossYearObj.visibility	= (dom||ie)? ""visible"" : ""show""" & Chr(13)
            strBuildUp += "leftOffset = parseInt(crossobj.left) + document.getElementById(""spanYear"").offsetLeft" & Chr(13)
            strBuildUp += "if (ie)" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "leftOffset += 6" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "crossYearObj.left =	leftOffset" & Chr(13)
            strBuildUp += "crossYearObj.top = parseInt(crossobj.top) +	26" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "/*** calendar ***/" & Chr(13)
            strBuildUp += "function WeekNbr(n) {" & Chr(13)
            strBuildUp += "// Algorithm used:" & Chr(13)
            strBuildUp += "// From Klaus Tondering's Calendar document (The Authority/Guru)" & Chr(13)
            strBuildUp += "// hhtp://www.tondering.dk/claus/calendar.html" & Chr(13)
            strBuildUp += "// a = (14-month) / 12" & Chr(13)
            strBuildUp += "// y = year + 4800 - a" & Chr(13)
            strBuildUp += "// m = month + 12a - 3" & Chr(13)
            strBuildUp += "// J = day + (153m + 2) / 5 + 365y + y / 4 - y / 100 + y / 400 - 32045" & Chr(13)
            strBuildUp += "// d4 = (J + 31741 - (J mod 7)) mod 146097 mod 36524 mod 1461" & Chr(13)
            strBuildUp += "// L = d4 / 1460" & Chr(13)
            strBuildUp += "// d1 = ((d4 - L) mod 365) + L" & Chr(13)
            strBuildUp += "// WeekNumber = d1 / 7 + 1" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "year = n.getFullYear();" & Chr(13)
            strBuildUp += "month = n.getMonth() + 1;" & Chr(13)
            strBuildUp += "if (startAt == 0) {" & Chr(13)
            strBuildUp += "day = n.getDate() + 1;" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "else {" & Chr(13)
            strBuildUp += "day = n.getDate();" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "a = Math.floor((14-month) / 12);" & Chr(13)
            strBuildUp += "y = year + 4800 - a;" & Chr(13)
            strBuildUp += "m = month + 12 * a - 3;" & Chr(13)
            strBuildUp += "b = Math.floor(y/4) - Math.floor(y/100) + Math.floor(y/400);" & Chr(13)
            strBuildUp += "J = day + Math.floor((153 * m + 2) / 5) + 365 * y + b - 32045;" & Chr(13)
            strBuildUp += "d4 = (((J + 31741 - (J % 7)) % 146097) % 36524) % 1461;" & Chr(13)
            strBuildUp += "L = Math.floor(d4 / 1460);" & Chr(13)
            strBuildUp += "d1 = ((d4 - L) % 365) + L;" & Chr(13)
            strBuildUp += "week = Math.floor(d1/7) + 1;" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "return week;" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "function constructCalendar () {" & Chr(13)
            strBuildUp += "var aNumDays = Array (31,0,31,30,31,30,31,31,30,31,30,31)" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "var dateMessage" & Chr(13)
            strBuildUp += "var	startDate =	new	Date (yearSelected,monthSelected,1)" & Chr(13)
            strBuildUp += "var endDate" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "if (monthSelected==1)" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "endDate	= new Date (yearSelected,monthSelected+1,1);" & Chr(13)
            strBuildUp += "endDate	= new Date (endDate	- (24*60*60*1000));" & Chr(13)
            strBuildUp += "numDaysInMonth = endDate.getDate()" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "else" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "numDaysInMonth = aNumDays[monthSelected];" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "datePointer	= 0" & Chr(13)
            strBuildUp += "dayPointer = startDate.getDay() - startAt" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "if (dayPointer<0)" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "dayPointer = 6" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "sHTML =	""<table	 border=0 style='font-family:verdana;font-size:10px;'><tr>""" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "if (showWeekNumber==1)" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "sHTML += ""<td width=27><b>"" + weekString + ""</b></td><td width=1 rowspan=7 bgcolor='#d0d0d0' style='padding:0px'><img src='""+imgDir+""" & m_dividerGif & "' width=1></td>""" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "for	(i=0; i<7; i++)	{" & Chr(13)
            strBuildUp += "sHTML += ""<td width='27' align='right'><B>""+ dayName[i]+""</B></td>""" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "sHTML +=""</tr><tr>""" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "if (showWeekNumber==1)" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "sHTML += ""<td align=right>"" + WeekNbr(startDate) + ""&nbsp;</td>""" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "for	( var i=1; i<=dayPointer;i++ )" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "sHTML += ""<td>&nbsp;</td>""" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "for	( datePointer=1; datePointer<=numDaysInMonth; datePointer++ )" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "dayPointer++;" & Chr(13)
            strBuildUp += "sHTML += ""<td align=right>""" & Chr(13)
            strBuildUp += "sStyle=styleAnchor" & Chr(13)
            strBuildUp += "if ((datePointer==odateSelected) &&	(monthSelected==omonthSelected)	&& (yearSelected==oyearSelected))" & Chr(13)
            strBuildUp += "{ sStyle+=styleLightBorder }" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "sHint = """"" & Chr(13)
            strBuildUp += "for (k=0;k<HolidaysCounter;k++)" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "if ((parseInt(Holidays[k].d)==datePointer)&&(parseInt(Holidays[k].m)==(monthSelected+1)))" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "if ((parseInt(Holidays[k].y)==0)||((parseInt(Holidays[k].y)==yearSelected)&&(parseInt(Holidays[k].y)!=0)))" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "sStyle+=""background-color:#FFDDDD;""" & Chr(13)
            strBuildUp += "sHint+=sHint==""""?Holidays[k].desc:""\n""+Holidays[k].desc" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "var regexp= /\""/g" & Chr(13)
            strBuildUp += "sHint=sHint.replace(regexp,""&quot;"")" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "dateMessage = ""onmousemove='window.status=\""""+selectDateMessage.replace(""[date]"",constructDate(datePointer,monthSelected,yearSelected))+""\""' onmouseout='window.status=\""\""' """ & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "if ((datePointer==dateNow)&&(monthSelected==monthNow)&&(yearSelected==yearNow))" & Chr(13)
            strBuildUp += "{ sHTML += ""<b><a ""+dateMessage+"" title=\"""" + sHint + ""\"" style='""+sStyle+""' href='javascript:dateSelected=""+datePointer+"";closeCalendar();'><font color=#ff0000>&nbsp;"" + datePointer + ""</font>&nbsp;</a></b>""}" & Chr(13)
            strBuildUp += "else if	(dayPointer % 7 == (startAt * -1)+1)" & Chr(13)
            strBuildUp += "{ sHTML += ""<a ""+dateMessage+"" title=\"""" + sHint + ""\"" style='""+sStyle+""' href='javascript:dateSelected=""+datePointer + "";closeCalendar();'>&nbsp;<font color=#909090>"" + datePointer + ""</font>&nbsp;</a>"" }" & Chr(13)
            strBuildUp += "else" & Chr(13)
            strBuildUp += "{ sHTML += ""<a ""+dateMessage+"" title=\"""" + sHint + ""\"" style='""+sStyle+""' href='javascript:dateSelected=""+datePointer + "";closeCalendar();'>&nbsp;"" + datePointer + ""&nbsp;</a>"" }" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "sHTML += """"" & Chr(13)
            strBuildUp += "if ((dayPointer+startAt) % 7 == startAt) {" & Chr(13)
            strBuildUp += "sHTML += ""</tr><tr>""" & Chr(13)
            strBuildUp += "if ((showWeekNumber==1)&&(datePointer<numDaysInMonth))" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "sHTML += ""<td align=right>"" + (WeekNbr(new Date(yearSelected,monthSelected,datePointer+1))) + ""&nbsp;</td>""" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "document.getElementById(""content"").innerHTML   = sHTML" & Chr(13)
            strBuildUp += "document.getElementById(""spanMonth"").innerHTML = ""&nbsp;"" +	monthName[monthSelected] + ""&nbsp;<IMG id='changeMonth' SRC='""+imgDir+""drop1.gif' WIDTH='12' HEIGHT='10' BORDER=0>""" & Chr(13)
            strBuildUp += "document.getElementById(""spanYear"").innerHTML =	""&nbsp;"" + yearSelected	+ ""&nbsp;<IMG id='changeYear' SRC='""+imgDir+""drop1.gif' WIDTH='12' HEIGHT='10' BORDER=0>""" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "function popUpCalendar(ctl,	ctl2, format) {" & Chr(13)
            strBuildUp += "var	leftpos=0" & Chr(13)
            strBuildUp += "var	toppos=0" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "if (bPageLoaded)" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "if ( crossobj.visibility ==	""hidden"" ) {" & Chr(13)
            strBuildUp += "ctlToPlaceValue	= ctl2" & Chr(13)
            strBuildUp += "dateFormat=format;" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "formatChar = "" """ & Chr(13)
            strBuildUp += "aFormat	= dateFormat.split(formatChar)" & Chr(13)
            strBuildUp += "if (aFormat.length<3)" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "formatChar = ""/""" & Chr(13)
            strBuildUp += "aFormat	= dateFormat.split(formatChar)" & Chr(13)
            strBuildUp += "if (aFormat.length<3)" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "formatChar = "".""" & Chr(13)
            strBuildUp += "aFormat	= dateFormat.split(formatChar)" & Chr(13)
            strBuildUp += "if (aFormat.length<3)" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "formatChar = ""-""" & Chr(13)
            strBuildUp += "aFormat	= dateFormat.split(formatChar)" & Chr(13)
            strBuildUp += "if (aFormat.length<3)" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "// invalid date	format" & Chr(13)
            strBuildUp += "formatChar=""""" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "tokensChanged =	0" & Chr(13)
            strBuildUp += "if ( formatChar	!= """" )" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "// use user's date" & Chr(13)
            strBuildUp += "aData =	ctl2.value.split(formatChar)" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "for	(i=0;i<3;i++)" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "if ((aFormat[i]==""d"") || (aFormat[i]==""dd""))" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "dateSelected = parseInt(aData[i], 10)" & Chr(13)
            strBuildUp += "tokensChanged ++" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "else if	((aFormat[i]==""m"") || (aFormat[i]==""mm""))" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "monthSelected =	parseInt(aData[i], 10) - 1" & Chr(13)
            strBuildUp += "tokensChanged ++" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "else if	(aFormat[i]==""yyyy"")" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "yearSelected = parseInt(aData[i], 10)" & Chr(13)
            strBuildUp += "tokensChanged ++" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "else if	(aFormat[i]==""mmm"")" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "for	(j=0; j<12;	j++)" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "if (aData[i]==monthName[j])" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "monthSelected=j" & Chr(13)
            strBuildUp += "tokensChanged ++" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "if ((tokensChanged!=3)||isNaN(dateSelected)||isNaN(monthSelected)||isNaN(yearSelected))" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "dateSelected = dateNow" & Chr(13)
            strBuildUp += "monthSelected =	monthNow" & Chr(13)
            strBuildUp += "yearSelected = yearNow" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "odateSelected=dateSelected" & Chr(13)
            strBuildUp += "omonthSelected=monthSelected" & Chr(13)
            strBuildUp += "oyearSelected=yearSelected" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "aTag = ctl" & Chr(13)
            strBuildUp += "do {" & Chr(13)
            strBuildUp += "aTag = aTag.offsetParent;" & Chr(13)
            strBuildUp += "leftpos	+= aTag.offsetLeft;" & Chr(13)
            strBuildUp += "toppos += aTag.offsetTop;" & Chr(13)
            strBuildUp += "} while(aTag.tagName!=""BODY"");" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "crossobj.left =	fixedX==-1 ? ctl.offsetLeft	+ leftpos :	fixedX" & Chr(13)
            strBuildUp += "crossobj.top = fixedY==-1 ?	ctl.offsetTop +	toppos + ctl.offsetHeight +	2 :	fixedY" & Chr(13)
            strBuildUp += "constructCalendar (1, monthSelected, yearSelected);" & Chr(13)
            strBuildUp += "crossobj.visibility=(dom||ie)? ""visible"" : ""show""" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "hideElement( 'SELECT', document.getElementById(""calendar"") );" & Chr(13)
            strBuildUp += "hideElement( 'APPLET', document.getElementById(""calendar"") );" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "bShow = true;" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "else" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "hideCalendar()" & Chr(13)
            strBuildUp += "if (ctlNow!=ctl) {popUpCalendar(ctl, ctl2, format)}" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "ctlNow = ctl" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "document.onkeypress = function hidecal1 () {" & Chr(13)
            strBuildUp += "if (event.keyCode==27)" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "hideCalendar()" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "document.onclick = function hidecal2 () {" & Chr(13)
            strBuildUp += "if (!bShow)" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "hideCalendar()" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "bShow = false" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "" & Chr(13)
            strBuildUp += "if(ie)" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "init()" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "else" & Chr(13)
            strBuildUp += "{" & Chr(13)
            strBuildUp += "window.onload=init" & Chr(13)
            strBuildUp += "}" & Chr(13)
            strBuildUp += "<"
            strBuildUp += "/"
            strBuildUp += "script>"
            Page.Response.Write(strBuildUp)
        End Sub


        'Register Controls
        Protected Overrides Sub CreateChildControls()
            placeJavascript()
            Dim txtTextBox As New System.Web.UI.WebControls.TextBox()
            If Len(m_ControlCssClass) > 0 Then
                txtTextBox.CssClass = m_ControlCssClass
            End If
            txtTextBox.ReadOnly = True
            If Not (m_text = "") Then
                txtTextBox.Text = m_text
            End If
            If Not (m_Css = "") Then
                txtTextBox.CssClass = m_Css
            End If
            txtTextBox.ID = "foo"
            txtTextBox.Attributes.Add("onclick", "popUpCalendar(document.all." & Me.ClientID & "_foo,document.all." & _
                Me.ClientID & "_foo, '" & m_DateType & "')")
            Me.Controls.Add(txtTextBox)
        End Sub

    End Class

End Namespace



