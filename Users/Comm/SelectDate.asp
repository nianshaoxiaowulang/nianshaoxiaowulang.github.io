<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>选择日期</title>
</head>

<link href="../Css/UserCss.css" rel="stylesheet">
<style type="text/css">
<!--
Body{
	cursor: default;

}
.IntialStyle{
	border-top-width: 1px;
	border-right-width: 1px;
	border-bottom-width: 1px;
	border-left-width: 1px;
	border-top-style: solid;
	border-right-style: solid;
	border-bottom-style: solid;
	border-left-style: solid;
	border-top-color: #000000;
	border-right-color: #FFFFFF;
	border-bottom-color: #FFFFFF;
	border-left-color: #000000;

}
.SelectStyle{

}
.DateMouseOver {
	border-top-width: 1px;
	border-right-width: 1px;
	border-bottom-width: 1px;
	border-left-width: 1px;
	border-top-style: solid;
	border-right-style: solid;
	border-bottom-style: solid;
	border-left-style: solid;
	border-top-color: #FFFFFF;
	border-right-color: #000000;
	border-bottom-color: #000000;
	border-left-color: #FFFFFF;
}
.DateStyle {
	cursor: default;
	border: 1px solid buttonface;
}
-->
</style>
<body>
<div align="center">
  <table border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td height="30" colspan="4" nowrap>
	  <select style="width:98%;" onChange="ChangeDateNum();" name="YearList">
         <%
		for i = 1920 to left(date,4)
		if i = left(date,4) then
			response.write"<option value="""&i&"""selected"
			response.write">"&i&"</option>"
		else
			response.write"<option value="""&i&""""
			response.write">"&i&"</option>"
		end if
		next
		%>
        </select> </td>
      <td height="30" colspan="4" nowrap><select style="width:98%;" onChange="ChangeDateNum();" name="MonthList">
          <option value="01">一月</option>
          <option value="02">二月</option>
          <option value="03">三月</option>
          <option value="04">四月</option>
          <option value="05">五月</option>
          <option value="06">六月</option>
          <option value="07">七月</option>
          <option value="08">八月</option>
          <option value="09">九月</option>
          <option value="10">十月</option>
          <option value="11">十一月</option>
          <option value="12">十二月</option>
        </select> </td>
      <td height="30" colspan="3" align="center">
<table onMouseOver="this.className='DateMouseOver';"   style="display:none" onMouseOut="this.className='';" onMouseDown="this.className='IntialStyle';" onClick="TimeClick();" width="80%" border="1" cellspacing="0" cellpadding="0">
          <tr>
            <td align="center" bordercolor="buttonface"><font size="3"><span id="TimeInput"></span></font></td>
          </tr>
        </table>
        双击日期选择</td>
    </tr>
    <tr> 
      <td width="30" height="30" align="center" class="DateStyle">1</td>
      <td width="30" height="30" align="center" class="DateStyle">2</td>
      <td width="30" height="30" align="center" class="DateStyle">3</td>
      <td width="30" height="30" align="center" class="DateStyle">4</td>
      <td width="30" height="30" align="center" class="DateStyle">5</td>
      <td width="30" height="30" align="center" class="DateStyle">6</td>
      <td width="30" height="30" align="center" class="DateStyle">7</td>
      <td width="30" height="30" align="center" class="DateStyle">8</td>
      <td width="30" height="30" align="center" class="DateStyle">9</td>
      <td width="30" height="30" align="center" class="DateStyle">10</td>
      <td width="30" height="30" align="center" class="DateStyle">11</td>
    </tr>
    <tr> 
      <td height="30" align="center" class="DateStyle">12</td>
      <td height="30" align="center" class="DateStyle">13</td>
      <td height="30" align="center" class="DateStyle">14</td>
      <td height="30" align="center" class="DateStyle">15</td>
      <td height="30" align="center" class="DateStyle">16</td>
      <td height="30" align="center" class="DateStyle">17</td>
      <td height="30" align="center" class="DateStyle">18</td>
      <td height="30" align="center" class="DateStyle">19</td>
      <td height="30" align="center" class="DateStyle">20</td>
      <td height="30" align="center" class="DateStyle">21</td>
      <td height="30" align="center" class="DateStyle">22</td>
    </tr>
    <tr> 
      <td height="30" align="center" class="DateStyle">23</td>
      <td height="30" align="center" class="DateStyle">24</td>
      <td height="30" align="center" class="DateStyle">25</td>
      <td height="30" align="center" class="DateStyle">26</td>
      <td height="30" align="center" class="DateStyle">27</td>
      <td height="30" align="center" class="DateStyle">28</td>
      <td height="30" id="Date29" align="center" class="DateStyle">29</td>
      <td height="30" id="Date30" align="center" class="DateStyle">30</td>
      <td height="30" id="Date31" align="center" class="DateStyle">31</td>
      <td height="30" align="center"></td>
      <td height="30" align="center"></td>
    </tr>
  </table>
  
</div>
</body>
</html>
<script language="JavaScript">
var bInitialized = false;
var AlreadySelectDate='';
window.setInterval('SetTimeInput();',1000);
function document.onreadystatechange()
{
	if (document.readyState!="complete") return;
	if (bInitialized) return;
	bInitialized = true;
	var i,Curr;
	for (i=0; i<document.body.all.length;i++)
	{
		Curr=document.body.all[i];
		if (Curr.className == "DateStyle") InitBtn(Curr);
	}
	var NowDate,YearStr,MonthStr,DateStr;
	NowDate=new Date();
	YearStr=NowDate.getYear();
	MonthStr=NowDate.getMonth()+1;
	DateStr=NowDate.getDate();
	SelectOption(document.all.YearList,YearStr);
	SelectOption(document.all.MonthList,MonthStr);
	SelectDate(DateStr);
	AlreadySelectDate=DateStr;
	SetTimeInput();
	ChangeDateNum();
}
function SetTimeInput()
{
	var NowDate=new Date();
	var MinuteStr= new String(NowDate.getMinutes());
	if (MinuteStr.length==1) MinuteStr='0'+MinuteStr;
	var SecondStr=new String(NowDate.getSeconds());
	if (SecondStr.length==1) SecondStr='0'+SecondStr;
	var TimeStr=NowDate.getHours()+':'+MinuteStr+':'+SecondStr;
	document.all.TimeInput.innerText=TimeStr;
}
function InitBtn(btn) 
{
	btn.onmouseover = BtnMouseOver;
	btn.onmouseout = BtnMouseOut;
	btn.onmousedown = BtnMouseDown;
	btn.onmouseup = BtnMouseOut;
	btn.onclick=DateClick;
	btn.ondblclick=DateDblClick;
	btn.disabled=false;
	return true;
}
function BtnMouseOver() 
{
	var image = event.srcElement;
	image.className = "DateMouseOver";
	event.cancelBubble = true;
}
function BtnMouseOut() 
{
	var image = event.srcElement;
	image.className = "DateStyle";
	event.cancelBubble = true;
}
function BtnMouseDown() 
{
	var image = event.srcElement;
	image.className = "IntialStyle";
	event.cancelBubble = true;
	event.returnValue=false;
	return false;
}
function SelectOption(SelectObj,Val)
{
	for (var i=0;i<SelectObj.options.length;i++)
	{
		if (SelectObj.options(i).value==Val) SelectObj.options(i).selected=true;
	}
}
function SelectDate(Val)
{
	for(var i=0;i<document.all.length;i++)
	{
		if (document.all(i).innerText==Val)
		{
			//document.all(i).className='IntialStyle';
			document.all(i).bgColor='highlight';
			document.all(i).style.color='white';
		}
	}
}
function DateClick()
{
	AlreadySelectDate=event.srcElement.innerText;
	for (var i=0;i<document.all.length;i++)
	{
		document.all(i).bgColor='';
		document.all(i).style.color='Black';
	}
	event.srcElement.bgColor='highlight';
	event.srcElement.style.color='white';
}
function DateDblClick()
{
	var TempDateStr='';
	TempDateStr=event.srcElement.innerText;
	if (TempDateStr.length==1) TempDateStr='0'+TempDateStr;
	window.returnValue=document.all.YearList.value+'-'+document.all.MonthList.value+'-'+TempDateStr;
	window.close();
}
function TimeClick()
{
	var TempDateStr='';
	TempDateStr=AlreadySelectDate;
	if (TempDateStr.length==1) TempDateStr='0'+TempDateStr;
	window.returnValue=document.all.YearList.value+'-'+document.all.MonthList.value+'-'+AlreadySelectDate+' '+document.all.TimeInput.innerText;
	window.close();
}
window.onunload=CheckReturnValue;
function CheckReturnValue()
{
	if (typeof(window.returnValue)!='string') window.returnValue='';
}
function ChangeDateNum()
{
	var YearStr=document.all.YearList.value;
	var MonthStr=document.all.MonthList.value;
	var DateNumber=GetDayNum(YearStr,MonthStr);
	switch (DateNumber)
	{
		case 31:
			document.all.Date29.style.display='';
			document.all.Date30.style.display='';
			document.all.Date31.style.display='';
			break;
		case 30:
			document.all.Date29.style.display='';
			document.all.Date30.style.display='';
			document.all.Date31.style.display='none';
			break;
		case 29:
			document.all.Date29.style.display='';
			document.all.Date30.style.display='none';
			document.all.Date31.style.display='none';
			break;
		case 28:
			document.all.Date29.style.display='none';
			document.all.Date30.style.display='none';
			document.all.Date31.style.display='none';
			break;
		default :
			document.all.Date29.style.display='none';
			document.all.Date30.style.display='none';
			document.all.Date31.style.display='none';
	}
}
function GetDayNum(YearVar, MonthVar)
{
    var Temp,LeapYear,i,BigMonth;
    var BigMonthArray=new Array(1,3,5,7,8,10,12);
    YearVar=parseInt(YearVar);
    MonthVar=parseInt(MonthVar);
    Temp=parseInt(YearVar/4);
    if (YearVar==Temp*4) LeapYear=true;
    else LeapYear = false
    for(i=0;i<BigMonthArray.length;i++)
	{
        if (MonthVar==BigMonthArray[i])
		{
			BigMonth=true;
            break;
		}
        else BigMonth=false;
    }
    if (BigMonth==true) return 31;
    else
	{
        if (MonthVar==2)
		{
            if (LeapYear==true) return 29;
            else return 28;
		}
        else  return 30;
    }
}
</script>