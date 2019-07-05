<% Option Explicit %>
<!--#include file="../../Inc/Cls_DB.asp" -->
<!--#include file="../../Inc/Const.asp" -->
<!--#include file="../../Inc/NoSqlHack.asp" -->
<%
Dim DBC,Conn,HelpConn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
DBC.ConnStr = "DBQ=" + Server.MapPath("Foosun_help.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
Set HelpConn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../Inc/Session.asp" -->
<%
Dim HelpKeyWord,HelpPage,HelpID,IDHelpStr,DefauleHelpStr,DataTypeStr,HelpKeyWordShow
HelpKeyWord = Replace(Replace(Replace(Request("KeyWord"),Chr(13),""),Chr(10),"")," ","")
HelpPage = Request("Page")
HelpID = Request("HelpID")
DefauleHelpStr = ""
If IsSqlDataBase=0 then
	DataTypeStr = "#"
Else
	DataTypeStr = "'"
End If
if HelpKeyWord = "" then
	DefauleHelpStr = GetDefauleHelpStr
else
	Dim RsHelpObj,HelpSql,RecCount,HelpStr,HelpIndex
	'Response.Write(Request.ServerVariables("HTTP_REFERER"))
	'Response.End
	HelpKeyWord = Replace(Replace(Replace(Replace(HelpKeyWord,"[",""),"]",""),"'",""),"""","")
	HelpPage = Replace(Replace(Replace(Replace(HelpPage,"[",""),"]",""),"'",""),"""","")
	if HelpPage <> "" then
		HelpSql = "Select HelpSingleContent,ID,FuncName,PageField from FS_Help where FileName='" & HelpPage & "' and PageField like '%" & HelpKeyWord & "%'"
	else
		HelpSql = "Select HelpSingleContent,ID,FuncName,PageField from FS_Help where PageField like '%" & HelpKeyWord & "%'"
	end if
	Set RsHelpObj = Server.CreateObject(G_FS_RS)
	RsHelpObj.Open HelpSql,HelpConn,1,1
	RecCount = RsHelpObj.RecordCount
	if RecCount = 0 then
		DefauleHelpStr = GetNoneHelpStr
	elseif RecCount = 1 then
		HelpID = RsHelpObj("ID")
		IDHelpStr = RsHelpObj("HelpSingleContent")
		HelpKeyWordShow = RsHelpObj("PageField")
	else
		HelpIndex = 1
		do while Not RsHelpObj.Eof
			if HelpStr = "" then
				HelpStr =  "<strong>" & HelpIndex & " : </strong>" & "<span style=""CURSOR:hand;"" onClick=""BrowHelp('" & RsHelpObj("ID") & "','" & HelpKeyWord & "','" & HelpPage & "');"">" & GetOneKeyWord(RsHelpObj("PageField"),HelpKeyWord)  & "</span>"
			else
				HelpStr = HelpStr & "<br>" & "<strong>" & HelpIndex & " : </strong>" & "<span style=""CURSOR:hand;"" onClick=""BrowHelp('" & RsHelpObj("ID") & "','" & HelpKeyWord & "','" & HelpPage & "');"">" & GetOneKeyWord(RsHelpObj("PageField"),HelpKeyWord) & "</span>"
			end if
			if IDHelpStr = "" then
				if HelpID = "" then
					IDHelpStr = ""
				else
					if RsHelpObj("ID") = CInt(HelpID) then
						IDHelpStr = RsHelpObj("HelpSingleContent")
					end if
					HelpKeyWordShow = RsHelpObj("PageField")
				end if
			end if
			HelpIndex = HelpIndex + 1
			RsHelpObj.MoveNext
		Loop
	end if
	Set RsHelpObj = Nothing
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style>
td,table,BODY {
background:center;
font-family:  "Verdana, Arial, Helvetica, sans-serif"; font-size: 11.5px;
FONT-SIZE: 9pt;
line-height:16px;
color: #393939;
text-decoration: none;
scrollbar-face-color: #f6f6f6;
scrollbar-highlight-color: #ffffff; scrollbar-shadow-color: #cccccc; scrollbar-3dlight-color: #cccccc; scrollbar-arrow-color: #000000; scrollbar-track-color: #EFEFEF; scrollbar-darkshadow-color: #ffffff;
}
</style>
<link href="../../CSS/FS_css.css" rel="stylesheet">
<title>帮助</title>
<script language="Javascript">
<!--
function zoomimg(img)
{
return false;
}
-->
</script>
</head>
<body topmargin="3" leftmargin="3" oncontextmenu="//return false;">
<table width="100%" height="100%" border="0" cellpadding="1" cellspacing="0">
  <%
if DefauleHelpStr <> "" then
%>
  <tr> 
    <td height="10" valign="top"><font color="393939"> 
      <% = DefauleHelpStr %>
      </font></td>
  </tr>
  <%
else
	if HelpStr <> "" then
%>
  <tr> 
    <td height="10" valign="top"><font color="393939"> 
      <% = HelpStr %>
      </font></td>
  </tr>
  <%
	end if
	Dim regEx
	if (Not IsNull(IDHelpStr)) OR IDHelpStr <> "" then
		Set regEx = New RegExp
		regEx.Pattern = "<\/*(ol|p|font|span|table|tr|td|tbody|li|ul|a|img)[^<>]*>"
		regEx.IgnoreCase = True
		regEx.Global = True
		IDHelpStr = regEx.Replace(IDHelpStr,"")
		Set regEx = Nothing
	end if
	if IDHelpStr <> "" then
%>
  <tr> 
    <td valign="top"><font color="393939"><strong> <font color="#FF3300"><% Response.write(GetOneKeyWord(HelpKeyWordShow,HelpKeyWord)) %>
      </font></strong><font color="#393939">的</font>帮助<strong>：</strong><br>
      <% = IDHelpStr %>
      </font></td>
  </tr>
  <tr> 
    <td height="10" align="right" valign="top"><span style="CURSOR:hand;" onClick="OpenMoreWindow('<% = HelpID %>','<% = HelpKeyWord %>');"><font color="#393939">详细帮助...</font></span></td>
  </tr>
  <%
	end if
end if
%>
</table>
</body>
</html>
<%
Set Conn = Nothing
Function GetDefauleHelpStr()
	GetDefauleHelpStr = "<font color=#FF3300><strong>帮助说明</strong></font><br>一、按<font color=#FF3300>""Ctrl+数字键1""</font>组合键，帮助窗口里面会显示鼠标所在地方或者窗口里面选中对象的帮助。<br>二、单击<span style=""CURSOR:hand;"" onclick=""OpenHelpWindow('MoreHelpInfo.htm');""><strong><font color=""#FF0000""><b>这里</b></font></strong></span>获得更多帮助使用说明。"
End Function
Function GetNoneHelpStr()
	GetNoneHelpStr = "<strong>说明</strong>：<br>系统没有找到<font color=#FF3300>"& HelpKeyWord &"</font>此关键字的帮助。您可以到<span style=""CURSOR:hand;"" onclick=""Search('" & HelpKeyWord & "');""><strong><font color=""#FF0000""><b>这里</b></font></strong></span>去查找帮助。或者到官方论坛发帖，请求帮助。论坛地址:http://bbs.foosun.net"
End Function
Function GetOneKeyWord(KeyWordStr,InstrStr)
	Dim TempArray,TempI
	GetOneKeyWord = ""
	if Not IsNull(KeyWordStr) then
		TempArray = Split(Replace(KeyWordStr,"，",","),",")
		for TempI = LBound(TempArray) to Ubound(TempArray)
			if Instr(TempArray(TempI),InstrStr) then GetOneKeyWord=TempArray(TempI)
		Next
	else
		GetOneKeyWord = ""
	end if
	if GetOneKeyWord = "" then GetOneKeyWord = InstrStr
End Function
%>
<script language="JavaScript">
function BrowHelp(ID,HelpKeyWord,HelpPage)
{
	location='?HelpID='+escape(ID)+'&KeyWord='+escape(HelpKeyWord)+'&Page='+escape(HelpPage)
}
function Search(HelpKeyWord)
{
	var HrefStr='http://help.foosun.net/Search.asp?Keyword='+escape(HelpKeyWord)+'&condition=content';
	window.open(HrefStr);
}
function OpenMoreWindow(HelpID,HelpKeyWord)
{
	window.open('ReadMore.asp?ID='+HelpID+'&HelpKeyWord='+escape(HelpKeyWord),'HelpWindow','width=720,height=380,top='+(screen.height-380)/2+',left='+(screen.width-720)/2+',resizable=yes,status=1,scrollbars=1');
}
function OpenHelpWindow(URL)
{
	window.open(URL,'HelpWindow','width=720,height=380,top='+(screen.height-380)/2+',left='+(screen.width-720)/2+',resizable=yes,status=1,scrollbars=1');
}
</script>
<%Set HelpConn = Nothing%>