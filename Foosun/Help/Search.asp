<% Option Explicit %>
<!--#include file="../../Inc/Cls_DB.asp" -->
<!--#include file="../../Inc/Const.asp" -->
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
Dim HelpKeyWord,HelpID
HelpKeyWord = Replace(Replace(Replace(Request("KeyWord"),Chr(13),""),Chr(10),"")," ","")
HelpID = Replace(Replace(Replace(Request("HelpID"),Chr(13),""),Chr(10),"")," ","")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style>
/*td,table,BODY {
background:center;
font-family:  "Verdana, Arial, Helvetica, sans-serif"; font-size: 11.5px;
FONT-SIZE: 9pt;
line-height:16px;
text-decoration: none;
scrollbar-face-color: #f6f6f6;
scrollbar-highlight-color: #ffffff; scrollbar-shadow-color: #cccccc; scrollbar-3dlight-color: #cccccc; scrollbar-arrow-color: #000000; scrollbar-track-color: #EFEFEF; scrollbar-darkshadow-color: #ffffff;
}
*/</style>
<link href="../../CSS/FS_css.css" rel="stylesheet">
<title>帮助</title>
</head>
<body topmargin="5" leftmargin="5" oncontextmenu="//return false;">
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#DEDEDE">
  <%
if HelpKeyWord = "" then
%>
  <tr> 
    <td height="10" valign="top" bgcolor="#EFEFEF"><font color="393939">没有搜索关键字</font></td>
  </tr>
<%
else
%>
  <tr> 
    <td height="10" valign="top" bgcolor="#EFEFEF"><strong>全部符合关键字""<font color="#FF3300"> 
      <% = HelpKeyWord %>
      </font> ""的帮助</strong>：</td>
  </tr>
<%
	Dim RsHelpObj,HelpSql,RecCount,HelpIndex
	HelpKeyWord = Replace(Replace(Replace(Replace(HelpKeyWord,"[",""),"]",""),"'",""),"""","")
	HelpSql = "Select * from FS_Help where PageField like '%" & HelpKeyWord & "%'"
	Set RsHelpObj = Server.CreateObject(G_FS_RS)
	RsHelpObj.Open HelpSql,HelpConn,1,1
	RecCount = RsHelpObj.RecordCount
	if RecCount = 0 then
%>
  <tr> 
    <td height="10" valign="top" bgcolor="#EFEFEF"><font color="393939">没有符合关键字的纪录</font></td>
  </tr>
<%
	elseif RecCount = 1 then
%>
  <tr> 
    <td bgcolor="#EFEFEF">
<% = GetHelpContent(RsHelpObj) %></td>
  </tr>
<%
	else
		HelpIndex = 1
		if HelpID = "" then HelpID = RsHelpObj("ID")
		do while Not RsHelpObj.Eof
			%>
			  <tr> 
    <td height="10" valign="top" bgcolor="#EFEFEF">
<% = HelpIndex %>、<span style="CURSOR:hand;" onClick="BrowHelp('<% = RsHelpObj("ID") %>','<% = HelpKeyWord %>');"><% = GetOneKeyWord(RsHelpObj("PageField"),HelpKeyWord) %></span></td>
			  </tr>
			<%
			if HelpID <> "" then
				if RsHelpObj("ID") = CInt(HelpID) then
%>
  <tr> 
	<td bgcolor="#EFEFEF">
<% = GetHelpContent(RsHelpObj) %></td>
  </tr>
<%
				end if
			end if
			HelpIndex = HelpIndex + 1
			RsHelpObj.MoveNext
		Loop
	end if
	Set RsHelpObj = Nothing
end if
%>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td height="40"> <div align="center"><a href="javascript:window.close()"><img src="../Images/Colse.gif" alt="关闭窗口" border="0"></a>　<a href="http://help.foosun.net/Search.asp?Keyword=<% = Server.HTMLEncode(Request("KeyWord")) %>&condition=content"; target="_blank"><img src="../Images/ReHelp.gif" width="119" height="28" border="0"></a></div></td>
  </tr>
</table>
</body>
</html>
<%
Set Conn = Nothing
Set HelpConn = Nothing
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

Function GetHelpContent(RsObj)
%>
<table cellpadding=4 width="98%" cellspacing=1 align=center bgcolor="#DEDEDE" style="padding:0px 4px;">
  <tr bgcolor="#EFEFEF"> 
    <td width="83" nowrap> <div align="right"><strong>页面功能</strong></div></td>
    <td width="889" bgcolor="#F7F7F7"><% = RsObj("FuncName") %></td>
  </tr>
  <tr bgcolor="#EFEFEF"> 
    <td nowrap> <div align="right"><strong>页面地址</strong></div></td>
    <td bgcolor="#F7F7F7"><% = RsObj("FileName") %></td>
  </tr>
  <tr bgcolor="#EFEFEF"> 
    <td nowrap> <div align="right"><strong>关键字</strong></div></td>
    <td bgcolor="#F7F7F7"><% = RsObj("PageField") %></td>
  </tr>
  <tr bgcolor="#EFEFEF"> 
    <td nowrap> <div align="right"><strong>简单说明</strong></div></td>
    <td height="58" valign="top" bgcolor="#F7F7F7"><% = RsObj("HelpSingleContent") %></td>
  </tr>
  <tr bgcolor="#EFEFEF"> 
    <td nowrap> <div align="right"><strong>详细说明</strong></div></td>
    <td bgcolor="#F7F7F7"><% = RsObj("HelpContent") %></td>
  </tr>
</table>
<%
End Function
%>
<script language="JavaScript">
function BrowHelp(ID,HelpKeyWord)
{
	location='?HelpID='+escape(ID)+'&KeyWord='+escape(HelpKeyWord);
}
</script>