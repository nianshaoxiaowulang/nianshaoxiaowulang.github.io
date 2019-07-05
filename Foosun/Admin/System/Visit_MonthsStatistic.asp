<% Option Explicit %>
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<%
'==============================================================================
'软件名称：风讯网站信息管理系统
'当前版本：Foosun Content Manager System(FoosunCMS V3.1.0930)
'最新更新：2005.10
'==============================================================================
'Copyright (C) 2002-2004 Foosun.Net  All rights reserved.
'商业注册联系：028-85098980-601,项目开发：028-85098980-606、609,客户支持：608
'产品咨询QQ：394226379,159410,125114015
'技术支持QQ：315485710,66252421 
'项目开发QQ：415637671，655071
'程序开发：四川风讯科技发展有限公司(Foosun Inc.)
'Email:service@Foosun.cn
'MSN：skoolls@hotmail.com
'论坛支持：风讯在线论坛(http://bbs.foosun.net)
'官方网站：www.Foosun.cn  演示站点：test.cooin.com 
'网站通系列(智能快速建站系列)：www.ewebs.cn
'==============================================================================
'免费版本请在程序首页保留版权信息，并做上本站LOGO友情连接
'风讯公司保留此程序的法律追究权利
'如需进行2次开发，必须经过风讯公司书面允许。否则将追究法律责任
'==============================================================================
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P080506") then Call ReturnError1()
%>
<html>
<head>
<title>按月信息统计</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<body bgcolor="#FFFFFF" topmargin="2" leftmargin="2">
<table height="26" width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
      <td height="28" class="ButtonListLeft"><div align="center"><strong>按月信息统计</strong></div></td>
  </tr>
</table>
<%
Dim RsMonthObj,Sql
Set RsMonthObj = Server.CreateObject(G_FS_RS)
Dim MaxVisitCount,VisitTime,VisitMonth,CurrentMonth
CurrentMonth=Month(Now())
If IsSqlDataBase=0 then
	Sql="Select VisitTime From FS_FlowStatistic where DATEDIFF('m',VisitTime,Now()) < 12 And DATEDIFF('m',VisitTime,Now()) >=0"
Else
	Sql="Select VisitTime From FS_FlowStatistic where DATEDIFF(month,VisitTime,GetDate()) < 12 And DATEDIFF(month,VisitTime,GetDate()) >=0"
End If
RsMonthObj.Open Sql,Conn,3,3
MaxVisitCount=0
Dim VisitNum(12),I
for I=1 To 12
	VisitNum(I)=0
next
Do While not RsMonthObj.Eof 
	VisitTime = RsMonthObj("VisitTime")
	VisitMonth = Month(VisitTime)
	for I=1 To 12
		if I=VisitMonth then
			VisitNum(I)=VisitNum(I)+1
		end if
	next
	RsMonthObj.MoveNext
Loop
for I=1 To 12
	if VisitNum(I)>=MaxVisitCount  then
		MaxVisitCount=VisitNum(I)
	end if
next
%>
<% 
	Dim VisitSize(12)
	For I=1 To 12
	if MaxVisitCount<>0 then
	VisitSize(I)=100*VisitNum(I)/MaxVisitCount
	else
	VisitSize(I)=0
	end if
	Next
%>
<table border=0 align="center" cellpadding=2>
	<tr>
		<td align=center>最近12个月统计图表</td>
	</tr>
	<tr>
		
    <td align=center><table border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><table border=0 align=center cellpadding=0 cellspacing=0>
              <tr> 
                <td height="25" valign=top align="right" nowrap> 
                  <% Response.Write(MaxVisitCount&"次")%>
               </td>
              </tr>
              <tr> 
                <td  height="25" valign=top align="right"  nowrap>
				<% if MaxVisitCount>3 then  
					Response.Write(Round(MaxVisitCount*0.75)&"次") 
					elseif MaxVisitCount>1  then 
					Response.Write((MaxVisitCount-1)&"次") 
					else Response.Write("&nbsp;") 
					end if
				%> 
				</td>
              </tr>
              <tr> 
                <td  height="25" valign=top  align="right" nowrap>
				<% if MaxVisitCount>3 then  
					Response.Write(Round(MaxVisitCount*0.5)&"次") 
					elseif MaxVisitCount>2  then 
					Response.Write((MaxVisitCount-2)&"次") 
					else Response.Write("&nbsp;") 
					end if
				%>
				</td>
              </tr>
              <tr>
                <td  height="25" valign=top  align="right" nowrap>
                <% if MaxVisitCount>3 then 
					 Response.Write(Round(MaxVisitCount*0.25)&"次") 
					else Response.Write("&nbsp;") 
					end if
				%>
				</td>
              </tr>
              <tr> 
                <td  height="31" valign=top  align="right" nowrap>0次</td>
              </tr>
            </table></td>
          <td valign="bottom">
<table align=center>
              <tr valign=bottom > 
			  <% if CurrentMonth =12 then %>
                <% For I=1 To 12 %>
                <td width="15" align=center nowrap background=../../Images/Visit/tu_back.gif><img src=../../Images/Visit/bar.gif width="15" height="<% =VisitSize(I) %>" id=htav><br>
                  <% = I%>
                </td>
                <% Next %>
			<% else %>
                <% For I=CurrentMonth+1 To 12 %>
                <td width="15" align=center nowrap background=../../Images/Visit/tu_back.gif><img src=../../Images/Visit/bar.gif width="15" height="<% =VisitSize(I) %>" id=htav><br>
                  <% = I%>
                </td>
                <% Next %>
                <% For I=1 To CurrentMonth  %>
                <td width="15" align=center nowrap background=../../Images/Visit/tu_back.gif><img src=../../Images/Visit/bar.gif width="15" height="<% =VisitSize(I) %>" id=htav><br>
                  <% = I%>
                </td>
                <% Next %>
			<% end if %>
                <td>单位（月）</td>
              </tr>
            </table></td>
        </tr>
      </table> </td>
	</tr>
</table>
<p>&nbsp;</p>
<%
Dim RsMonthsObj
Set RsMonthsObj = Server.CreateObject(G_FS_RS)
Sql="Select VisitTime From FS_FlowStatistic"
RsMonthsObj.Open Sql,Conn,3,3
MaxVisitCount=0
for I=1 To 12
	VisitNum(I)=0
next
Do While not RsMonthsObj.Eof 
	VisitTime = RsMonthsObj("VisitTime")
	VisitMonth = Month(VisitTime)
	for I=1 To 12
		if I=VisitMonth then
			VisitNum(I)=VisitNum(I)+1
		end if
	next
	RsMonthsObj.MoveNext
Loop
for I=1 To 12
	if VisitNum(I)>=MaxVisitCount  then
		MaxVisitCount=VisitNum(I)
	end if
next
%>
<% 
	Dim VisitsSize(12),AllCount
	AllCount=0
	For I=1 To 12
		Allcount=AllCount+VisitNum(I)
	Next
	For I=0 To 12
		if VisitNum(I)<>0 then
			VisitsSize(I)=100*VisitNum(I)/AllCount
		else
			VisitsSize(I)=0
		end if
	Next
%>
<table border=0 align="center" cellpadding=2>
	<tr>
		<td align=center>访问量12个月分配图表</td>
	</tr>
	<tr>
		
    <td align=center><table border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><table border=0 align=center cellpadding=0 cellspacing=0>
              <tr> 
                <td height="25" valign=top align="right" nowrap>100%</td>
              </tr>
              <tr> 
                <td  height="25" valign=top align="right"  nowrap>75%</td>
              </tr>
              <tr> 
                <td  height="25" valign=top  align="right" nowrap>50%</td>
              </tr>
              <tr>
                <td  height="25" valign=top  align="right" nowrap>25%</td>
              </tr>
              <tr> 
                <td  height="31" valign=top  align="right" nowrap>0</td>
              </tr>
            </table></td>
          <td valign="bottom">
<table align=center>
              <tr valign=bottom >
                <td width="15" align=center nowrap background=../images/tu_back.gif><img src=../../Images/Visit/bar.gif width="15" height="100" id=htav><br>
                  总</td>
                <% For I=1 To 12 %>
                <td width="15" align=center nowrap background=../../Images/Visit/tu_back.gif><img src=../../Images/Visit/bar.gif width="15" height="<% =VisitsSize(I) %>" id=htav><br>
                  <% = I%>
                </td>
                <% Next %>
                <td>单位（月）</td>
              </tr>
            </table></td>
        </tr>
      </table> </td>
	</tr>
</table>
</body>
</html>
<%
Conn.Close
Set Conn=Nothing
%>
