<% Option Explicit %>
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<%
'==============================================================================
'������ƣ���Ѷ��վ��Ϣ����ϵͳ
'��ǰ�汾��Foosun Content Manager System(FoosunCMS V3.1.0930)
'���¸��£�2005.10
'==============================================================================
'Copyright (C) 2002-2004 Foosun.Net  All rights reserved.
'��ҵע����ϵ��028-85098980-601,��Ŀ������028-85098980-606��609,�ͻ�֧�֣�608
'��Ʒ��ѯQQ��394226379,159410,125114015
'����֧��QQ��315485710,66252421 
'��Ŀ����QQ��415637671��655071
'���򿪷����Ĵ���Ѷ�Ƽ���չ���޹�˾(Foosun Inc.)
'Email:service@Foosun.cn
'MSN��skoolls@hotmail.com
'��̳֧�֣���Ѷ������̳(http://bbs.foosun.net)
'�ٷ���վ��www.Foosun.cn  ��ʾվ�㣺test.cooin.com 
'��վͨϵ��(���ܿ��ٽ�վϵ��)��www.ewebs.cn
'==============================================================================
'��Ѱ汾���ڳ�����ҳ������Ȩ��Ϣ�������ϱ�վLOGO��������
'��Ѷ��˾�����˳���ķ���׷��Ȩ��
'�������2�ο��������뾭����Ѷ��˾������������׷����������
'==============================================================================
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P080505") then Call ReturnError1()
%>
<html>
<head>
<title>������Ϣͳ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<body bgcolor="#FFFFFF" topmargin="2" leftmargin="2">
<table height="26" width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
      <td height="28" class="ButtonListLeft"><div align="center"><strong>������Ϣͳ��</strong></div></td>
  </tr>
</table>
<%
Dim RsDayObj,Sql
Set RsDayObj = Server.CreateObject(G_FS_RS)
Dim MaxVisitCount,VisitTime,VisitDay,CurrentDay
CurrentDay=Day(Now())
Dim DaysOfMonth
if Month(Now())=1 then
	DaysOfMonth=GetDayNum(Year(Now())-1, 12)
else
	DaysOfMonth=GetDayNum(Year(Now()), Month(Now())-1)
end if
If IsSqlDataBase=0 then
	Sql="Select VisitTime From FS_FlowStatistic where DATEDIFF('d',VisitTime,Now()) <= 'DaysOfMonth' And DATEDIFF('d',VisitTime,Now()) >=0"
Else
	Sql="Select VisitTime From FS_FlowStatistic where DATEDIFF(day,VisitTime,GetDate()) <= " & DaysOfMonth & " And DATEDIFF(day,VisitTime,GetDate()) >=0"
End If
RsDayObj.Open Sql,Conn,3,3
MaxVisitCount=0
Dim VisitNum(31),I
for I=1 To DaysOfMonth
	VisitNum(I)=0
next
Do While not RsDayObj.Eof 
	VisitTime = RsDayObj("VisitTime")
	VisitDay = Day(VisitTime)
	for I=1 To DaysOfMonth
		if I=VisitDay then
			VisitNum(I)=VisitNum(I)+1
		end if
	next
	RsDayObj.MoveNext
Loop
for I=1 To DaysOfMonth
	if VisitNum(I)>=MaxVisitCount  then
		MaxVisitCount=VisitNum(I)
	end if
next
%>
<% 
	Dim VisitSize(31)
	For I=1 To DaysOfMonth
	if MaxVisitCount<>0 then
	VisitSize(I)=100*VisitNum(I)/MaxVisitCount
	else
	VisitSize(I)=0
	end if
	Next
%>
<table border=0 align="center" cellpadding=2>
	<tr>
		<td align=center>���<% =DaysOfMonth %>��ͳ��ͼ��</td>
	</tr>
	<tr>
		
    <td align=center><table border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><table border=0 align=center cellpadding=0 cellspacing=0>
              <tr> 
                <td height="25" valign=top align="right" nowrap> 
                  <% Response.Write(MaxVisitCount&"��")%>
               </td>
              </tr>
              <tr> 
                <td  height="25" valign=top align="right"  nowrap>
				<% if MaxVisitCount>3 then  
					Response.Write(Round(MaxVisitCount*0.75)&"��") 
					elseif MaxVisitCount>1  then 
					Response.Write((MaxVisitCount-1)&"��") 
					else Response.Write("&nbsp;") 
					end if
				%> 
				</td>
              </tr>
              <tr> 
                <td  height="25" valign=top  align="right" nowrap>
				<% if MaxVisitCount>3 then  
					Response.Write(Round(MaxVisitCount*0.5)&"��") 
					elseif MaxVisitCount>2  then 
					Response.Write((MaxVisitCount-2)&"��") 
					else Response.Write("&nbsp;") 
					end if
				%>
				</td>
              </tr>
              <tr>
                <td  height="25" valign=top  align="right" nowrap>
                <% if MaxVisitCount>3 then 
					 Response.Write(Round(MaxVisitCount*0.25)&"��") 
					else Response.Write("&nbsp;") 
					end if
				%>
				</td>
              </tr>
              <tr> 
                <td  height="31" valign=top  align="right" nowrap>0��</td>
              </tr>
            </table></td>
          <td valign="bottom">
<table align=center>
              <tr valign=bottom > 
                <% For I=CurrentDay+1 To DaysOfMonth %>
                <td width="15" align=center nowrap background=../../Images/Visit/tu_back.gif><img src=../../Images/Visit/bar.gif width="15" height="<% =VisitSize(I) %>" id=htav><br>
                  <% = I%>
                </td>
                <% Next %>
                <% For I=1 To CurrentDay  %>
                <td width="15" align=center nowrap background=../../Images/Visit/tu_back.gif><img src=../../Images/Visit/bar.gif width="15" height="<% =VisitSize(I) %>" id=htav><br>
                  <% = I%>
                </td>
                <% Next %>
                <td>��λ���գ�</td>
              </tr>
            </table></td>
        </tr>
      </table> </td>
	</tr>
</table>
<p>&nbsp;</p>
<%
Dim RsDaysObj
Set RsDaysObj = Server.CreateObject(G_FS_RS)
Sql="Select VisitTime From FS_FlowStatistic"
RsDaysObj.Open Sql,Conn,3,3
MaxVisitCount=0
for I=1 To 30
	VisitNum(I)=0
next
Do While not RsDaysObj.Eof 
	VisitTime = RsDaysObj("VisitTime")
	VisitDay = Day(VisitTime)
	for I=1 To 31
		if I=VisitDay then
			VisitNum(I)=VisitNum(I)+1
		end if
	next
	RsDaysObj.MoveNext
Loop
for I=1 To 31
	if VisitNum(I)>=MaxVisitCount  then
		MaxVisitCount=VisitNum(I)
	end if
next
%>
<% 
	Dim VisitsSize(31),AllCount
	AllCount=0
	For I=1 To 31
		Allcount=AllCount+VisitNum(I)
	Next
	For I=0 To 31
		if VisitNum(I)<>0 then
			VisitsSize(I)=100*VisitNum(I)/AllCount
		else
			VisitsSize(I)=0
		end if
	Next
%>
<table border=0 align="center" cellpadding=2>
	<tr>
		<td align=center>�������������ͼ��</td>
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
                <td width="15" align=center nowrap background=../../Images/Visit/tu_back.gif><img src=../../Images/Visit/bar.gif width="15" height="100" id=htav><br>
                  ��</td>
                <% For I=1 To 31 %>
                <td width="15" align=center nowrap background=../../Images/Visit/tu_back.gif><img src=../../Images/Visit/bar.gif width="15" height="<% =VisitsSize(I) %>" id=htav><br>
                  <% = I%>
                </td>
                <% Next %>
                <td>��λ���գ�</td>
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
Function GetDayNum(YearVar, MonthVar)
    Dim Temp, LeapYear, BigMonthArray, i, BigMonth
    BigMonthArray = Array(1, 3, 5, 7, 8, 10, 12)
    YearVar = CInt(YearVar)
    MonthVar = CInt(MonthVar)
    Temp = CInt(YearVar / 4)
    If YearVar = Temp * 4 Then
        LeapYear = True
    Else
        LeapYear = False
    End If
    For i = LBound(BigMonthArray) To UBound(BigMonthArray)
        If MonthVar = BigMonthArray(i) Then
            BigMonth = True
            Exit For
        Else
            BigMonth = False
        End If
    Next
    If BigMonth = True Then
        GetDayNum = 31
    Else
        If MonthVar = 2 Then
            If LeapYear = True Then
                GetDayNum = 29
            Else
                GetDayNum = 28
            End If
        Else
            GetDayNum = 30
        End If
    End If
End Function
%>
