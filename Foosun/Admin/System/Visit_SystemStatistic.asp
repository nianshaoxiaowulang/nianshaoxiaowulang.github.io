<% Option Explicit %>
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<%
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P080507") then Call ReturnError1()
%>
<html>
<head>
<title>ϵͳ/�����ͳ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<body bgcolor="#FFFFFF" topmargin="2" leftmargin="2">
<table height="26" width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
      <td height="28" class="ButtonListLeft"><div align="center"><strong>ϵͳ/�����ͳ��</strong></div></td>
  </tr>
</table>
<%
Dim RsOsObj,Sql
Set RsOsObj = Server.CreateObject(G_FS_RS)
Sql="Select OSType,ExploreType From FS_FlowStatistic"
RsOsObj.Open Sql,Conn,3,3
Dim OsType,ExploreType
Dim Num2000,NumXP,Num2003,NumNT,Num9x,NumUnix,NumOthers
Num2000=0
NumXP=0
Num2003=0
NumNT=0
Num9x=0
NumUnix=0
NumOthers=0
Dim NumIE6,NumIE5,NumNetscape,NumOpera,NumNetCaptor,NumIE4,NumOther
NumIE6=0
NumIE5=0
NumNetscape=0
NumOpera=0
NumNetCaptor=0
NumIE4=0
NumOther=0
Do While not RsOsObj.Eof
	OsType= RsOsObj("OsType")
	ExploreType= RsOsObj("ExploreType")
	Select Case OsType
	Case "Windows 2000"
		Num2000=Num2000+1
	Case "Windows XP"
		NumXP=NumXP+1
	Case "Windows 2003"
		Num2003=Num2003+1
	Case "Windows NT"
		NumNT=NumNT+1
	Case "Windows 9x"
		Num9x=Num9x+1
	Case "Unix & Unix ��"
		NumUnix=NumUnix+1
	Case "Others"
		NumOthers=NumOthers+1
	Case Else
	End Select
	Select Case ExploreType
	Case "Internet Explorer 6.x"
		NumIE6=NumIE6+1
	Case "Internet Explorer 5.x"
		NumIE5=NumIE5+1
	Case "Netscape"
		NumNetscape=NumNetscape+1
	Case "Opera"
		NumOpera=NumOpera+1
	Case "NetCaptor"
		NumNetCaptor=NumNetCaptor+1
	Case "Internet Explorer 4.x"
		NumIE4=NumIE4+1
	Case "Others"
		NumOther=NumOther+1
	Case Else
	End Select
	RsOsObj.MoveNext
Loop
%>
<%
Dim AllNum
AllNum=Num2000+NumXP+Num2003+NumNT+Num9x+NumUnix+NumOthers
Dim AllNums
AllNums=NumIE6+NumIE5+NumNetscape+NumOpera+NumNetCaptor+NumIE4+NumOther
%>
<table width="96%" border=0 cellpadding=2>
	<tr>
		<td width=45 % align=center><div align="center">ϵͳͳ��ͼ��</div></td>
		<td width=45 % align=center><div align="center">�����ͳ��ͼ��</div></td>
	</tr>
	<tr valign=top>
		<td height=22 align=center>
			<table align=center>
        <tr valign=cente>
					
          <td align=right nowap>Windows 2000</td>
					
          <td align=left background=../../Images/Visit/tu_back_2.gif valign=middle width=150><img src=../../Images/Visit/bar2.gif width="
					<% if AllNum<>0 then
						Response.Write(150*Num2000/AllNum)
						else
						Response.Write(0)
						end if
					 %>
					" height=15></td>
					<td nowap><% = Num2000 %></td>
				</tr>
				<tr valign=cente>
					
          <td align=right nowap>Windows XP</td>
					
          <td align=left background=../../Images/Visit/tu_back_2.gif valign=middle width=150><img src=../../Images/Visit/bar2.gif height=15></td>
					<td nowap><% = NumXP %></td>
				</tr>
				<tr valign=cente>
					
          <td align=right nowap>Windows 2003</td>
					
          <td align=left background=../../Images/Visit/tu_back_2.gif valign=middle width=150><img src=../../Images/Visit/bar2.gif height=15></td>
					<td nowap><% = Num2003 %></td>
				</tr>
				<tr valign=cente>
					
          <td align=right nowap>Windows NT</td>
					
          <td align=left background=../../Images/Visit/tu_back_2.gif valign=middle width=150><img src=../../Images/Visit/bar2.gif height=15></td>
					<td nowap><% = NumNT %></td>
				</tr>
				<tr valign=cente>
					
          <td align=right nowap>Windows 9x</td>
					
          <td align=left background=../../Images/Visit/tu_back_2.gif valign=middle width=150><img src=../../Images/Visit/bar2.gif height=15></td>
					<td nowap><% = Num9x %></td>
				</tr>
				<tr valign=cente>
					
          <td align=right nowap>Unix & Unix ��</td>
					
          <td align=left background=../../Images/Visit/tu_back_2.gif valign=middle width=150><img src=../../Images/Visit/bar2.gif height=15></td>
					<td nowap><% = NumUnix %></td>
				</tr>
				<tr valign=cente>
					
          <td align=right nowap>Others</td>
					
          <td align=left background=../../Images/Visit/tu_back_2.gif valign=middle width=150><img src=../../Images/Visit/bar2.gif height=15></td>
					<td nowap><% = NumOthers %></td>
				</tr>
				<tr valign=cente>
					
          <td align=right nowap>��</td>
					
          <td align=left background=../../Images/Visit/tu_back_2.gif valign=middle width=150><img src=../../Images/Visit/bar2.gif width="150" height=15></td>
					<td nowap><% = AllNum %></td>
				</tr>
			</table>
		</td>
		<td align=center>
			<table align=center>
        <tr valign=cente>
					
          <td align=right nowap>Internet Explorer 6.x</td>
					
          <td align=left background=../../Images/Visit/tu_back_2.gif valign=middle width=150><img src=../../Images/Visit/bar2.gif height=15></td>
					<td nowap><% = NumIE6 %></td>
				</tr>
				<tr valign=cente>
					
          <td align=right nowap>Internet Explorer 5.x</td>
					
          <td align=left background=../../Images/Visit/tu_back_2.gif valign=middle width=150><img src=../../Images/Visit/bar2.gif height=15></td>
					<td nowap><% = NumIE5 %></td>
				</tr>
				<tr valign=cente>
					
          <td align=right nowap>Netscape</td>
					
          <td align=left background=../../Images/Visit/tu_back_2.gif valign=middle width=150><img src=../../Images/Visit/bar2.gif height=15></td>
					<td nowap><% = NumNetscape %></td>
				</tr>
				<tr valign=cente>
					
          <td align=right nowap>Opera</td>
					
          <td align=left background=../../Images/Visit/tu_back_2.gif valign=middle width=150><img src=../../Images/Visit/bar2.gif height=15></td>
					<td nowap><% = NumOpera %></td>
				</tr>
				<tr valign=cente>
					
          <td align=right nowap>NumNetCaptor</td>
					
          <td align=left background=../../Images/Visit/tu_back_2.gif valign=middle width=150><img src=../../Images/Visit/bar2.gif height=15></td>
					<td nowap><% = NumNetCaptor %></td>
				</tr>
				<tr valign=cente>
					
          <td align=right nowap>Internet Explorer 4.x</td>
					
          <td align=left background=../../Images/Visit/tu_back_2.gif valign=middle width=150><img src=../../Images/Visit/bar2.gif height=15></td>
					<td nowap><% = NumIE4 %></td>
				</tr>
				<tr valign=cente>
					
          <td align=right nowap>Others</td>
					
          <td align=left background=../../Images/Visit/tu_back_2.gif valign=middle width=150><img src=../../Images/Visit/bar2.gif height=15></td>
					<td nowap><% = NumOther %></td>
				</tr>
				<tr valign=cente>
					
          <td align=right nowap>��</td>
					
          <td align=left background=../../Images/Visit/tu_back_2.gif valign=middle width=150><img src=../../Images/Visit/bar2.gif width="150" height=15></td>
					<td nowap><% = AllNums %></td>
				</tr>
			</table>
		</td>
	</tr>
</table>
</body>
</html>