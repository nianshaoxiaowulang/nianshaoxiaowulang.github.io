<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
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
<!--#include file="../refresh/function.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P010514") then Call ReturnError()
Dim NewsID,DownLoadID,OperateType,TempStr,TargetClassID
DownLoadID = Request("DownLoadID")
NewsID = Cstr(Request("NewsID"))
TargetClassID=request.form("ClassID")
If TargetClassID="" then 
	Dim TempClassListStr
		TempClassListStr = ClassList
	Function ClassList()
		Dim Rs
		Set Rs = Conn.Execute("select ClassID,ClassCName from FS_newsclass where ParentID = '0' and delflag=0 and isoutclass=0 order by AddTime desc")
		do while Not Rs.Eof
			ClassList = ClassList & "<option value="&Rs("ClassID") & ">" & Rs("ClassCName") & chr(13)
			ClassList = ClassList & ChildClassList(Rs("ClassID"),"")
			Rs.MoveNext	
		loop
		Rs.Close
		Set Rs = Nothing
	End Function
	Function ChildClassList(ClassID,Temp)
		Dim TempRs,TempStr
		Set TempRs = Conn.Execute("Select ClassID,ClassCName,ChildNum from FS_NewsClass where ParentID = '" & ClassID & "' and delflag=0 and isoutclass=0 order by AddTime desc ")
		TempStr = Temp & " - "
		do while Not TempRs.Eof
			ChildClassList = ChildClassList & "<option value="&TempRs("ClassID")& ">" & TempStr & TempRs("ClassCName") & "</option>" & chr(13)
			ChildClassList = ChildClassList & ChildClassList(TempRs("ClassID"),TempStr)
			TempRs.MoveNext
		loop
		TempRs.Close
		Set TempRs = Nothing
	End Function
	if NewsID <> "" OR DownLoadID <> "" then
	
	Else
		Response.Write("<script>alert(""�������ݴ���"");window.close();</script>")
		response.end
	end if 
	%>
	<html>
	<head>
	<link href="../../../CSS/ModeWindow.css" rel="stylesheet" type="text/css">
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<title>��Ŀת��</title>
	</head>
	<body leftmargin="0" topmargin="0">
	  <form action="?downloadID=<%=DownLoadID%>&NewsId=<%=NewsID%>" method="post" name="ClassForm">
	  <table width="100%">
	  <tr height="30" valign="bottom">
		<td width="70%" align="right">ѡ����Ҫת��������Ŀ��
		</td>
		<td align="left"><select name="ClassID">
		<% =TempClassListStr %>
		</select>
		</td>
		</tr>
		<tr height="20">
		<td>
		  </td>
	  </tr>
		<tr>
		<td align="center" colspan="2">
		<input name="NumClass"  type="submit" id="NumClass" value="ȷ ��">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name="CloseOk"  type="button" id="NumClass" value="�� ��" onClick="window.close();">
		  </td>
	  </tr>
	  </table>
	  </form>
	</body>
	</html>
	<%
else
	if NewsID <> "" then
		MoveNewsFile NewsID,"1",TargetClassID
		NewsID = Replace(NewsID,"***","','")
			Conn.Execute("Update FS_News Set ClassID='" & TargetClassID & "' where NewsID in ('" & NewsID & "')")
	end if
	if DownLoadID <> "" then
		MoveNewsFile DownLoadID,"2",TargetClassID
		DownLoadID = Replace(DownLoadID,"***","','")
			Conn.Execute("Update FS_DownLoad Set ClassID='" & TargetClassID & "' where DownLoadID in ('" & DownLoadID & "')")
	end if
	Response.Write("<script>window.close();</script>")
end if
%>