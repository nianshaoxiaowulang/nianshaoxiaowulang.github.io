<% Option Explicit %>
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="inc/Config.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
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
Dim DBC,Conn,CollectConn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
DBC.ConnStr = CollectDBConnectionStr
Set CollectConn = DBC.OpenConnection()
Set DBC = Nothing
'�ж�Ȩ��
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P080100") then Call ReturnError1()
'�ж�Ȩ�޽���
Dim SelectPath
if SysRootDir = "" then
	SelectPath = "/" & UpFiles
else
	SelectPath = "/" & SysRootDir & "/" & UpFiles
end if
Dim Rs
if Request("Action") = "Del" then
	if Not JudgePopedomTF(Session("Name"),"P080103") then Call ReturnError1()
	if Request("id") <> "" then
		CollectConn.Execute("delete from FS_Site where ID in (" & Replace(Request("id"),"***",",") & ")")
	end If
	if Request("SiteFolderID") <> "" then
		CollectConn.Execute("delete from FS_SiteFolder where ID in (" & Replace(Request("SiteFolderID"),"***",",") & ")")
	end if
	Response.Redirect("site.asp")
	Response.End
elseif Request("Action") = "Lock" then
	if Request("LockID") <> "" then
		CollectConn.Execute("Update FS_Site Set IsLock=1 where ID in (" & Replace(Request("LockID"),"***",",") & ")")
		Response.Redirect("site.asp")
		Response.End
	end if
elseif Request("Action") = "UNLock" then
	if Request("LockID") <> "" then
		CollectConn.Execute("Update FS_Site Set IsLock=0 where ID in (" & Replace(Request("LockID"),"***",",") & ")")
		Response.Redirect("site.asp")
		Response.End
	end if
end if
if Request.Form("vs")="add" then
	if Not JudgePopedomTF(Session("Name"),"P080101") then Call ReturnError1()
    if Request.Form("SaveIMGPath") = "" OR Request.Form("SiteName")="" Or Request.Form("SysTemplet")="" or Request.Form("objURL")=""  or Request.Form("SysClass")=""  then
		Response.write"<script>alert(""����д������"");location.href=""javascript:history.back()"";</script>"
		Response.end
	end if
    Dim Sql
	Set Rs = Server.CreateObject ("ADODB.RecordSet")
	Sql = "Select * from FS_Site where 1=0"
	Rs.Open Sql,CollectConn,1,3
	Rs.AddNew
	Rs("SiteName") = NoCSSHackAdmin(Request.Form("SiteName"),"վ������")
	Rs("SysTemplet") = Request.Form("SysTemplet")
	Rs("objURL") = Request.Form("objURL")
	Rs("folder") = Request.Form("SiteFolder")
	Rs("SysClass") = Request.Form("SysClass")
	Rs("SaveIMGPath") = Request.Form("SaveIMGPath")
	if Request.Form("IsIFrame") = "1" then
		Rs("IsIFrame") = True
	else
		Rs("IsIFrame") = False
	end if
	if Request.Form("IsReverse") = "1" then
		Rs("IsReverse") = 1
	else
		Rs("IsReverse") = 0
	end if
	if Request.Form("IsScript") = "1" then
		Rs("IsScript") = True
	else
		Rs("IsScript") = False
	end if
	if Request.Form("IsClass") = "1" then
		Rs("IsClass") = True
	else
		Rs("IsClass") = False
	end if
	if Request.Form("IsFont") = "1" then
		Rs("IsFont") = True
	else
		Rs("IsFont") = False
	end if
	if Request.Form("IsSpan") = "1" then
		Rs("IsSpan") = True
	else
		Rs("IsSpan") = False
	end if
	if Request.Form("IsObject") = "1" then
		Rs("IsObject") = True
	else
		Rs("IsObject") = False
	end if
	if Request.Form("IsStyle") = "1" then
		Rs("IsStyle") = True
	else
		Rs("IsStyle") = False
	end if
	if Request.Form("IsDiv") = "1" then
		Rs("IsDiv") = True
	else
		Rs("IsDiv") = False
	end if
	if Request.Form("IsA") = "1" then
		Rs("IsA") = True
	else
		Rs("IsA") = False
	end if
	if Request.Form("Audit") = "1" then
		Rs("Audit") = True
	else
		Rs("Audit") = False
	end if
	if Request.Form("TextTF") = "1" then
		Rs("TextTF") = True
	else
		Rs("TextTF") = False
	end if
	if Request.Form("SaveRemotePic") = "1" then
		Rs("SaveRemotePic") = True
	else
		Rs("SaveRemotePic") = False
	end if
	if Request.Form("Islock") <> "" then
		Rs("Islock") = True
	else
		Rs("Islock") = False
	end if
	Rs.UpDate
	Rs.Close
	Set Rs = Nothing
	Set Conn = Nothing
	Set CollectConn = Nothing
	Response.Redirect("Site.asp")
	Response.End
elseif Request("vs")="addfolder" then
	if Not JudgePopedomTF(Session("Name"),"P080101") then Call ReturnError1()
	Dim SiteFolder,SiteFolderDetail,SqlStr
	SiteFolder = NoCSSHackAdmin(Request.Form("SiteFolder"),"վ����Ŀ")
	SiteFolderDetail = Request.Form("SiteFolderDetail")
	If SiteFolder = "" or SiteFolderDetail = "" Then
		Response.write"<script>alert(""����д������"");location.href=""javascript:history.back()"";</script>"
		Response.end
	End If
	Set Rs = Server.CreateObject ("ADODB.RecordSet")
	SqlStr = "Select * from FS_SiteFolder where 1=0"
	Rs.Open SqlStr,CollectConn,1,3
	Rs.AddNew
	Rs("SiteFolder") = SiteFolder
	Rs("SiteFolderDetail") = SiteFolderDetail
	Rs.UpDate
	Rs.Close
	Set Rs = Nothing
	Set Conn = Nothing
	Set CollectConn = Nothing
	Response.Redirect("Site.asp")
	Response.end
elseif Request("vs")="Copy" then
	if Not JudgePopedomTF(Session("Name"),"P080104") then Call ReturnError1()
	Dim SiteID,SiteFolderID,RsCopySourceObj,RsCopyObjectObj,FiledObj
	SiteID = Request("SiteID")
	SiteFolderID = Request("SiteFolderID")
	if SiteID <> "" then
		Set RsCopySourceObj = CollectConn.Execute("Select * from FS_Site where ID in (" & Replace(SiteID,"***",",") & ")")
		do while Not RsCopySourceObj.Eof
			Set RsCopyObjectObj = Server.CreateObject("ADODB.RecordSet")
			RsCopyObjectObj.Open "Select * from FS_Site where 1=0",CollectConn,3,3
			RsCopyObjectObj.AddNew
			For Each FiledObj In RsCopyObjectObj.Fields
				if LCase(FiledObj.name) <> "id" then
					RsCopyObjectObj(FiledObj.name) = RsCopySourceObj(FiledObj.name)
				end if
			Next
			RsCopyObjectObj.Update
			RsCopySourceObj.MoveNext
		Loop
		Set RsCopySourceObj = Nothing
		Set RsCopyObjectObj = Nothing
	end If
	if SiteFolderID <> "" then
		Set RsCopySourceObj = CollectConn.Execute("Select * from FS_SiteFolder where ID in (" & Replace(SiteFolderID,"***",",") & ")")
		do while Not RsCopySourceObj.Eof
			Set RsCopyObjectObj = Server.CreateObject("ADODB.RecordSet")
			RsCopyObjectObj.Open "Select * from FS_SiteFolder where 1=0",CollectConn,3,3
			RsCopyObjectObj.AddNew
			For Each FiledObj In RsCopyObjectObj.Fields
				if LCase(FiledObj.name) <> "id" then
					RsCopyObjectObj(FiledObj.name) = RsCopySourceObj(FiledObj.name)
				end if
			Next
			RsCopyObjectObj.Update
			RsCopySourceObj.MoveNext
		Loop
		Set RsCopySourceObj = Nothing
		Set RsCopyObjectObj = Nothing
	end if
	Set Conn = Nothing
	Set CollectConn = Nothing
	Response.Redirect("Site.asp")
	Response.end
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�Զ����Ųɼ���վ������</title>
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script language="JavaScript" src="../../SysJS/PublicJS.js"></script>
<% if Request("Action") <> "Addsite" and Request("Action") <> "Addsitefolder" then %>
<script language="JavaScript" src="../../SysJS/ContentMenu.js"></script>
<% end if %>
<body<% if Request("Action") <> "Addsite" and Request("Action") <> "Addsitefolder" then %> onselectstart="return false;" onClick="SelectSite();"<% end if %> leftmargin="2" topmargin="2">
<%
if Request("Action") = "Addsite" then
	Call Add()
ElseIf Request("Action") = "Addsitefolder" Then
	Call AddFolder()
ElseIf Request("Action") = "SubFolder" Then
	Call SubMain()
Else
	Call Main()
end if
Sub Main()
	Session("SessionReturnValue") = ""
	if Not JudgePopedomTF(Session("Name"),"P080100") then Call ReturnError1()
%>
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=55 align="center" alt="��Ӳɼ���Ŀ" onClick="AddSiteFolder();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">�½���Ŀ</td>
		  <td width=2 class="Gray">|</td>
		  <td width=55 align="center" alt="��Ӳɼ�վ��" onClick="AddSite();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">�½�վ��</td>
		  <td width=2 class="Gray">|</td>
		  <td width=55 align="center" alt="�޸�վ������" onClick="EditSite();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">�޸�����</td>
		  <td width=2 class="Gray">|</td>
		  <td width=55 align="center" alt="�޸�վ����" onClick="EditSiteGuide();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">�޸���</td>
		  <td width=2 class="Gray">|</td>
          <td width=35 align="center" alt="ɾ��վ��" onClick="DelSite();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">ɾ��</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="����վ��" onClick="CopySite();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="��ʼ�ɼ�" onClick="StartCollect();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">�ɼ�</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="�����ϴβɼ�" onClick="ResumeCollect();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="����" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  <td>&nbsp; </td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
  	<td colspan="5" height="2"></td>
  </tr>
  <tr> 
    <td width="19%" height="26" nowrap bgcolor="#FFFFFF" class="ButtonListLeft"> <div align="center">����</div></td>
    <td width="9%" height="26" nowrap bgcolor="#FFFFFF" class="ButtonList"> <div align="center">״̬</div></td>
    <td width="9%" height="26" bgcolor="#FFFFFF" class="ButtonList" nowrap> <div align="center">�ɼ�����ҳ</div></td>
    <td width="20%" height="26" nowrap bgcolor="#FFFFFF" class="ButtonList"> <div align="center">�ɼ�����Ŀ</div></td>
    <td width="12%" height="26" nowrap class="ButtonList"> <div align="center">��ʼ�ɼ�</div></td>
  </tr>
  <%
	Dim RsSite,SiteSql,CheckInfo
	Dim RsSiteFolder
	Set RsSiteFolder = CollectConn.Execute("select * from FS_SiteFolder order by id DESC")
	Do While not RsSiteFolder.EOF
	%>
  <tr title="վ��������Ŀ¼���������"> 
    <td height="26" nowrap> <table border="0" cellspacing="0" cellpadding="0">
        <tr> 	
          <td><img src="../../Images/Folder/folderclosed.gif" width="24" height="22"></td>
          <td nowrap><span  class="TempletItem" SiteFolderID="<% = RsSiteFolder("ID") %>" onDblClick="ChangeFolder(this);"> 
            <%= RsSiteFolder("SiteFolder")%>
            </span></td>
			 </tr>
      </table></td>
    <td nowrap> <div align="center"> &nbsp; </div></td>
    <td nowrap> <div align="center"> &nbsp; </div></td>
    <td nowrap> <div align="center"> &nbsp; </div></td>
    <td nowrap> <div align="center"> &nbsp;  </div></td>
		 </tr>
	<%
		RsSiteFolder.MoveNext
	Loop
	Set RsSiteFolder = Nothing


	Dim IsCollect,SysClassCName,RsTempObj,CollectPromptInfo
	Set RsSite = Server.CreateObject ("ADODB.RecordSet")
	SiteSql="Select * from FS_Site where folder=0 order by id desc"
		RsSite.Open SiteSql,CollectConn,1,1
	Do While not RsSite.eof
		if  RsSite("LinkHeadSetting") <> "" And  RsSite("LinkFootSetting") <> "" And RsSite("PagebodyHeadSetting") <> "" And  RsSite("PagebodyFootSetting") <> "" And  RsSite("PageTitleHeadSetting") <> "" And  RsSite("PageTitleFootSetting") <> "" then
			if RsSite("IsLock") = True then
				IsCollect = False
				CollectPromptInfo = "վ���Ѿ�������,���ܲɼ�"
			else
				IsCollect = True
				CollectPromptInfo = "���Բɼ�,�����Ƿ�������ȷ�������ܽ��вɼ�"
			end if
		else
			IsCollect = False
			CollectPromptInfo = "���ܲɼ�,���ƥ�������������"
		end if
		
		Set RsTempObj = Conn.Execute("Select ClassCName from FS_NewsClass where ClassID='" & RsSite("SysClass") & "'")
		if Not RsTempObj.Eof then
			SysClassCName = RsTempObj("ClassCName")
		else
			SysClassCName = "��Ŀ������"
			IsCollect = False
			CollectPromptInfo = "Ŀ����Ŀ������,���ܲɼ�"
		end if
		Set RsTempObj = Nothing
	%>
	<tr title="<% = CollectPromptInfo %>"> 
    <td height="26" nowrap> <table border="0" cellspacing="0" cellpadding="0">
        <tr> 
		<td><img src="../../Images/Station.gif" width="24" height="22"></td>
          <td nowrap><span IsCollect="<% = IsCollect %>" class="TempletItem" SiteID="<% = RsSite("ID") %>"> 
            <% = RsSite("SiteName") %>
            </span></td>
		 </tr>
      </table></td>
    <td nowrap> <div align="center"> 
        <%
		if RsSite("IsLock") = True then
			Response.Write("����")
		ElseIf IsCollect = False Then
			Response.Write("��Ч")
		else
			Response.Write("��Ч")
		end if
		%>
      </div></td>
    <td nowrap> <div align="center"><a href="<% = RsSite("objURL") %>" target="_blank"><img src="Images/objpage.gif" alt="�������" width="20" height="20" border="0"></a></div></td>
    <td nowrap> <div align="center"> 
        <% = SysClassCName %>
      </div></td>
    <td nowrap> <div align="center"> 
        <% if IsCollect = true then %>
        <div align="center"><span onClick="StartOneSiteCollect('<% = RsSite("Id") %>');" style="cursor:hand;"><img src="Images/collect.gif" width="20" height="20" border="0"></span></div>
        <% else %>
        <div align="center"><img src="Images/uncollect.gif" width="20" height="20" border="0"></div>
        <% end if %>
      </div></td>
  </tr>
  <%
		RsSite.MoveNext
	loop
	RsSite.close
	Set RsSite = Nothing
end Sub

Sub	SubMain()
	Session("SessionReturnValue") = ""
	if Not JudgePopedomTF(Session("Name"),"P080100") then Call ReturnError1()
%>
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=55 align="center" alt="��Ӳɼ���Ŀ" onClick="AddSiteFolder();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">�½���Ŀ</td>
		  <td width=2 class="Gray">|</td>
		  <td width=55 align="center" alt="��Ӳɼ�վ��" onClick="AddSubSite();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">�½�վ��</td>
		  <td width=2 class="Gray">|</td>
		  <td width=55 align="center" alt="�޸�վ������" onClick="EditSite();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">�޸�����</td>
		  <td width=2 class="Gray">|</td>
		  <td width=55 align="center" alt="�޸�վ����" onClick="EditSiteGuide();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">�޸���</td>
		  <td width=2 class="Gray">|</td>
          <td width=35 align="center" alt="ɾ��վ��" onClick="DelSite();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">ɾ��</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="����վ��" onClick="CopySite();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="��ʼ�ɼ�" onClick="StartCollect();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">�ɼ�</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="�����ϴβɼ�" onClick="ResumeCollect();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">�̲�</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="����" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  <td>&nbsp; </td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <script language="javascript">
	function AddSubSite()
	{
		window.location="site.asp?Action=Addsite&FolderID="+<%=Request("FolderID")%>;
	}
  </script>
  <tr>
  	<td colspan="5" height="2"></td>
  </tr>
  <tr> 
    <td width="19%" height="26" nowrap bgcolor="#FFFFFF" class="ButtonListLeft"> <div align="center">����</div></td>
    <td width="9%" height="26" nowrap bgcolor="#FFFFFF" class="ButtonList"> <div align="center">״̬</div></td>
    <td width="9%" height="26" bgcolor="#FFFFFF" class="ButtonList" nowrap> <div align="center">�ɼ�����ҳ</div></td>
    <td width="20%" height="26" nowrap bgcolor="#FFFFFF" class="ButtonList"> <div align="center">�ɼ�����Ŀ</div></td>
    <td width="12%" height="26" nowrap class="ButtonList"> <div align="center">��ʼ�ɼ�</div></td>
  </tr>
  <tr title="��������ϼ�Ŀ¼"> 
    <td height="26" nowrap> <table border="0" cellspacing="0" cellpadding="0">
        <tr> 	
          <td><img src="../../Images/arrow.gif" width="24" height="22"></td>
          <td nowrap><span  class="TempletItem" onDblClick="backTop()" onclick="backTop()" style="cursor:hand">�����ϼ�</span></td>
			 </tr>
      </table></td>
    <td nowrap> <div align="center"> &nbsp; </div></td>
    <td nowrap> <div align="center"> &nbsp; </div></td>
    <td nowrap> <div align="center"> &nbsp; </div></td>
    <td nowrap> <div align="center"> &nbsp;  </div></td>
 </tr>
  <%
	Dim RsSite,SiteSql,CheckInfo
	Dim IsCollect,SysClassCName,RsTempObj,CollectPromptInfo
	Set RsSite = Server.CreateObject ("ADODB.RecordSet")
	SiteSql="Select * from FS_Site where folder="& CLng(Request("FolderID")) &" order by id desc"
		RsSite.Open SiteSql,CollectConn,1,1
	Do While not RsSite.eof
		if  RsSite("LinkHeadSetting") <> "" And  RsSite("LinkFootSetting") <> "" And RsSite("PagebodyHeadSetting") <> "" And  RsSite("PagebodyFootSetting") <> "" And  RsSite("PageTitleHeadSetting") <> "" And  RsSite("PageTitleFootSetting") <> "" then
			if RsSite("IsLock") = True then
				IsCollect = False
				CollectPromptInfo = "վ���Ѿ�������,���ܲɼ�"
			else
				IsCollect = True
				CollectPromptInfo = "���Բɼ�,�����Ƿ�������ȷ�������ܽ��вɼ�"
			end if
		else
			IsCollect = False
			CollectPromptInfo = "���ܲɼ�,���ƥ�������������"
		end if
		
		Set RsTempObj = Conn.Execute("Select ClassCName from FS_NewsClass where ClassID='" & RsSite("SysClass") & "'")
		if Not RsTempObj.Eof then
			SysClassCName = RsTempObj("ClassCName")
		else
			SysClassCName = "��Ŀ������"
			IsCollect = False
			CollectPromptInfo = "Ŀ����Ŀ������,���ܲɼ�"
		end if
		Set RsTempObj = Nothing
	%>
	<tr title="<% = CollectPromptInfo %>"> 
    <td height="26" nowrap> <table border="0" cellspacing="0" cellpadding="0">
        <tr> 
		<td><img src="../../Images/Station.gif" width="24" height="22"></td>
          <td nowrap><span IsCollect="<% = IsCollect %>" class="TempletItem" SiteID="<% = RsSite("ID") %>"> 
            <% = RsSite("SiteName") %>
            </span></td>
		 </tr>
      </table></td>
    <td nowrap> <div align="center"> 
        <%
		if RsSite("IsLock") = True then
			Response.Write("����")
		ElseIf IsCollect = False Then
			Response.Write("��Ч")
		else
			Response.Write("��Ч")
		end if
		%>
      </div></td>
    <td nowrap> <div align="center"><a href="<% = RsSite("objURL") %>" target="_blank"><img src="Images/objpage.gif" alt="�������" width="20" height="20" border="0"></a></div></td>
    <td nowrap> <div align="center"> 
        <% = SysClassCName %>
      </div></td>
    <td nowrap> <div align="center"> 
        <% if IsCollect = true then %>
        <div align="center"><span onClick="StartOneSiteCollect('<% = RsSite("Id") %>');" style="cursor:hand;"><img src="Images/collect.gif" width="20" height="20" border="0"></span></div>
        <% else %>
        <div align="center"><img src="Images/uncollect.gif" width="20" height="20" border="0"></div>
        <% end if %>
      </div></td>
  </tr>
  <%
		RsSite.MoveNext
	loop
	RsSite.close
	Set RsSite = Nothing
end sub

Sub Add()
	if Not JudgePopedomTF(Session("Name"),"P080101") then Call ReturnError1()
	Dim TempletDirectory
	if SysRootDir <> "" then
		TempletDirectory = "/" & SysRootDir & "/" & TempletDir
	else
		TempletDirectory = "/" & TempletDir
	end if
%>
<form name="AddSiteForm" method="post" action="">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=35 align="center" alt="����" onClick="document.AddSiteForm.submit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="����" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
            <td>&nbsp; <input name="vs" type="hidden" id="vs2" value="add"> </td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
  <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="dddddd">
    <tr bgcolor="#FFFFFF"> 
      <td width="100" height="26"> 
        <div align="center">�ɼ�վ������</div></td>
      <td> 
        <input name="SiteName" style="width:100%;" type="text" id="SiteName2"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26"> 
        <div align="center">�ɼ�վ�����</div></td>
      <td> 
        <select name="SiteFolder" style="width:100%;" id="SiteFolder">
		<option value="0">����Ŀ</option>
          <% = FolderList %>
        </select></td>
    </tr>
	<tr bgcolor="#FFFFFF"> 
      <td height="26"> 
        <div align="center">���Ŀ����Ŀ</div></td>
      <td> 
        <select name="SysClass" style="width:100%;" id="select">
          <% = ClassList %>
        </select></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26"> 
        <div align="center">�ɼ�����ҳ</div></td>
      <td> 
        <input style="width:100%;" name="objURL" type="text" id="objURL" value="http://"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26"> 
        <div align="center">��������</div></td>
      <td> 
        <input readonly name="SysTemplet" type="text" id="SysTemplet" style="width:80%;"> 
        <input name="Submitaaa" type="button" id="Submitaaa" value="ѡ��ģ��" onClick="OpenWindowAndSetValue('../../FunPages/SelectFileFrame.asp?CurrPath=<% = TempletDirectory %>',400,300,window,document.AddSiteForm.SysTemplet);"> 
        <div align="right"></div></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26"> 
        <div align="center">�ɼ�����</div></td>
      <td>���� 
        <input name="islock" type="checkbox" id="islock" value="1">
        ����Զ��ͼƬ 
        <input type="checkbox" name="SaveRemotePic" value="1">
        �����Ƿ��Ѿ���� 
        <input name="Audit" type="checkbox" value="1" checked>
        �Ƿ���ɼ� 
        <input name="IsReverse" type="checkbox" id="IsReverse" value="1"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26"> 
        <div align="center">����ͼƬ·��</div></td>
      <td> 
        <input type="text" readonly name="SaveIMGPath" style="width:80%;" value="/<% = UpFiles & "/" & BeyondPicDir %>">
        <input name="Submit111" id="SelectPath" type="button" value="ѡ��·��" onClick="OpenWindowAndSetValue('../../FunPages/SelectPathFrame.asp?CurrPath=<% = SelectPath %>',400,300,window,document.AddSiteForm.SaveIMGPath);"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26"> 
        <div align="center">����ѡ��</div></td>
      <td>HTML 
        <input type="checkbox" name="TextTF" value="1">
        STYLE <input type="checkbox" name="IsStyle" value="1">
        DIV<input type="checkbox" name="IsDiv" value="1">
        A<input type="checkbox" name="IsA" value="1">
        CLASS<input type="checkbox" name="IsClass" value="1">
        FONT<input type="checkbox" name="IsFont" value="1">
        SPAN<input type="checkbox" name="IsSpan" value="1">
        OBJECT<input type="checkbox" name="IsObject" value="1">
        IFRAME<input type="checkbox" name="IsIFrame" value="1">
        SCRIPT<input type="checkbox" name="IsScript" value="1">
        </td>
    </tr>
</table>
</form>
<% End Sub

Sub AddFolder()
	if Not JudgePopedomTF(Session("Name"),"P080101") then Call ReturnError1()
	Dim TempletDirectory
	if SysRootDir <> "" then
		TempletDirectory = "/" & SysRootDir & "/" & TempletDir
	else
		TempletDirectory = "/" & TempletDir
	end if
%>
<form name="AddSiteFolderForm" method="post" action="">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=35 align="center" alt="����" onClick="document.AddSiteFolderForm.submit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="����" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
          <td>&nbsp; <input name="vs" type="hidden" id="vs2" value="addfolder"> </td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
  <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="dddddd">
    <%
	Dim SiteFolderID,RsSiteFolder
	SiteFolderID = Request("SiteFolderID")
	If SiteFolderID<>"" Then
		Set RsSiteFolder = CollectConn.Execute("select * from FS_SiteFolder where ID=" & SiteFolderID)
		If RsSiteFolder.EOF Then
			Response.write "<script>alert('����Ŀ�����ڣ�');history.back();</script>"
			Response.End
		End If
	%>
	<tr bgcolor="#FFFFFF"> 
      <td width="100" height="26"> 
        <div align="center">��Ŀ����:</div></td>
	  <td> 
        <input style="width:100%" type="text" name="SiteFolder" value="<%=RsSiteFolder("SiteFolder")%>"></td>
	</tr>
	<tr bgcolor="#FFFFFF"> 
      <td width="100" height="26"> 
        <div align="center">��Ŀ˵��:</div></td>
	  <td> 
        <textarea style="width:100%" name="SiteFolderDetail" rows="10"><%=RsSiteFolder("SiteFolderDetail")%></textarea></td>
	</tr>
	<%
	Else
	%>
	<tr bgcolor="#FFFFFF"> 
      <td width="100" height="26"> 
        <div align="center">��Ŀ����:</div></td>
	  <td> 
        <input style="width:100%" type="text" name="SiteFolder"></td>
	</tr>
	<tr bgcolor="#FFFFFF"> 
      <td width="100" height="26"> 
        <div align="center">��Ŀ˵��:</div></td>
	  <td> 
        <textarea style="width:100%" name="SiteFolderDetail" rows="10"></textarea></td>
	</tr>
	<%
	End If
	%>
</table>
</form>
<% End Sub %>

</body>
</html>
<%
Function ClassList()
	Dim ClassListObj
	Set ClassListObj = Conn.Execute("Select ClassID,ClassCName from FS_NewsClass where ParentID='0' order by ClassID desc")
	do while Not ClassListObj.Eof
		ClassList = ClassList & "<option value="&ClassListObj("ClassID")&"" & ">" & ClassListObj("ClassCName") & "</option><br>"
		ClassList = ClassList & ChildClassList(ClassListObj("ClassID"),"")
		ClassListObj.MoveNext	
	loop
	ClassListObj.Close
	Set ClassListObj = Nothing
End Function

Function FolderList()
	Dim FolderListObj,StrSelected
	Set FolderListObj = Collectconn.Execute("Select * from FS_SiteFolder order by ID desc")
	do while Not FolderListObj.Eof
		If CInt(Request("FolderID"))=FolderListObj("ID") Then
			StrSelected="selected"
		Else
			StrSelected=""
		End If
		FolderList = FolderList & "<option value="&FolderListObj("ID")&" " & StrSelected & ">&nbsp;&nbsp;|--" & FolderListObj("SiteFolder") & "</option><br>"
		FolderListObj.MoveNext	
	loop
	FolderListObj.Close
	Set FolderListObj = Nothing
End Function

Function ChildClassList(ClassID,Temp)
	Dim TempRs,TempStr
	Set TempRs = Conn.Execute("Select ClassID,ClassCName from FS_NewsClass where ParentID='" & ClassID & "' order by ClassID desc")
	TempStr = Temp & " |- "
	do while Not TempRs.Eof
		ChildClassList = ChildClassList & "<option value="&TempRs("ClassID")&"" & ">" & TempStr & TempRs("ClassCName") & "</option><br>"
		ChildClassList = ChildClassList & ChildClassList(TempRs("ClassID"),TempStr)
		TempRs.MoveNext
	loop
	TempRs.Close
	Set TempRs = Nothing
End Function

Set Conn = Nothing
Set CollectConn = Nothing
%>
<script language="JavaScript">
var DocumentReadyTF=false;
var ListObjArray = new Array();
var ContentMenuArray=new Array();
function document.onreadystatechange()
{
	if (DocumentReadyTF) return;
	<% if Request("Action") <> "Addsite" and Request("Action") <> "AddsiteFolder" then %>
	IntialListObjArray();
	InitialContentListContentMenu();
	<% end if %>
	DocumentReadyTF=true;
}
function InitialContentListContentMenu()
{
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.AddSite();",'�½�վ��','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.EditSite();",'�޸�','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.DelSite();",'ɾ��','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('seperator','','');
//	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.StartCollect();','�ɼ�','disabled');
//	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.ResumeCollect();','����','disabled');
//	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.StartCollectAtServer();','��̨�ɼ�','disabled');
//	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.EditSiteGuide();','��','disabled');
//	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.CopySite();','����','disabled');
//	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('seperator','','');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.Lock(true);','����','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.Lock(false);','����','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('seperator','','');
	//ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.Export();','����','disabled');
	//ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.Import();','����','disabled');
	//ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('seperator','','');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('top.GetEkMainObject().location.reload();','ˢ��','');
//	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('prompt(\'��ҳ��·������\',\'<%=Request.ServerVariables("SCRIPT_NAME")%>\');','·������','');
}
function ContentMenuFunction(ExeFunction,Description,EnabledStr)
{
	this.ExeFunction=ExeFunction;
	this.Description=Description;
	this.EnabledStr=EnabledStr;
}
function ContentMenuShowEvent()
{
	ChangeContentMenuStatus();
}
function ChangeContentMenuStatus()
{
	var EventObjInArray=false,SelectContent='',SelectFolder='',DisabledContentMenuStr='';
	for (var i=0;i<ListObjArray.length;i++)
	{
		if (event.srcElement==ListObjArray[i].Obj)
		{
			if (ListObjArray[i].Selected==true) EventObjInArray=true;
			break;
		}
	}
	for (var i=0;i<ListObjArray.length;i++)
	{
		if (event.srcElement==ListObjArray[i].Obj)
		{
			ListObjArray[i].Obj.className='TempletSelectItem';
			ListObjArray[i].Selected=true;
			if (ListObjArray[i].Obj.SiteID!=null)
			{
				if (SelectContent=='') SelectContent=ListObjArray[i].Obj.SiteID;
				else SelectContent=SelectContent+'***'+ListObjArray[i].Obj.SiteID;
			}
			if (ListObjArray[i].Obj.SiteFolderID!=null)
			{
				if (SelectFolder=='') SelectFolder=ListObjArray[i].Obj.SiteFolderID;
				else SelectFolder=SelectFolder+'***'+ListObjArray[i].Obj.SiteFolderID;
			}
		}
		else
		{
			if (!EventObjInArray)
			{
				ListObjArray[i].Obj.className='TempletItem';
				ListObjArray[i].Selected=false;
			}
			else
			{
				if (ListObjArray[i].Selected==true)
				{
					if (ListObjArray[i].Obj.SiteID!=null)
					{
						if (SelectContent=='') SelectContent=ListObjArray[i].Obj.SiteID;
						else SelectContent=SelectContent+'***'+ListObjArray[i].Obj.SiteID;
					}
					if (ListObjArray[i].Obj.SiteFolderID!=null)
					{
						if (SelectFolder=='') SelectFolder=ListObjArray[i].Obj.SiteFolderID;
						else SelectFolder=SelectFolder+'***'+ListObjArray[i].Obj.SiteFolderID;
					}
				}
			}
		}
	}
	if (SelectContent=='' && SelectFolder=='') DisabledContentMenuStr=',�޸�,ɾ��,��,����,�ɼ�,����,��̨�ɼ�,����������,����,';
	else if (SelectContent!='' && SelectFolder!='')
	{
		DisabledContentMenuStr=',�޸�,��,�ɼ�,����,��̨�ɼ�,����������,����,';
	}
	else if (SelectContent!='')
	{
		if (SelectContent.indexOf('***')==-1) DisabledContentMenuStr='';
		else DisabledContentMenuStr=',�޸�,��,����,';
	}
	else
	{		
		if (SelectFolder.indexOf('***')==-1) DisabledContentMenuStr=',��,�ɼ�,����,��̨�ɼ�,����������,����,';
		else DisabledContentMenuStr=',�޸�,��,�ɼ�,����,��̨�ɼ�,����������,����,';
	}
	for (var i=0;i<ContentMenuArray.length;i++)
	{
		if (DisabledContentMenuStr.indexOf(ContentMenuArray[i].Description)!=-1) ContentMenuArray[i].EnabledStr='disabled';
		else  ContentMenuArray[i].EnabledStr='';
	}
}
function FolderFileObj(Obj,Index,Selected)
{
	this.Obj=Obj;
	this.Index=Index;
	this.Selected=Selected;
}
function IntialListObjArray()
{
	var CurrObj=null,j=1;
	for (var i=0;i<document.all.length;i++)
	{
		CurrObj=document.all(i);
		if (CurrObj.SiteID!=null)
		{
			ListObjArray[ListObjArray.length]=new FolderFileObj(CurrObj,j,false);
			j++;
		}
		if (CurrObj.SiteFolderID!=null)
		{
			ListObjArray[ListObjArray.length]=new FolderFileObj(CurrObj,j,false);
			j++;
		}
	}
}
function SelectSite()
{
	var el=event.srcElement;
	var i=0;
	if ((event.ctrlKey==true)||(event.shiftKey==true))
	{
		if (event.ctrlKey==true)
		{
			for (i=0;i<ListObjArray.length;i++)
			{
				if (el==ListObjArray[i].Obj)
				{
					if (ListObjArray[i].Selected==false)
					{
						ListObjArray[i].Obj.className='TempletSelectItem';
						ListObjArray[i].Selected=true;
					}
					else
					{
						ListObjArray[i].Obj.className='TempletItem';
						ListObjArray[i].Selected=false;
					}
				}
			}
		}
		if (event.shiftKey==true)
		{
			var MaxIndex=0,ObjInArray=false,EndIndex=0,ElIndex=-1;
			for (i=0;i<ListObjArray.length;i++)
			{
				if (ListObjArray[i].Selected==true)
				{
					if (ListObjArray[i].Index>=MaxIndex) MaxIndex=ListObjArray[i].Index;
				}
				if (el==ListObjArray[i].Obj)
				{
					ObjInArray=true;
					ElIndex=i;
					EndIndex=ListObjArray[i].Index;
				}
			}
			if (ElIndex>MaxIndex)
				for (i=MaxIndex-1;i<EndIndex;i++)
				{
					ListObjArray[i].Obj.className='TempletSelectItem';
					ListObjArray[i].Selected=true;
				}
			else
			{
				for (i=EndIndex;i<MaxIndex-1;i++)
				{	
					ListObjArray[i].Obj.className='TempletSelectItem';
					ListObjArray[i].Selected=true;
				}
				ListObjArray[ElIndex].Obj.className='TempletSelectItem';
				ListObjArray[ElIndex].Selected=true;
			}
		}
	}
	else
	{
		for (i=0;i<ListObjArray.length;i++)
		{
			if (el==ListObjArray[i].Obj)
			{
				ListObjArray[i].Obj.className='TempletSelectItem';
				ListObjArray[i].Selected=true;
			}
			else
			{
				ListObjArray[i].Obj.className='TempletItem';
				ListObjArray[i].Selected=false;
			}
		}
	}
}
function AddSiteFolder()
{
	location='?Action=Addsitefolder';
}
function AddSite()
{
	location='?Action=Addsite';
}

function Lock(Falg)
{
	var SelectedSite='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.SiteID!=null)
			{
				if (SelectedSite=='') SelectedSite=ListObjArray[i].Obj.SiteID;
				else  SelectedSite=SelectedSite+'***'+ListObjArray[i].Obj.SiteID;
			}
		}
	}
	if (SelectedSite!='')
	{
		if (Falg) location='?Action=Lock&LockID='+SelectedSite;
		else location='?Action=UNLock&LockID='+SelectedSite;
	}
	else alert('��ѡ��վ��');
}
function EditSite()
{
	var SelectedSite='',SelectedSiteFolder='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.SiteID!=null)
			{
				if (SelectedSite=='') SelectedSite=ListObjArray[i].Obj.SiteID;
				else  SelectedSite=SelectedSite+'***'+ListObjArray[i].Obj.SiteID;
			}
			if (ListObjArray[i].Obj.SiteFolderID!=null)
			{
				if (SelectedSiteFolder=='') SelectedSiteFolder=ListObjArray[i].Obj.SiteFolderID;
				else  SelectedSiteFolder=SelectedSiteFolder+'***'+ListObjArray[i].Obj.SiteFolderID;
			}
		}
	}
	if (SelectedSite!='')
	{
		SelectedSiteFolder='';
		if (SelectedSite.indexOf('***')==-1) location='SitemodifyOne.asp?SiteID='+SelectedSite;
		else alert('��ѡ��һ��վ��');
	}
	else if (SelectedSiteFolder!='')
	{
		if (SelectedSiteFolder.indexOf('***')==-1)
		location='site.asp?Action=Addsitefolder&SiteFolderID='+SelectedSiteFolder;
		else alert('��ѡ��һ��Ŀ¼');
	}
	else
	{
		alert('��ѡ��վ�������Ŀ');
	} 

}
function EditSiteGuide()
{
	var SelectedSite='',SelectedSiteFolder='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.SiteID!=null)
			{
				if (SelectedSite=='') SelectedSite=ListObjArray[i].Obj.SiteID;
				else  SelectedSite=SelectedSite+'***'+ListObjArray[i].Obj.SiteID;
			}
			else if (ListObjArray[i].Obj.SiteFolderID!=null)
			{
				if (SelectedSiteFolder=='') SelectedSiteFolder=ListObjArray[i].Obj.SiteFolderID;
				else  SelectedSiteFolder=SelectedSiteFolder+'***'+ListObjArray[i].Obj.SiteFolderID;
			}
		}
	}
	if (SelectedSite!='')
	{
		if (SelectedSite.indexOf('***')==-1) location='Sitemodify.asp?SiteID='+SelectedSite;
		else alert('��ѡ��һ��վ��');
	}
	else alert('��ѡ��һ��վ��');
}
function DelSite()
{
	var SelectedSite='',SelectedSiteFolder='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.SiteID!=null)
			{
				if (SelectedSite=='') SelectedSite=ListObjArray[i].Obj.SiteID;
				else  SelectedSite=SelectedSite+'***'+ListObjArray[i].Obj.SiteID;
			}
			if (ListObjArray[i].Obj.SiteFolderID!=null)
			{
				if (SelectedSiteFolder=='') SelectedSiteFolder=ListObjArray[i].Obj.SiteFolderID;
				else  SelectedSiteFolder=SelectedSiteFolder+'***'+ListObjArray[i].Obj.SiteFolderID;
			}
		}
	}
	if (SelectedSite!='' || SelectedSiteFolder!='')
	{
		if (confirm('ȷ��Ҫɾ����')==true)
		window.location='?action=Del&Id='+SelectedSite+'&SiteFolderID='+SelectedSiteFolder;
	}
	else alert('��ѡ��վ�����Ŀ');
}
function CopySite()
{
	var SelectedSite='',SelectedSiteFolder='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.SiteID!=null)
			{
				if (SelectedSite=='') SelectedSite=ListObjArray[i].Obj.SiteID;
				else  SelectedSite=SelectedSite+'***'+ListObjArray[i].Obj.SiteID;
			}
			if (ListObjArray[i].Obj.SiteFolderID!=null)
			{
				if (SelectedSiteFolder=='') SelectedSiteFolder=ListObjArray[i].Obj.SiteFolderID;
				else  SelectedSiteFolder=SelectedSiteFolder+'***'+ListObjArray[i].Obj.SiteFolderID;
			}
		}
	}
	if (SelectedSite!='' || SelectedSiteFolder!='')
	{
		location='Site.asp?vs=Copy&SiteID='+SelectedSite+'&SiteFolderID='+SelectedSiteFolder;
	}
	else alert('��ѡ��վ�������Ŀ');
}
function StartCollect()
{
	var SelectedSite='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.SiteID!=null)
			{
				if (ListObjArray[i].Obj.IsCollect=='True')
				{
					if (SelectedSite=='') SelectedSite=ListObjArray[i].Obj.SiteID;
					else  SelectedSite=SelectedSite+'***'+ListObjArray[i].Obj.SiteID;
				}
			}
		}
	}
	if (SelectedSite!='')
	{
		Num = InsertScript();
	//	alert(Num);
		if (Num!='back'&& Num!='0')
		{
			if (Num==""||Num==null)
			{
				Num="allNews"
			}
			location='Collecting.asp?SiteID='+SelectedSite+'&Num='+Num;
		}
	}
	else alert('��ѡ��վ�㣬��������������');
}
function StartOneSiteCollect(ID)
{
	Num = InsertScript();
	//alert(Num);
	if (Num!='back'&& Num!='0')
	{
		if (Num==""||Num==null)
		{
			Num="allNews"
		}
		location='Collecting.asp?SiteID='+ID+'&Num='+Num;
	}
}
function StartCollectAtServer()
{
	var SelectedSite='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.SiteID!=null)
			{
				if (ListObjArray[i].Obj.IsCollect=='True')
				{
					if (SelectedSite=='') SelectedSite=ListObjArray[i].Obj.SiteID;
					else  SelectedSite=SelectedSite+'***'+ListObjArray[i].Obj.SiteID;
				}
			}
		}
	}
	if (SelectedSite!='')
	{
		if (confirm(CopyRightStr))
		{
			var ObjStr="Microsoft.XMLHTTP"
			var HTTP=new ActiveXObject(ObjStr);
			HTTP.onreadystatechange=function()
			{
				if (HTTP.readyState == 4) // �������
				{
				alert(HTTP.status);
					if (HTTP.status == 200) // ���سɹ�
					{
						alert (HTTP.responseText);
					}
				}
			} 
			HTTP.open("get",'CollectingAtServer.asp?SiteID='+SelectedSite,true);
			HTTP.send();
			HTTP=null;
		}
	}
	else alert('��ѡ��վ��');
}

function ResumeCollect()
{
	var SelectedSite='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.SiteID!=null)
			{
				if (ListObjArray[i].Obj.IsCollect=='True')
				{
					if (SelectedSite=='') SelectedSite=ListObjArray[i].Obj.SiteID;
					else  SelectedSite=SelectedSite+'***'+ListObjArray[i].Obj.SiteID;
				}
			}
		}
	}
	if (SelectedSite!='')
	{
		if(confirm("      ��ӭʹ�÷�Ѷ���Ųɼ�ϵͳ�������Ƶ���Ȩ�������Ѷ�Ƽ���չ���޹�˾�޹أ�ȷ��Ҫ�����ɼ���"))location='Collecting.asp?CollectType=ResumeCollect&SiteID='+SelectedSite;		
	}
	else alert('��ѡ��վ��');
}

function Export()
{
	var SelectedSite='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.SiteID!=null)
			{
				if (SelectedSite=='') SelectedSite=ListObjArray[i].Obj.SiteID;
				else  SelectedSite=SelectedSite+'***'+ListObjArray[i].Obj.SiteID;
			}
		}
	}
	if (SelectedSite!='')
	{
		location='Export.asp?ID='+SelectedSite;
	}
	else alert('��ѡ��վ��');
}
function Import()
{
	location='Import.asp';
}

function InsertScript()
{
	var ReturnValue='';
	ReturnValue=showModalDialog("NewsNum.asp",window,'dialogWidth:260pt;dialogHeight:120pt;status:no;help:no;scroll:no;');
	return ReturnValue;
}
function ChangeFolder(Obj)
{
	location.href='site.asp?Action=SubFolder&FolderID='+Obj.SiteFolderID;
}
function backTop()
{
	location.href = 'site.asp'
}
</script>