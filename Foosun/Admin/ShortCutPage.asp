<% Option Explicit %>
<!--#include file="../../Inc/Const.asp" -->
<!--#include file="../../Inc/Cls_DB.asp" -->
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
<!--#include file="../../Inc/Session.asp" -->
<!--#include file="../../Inc/CheckPopedom.asp" -->
<%
Dim RsMenuConfigObj,HaveValueTF
Set RsMenuConfigObj = Conn.execute("Select IsShop From FS_Config")
if RsMenuConfigObj("IsShop") = 1 then
	HaveValueTF = True
Else
	HaveValueTF = False
End if
Set RsMenuConfigObj = Nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��ݲ˵�</title>
</head>
<link href="../../CSS/FS_css.css" rel="stylesheet">
<body leftmargin="2" onselectstart="return false;">
<table width="100%" border="0" cellspacing="0" cellpadding="2">
  <%
if JudgePopedomTF(Session("Name"),"P040100") then
%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
    <td width="50%"><table border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span class="TempletItem" onClick="ClickBtn(this,'SysAdminGroup')">����Ա��</span></td>
        </tr>
      </table></td>
  </tr>
  <%
end if
if JudgePopedomTF(Session("Name"),"P040200") then
%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="AdsManage" style="display:;"> 
    <td><table border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span class="TempletItem" onClick="ClickBtn(this,'SysAdminList')">����Ա</span></td>
        </tr>
      </table></td>
  </tr>
  <%
end if
if JudgePopedomTF(Session("Name"),"P040300") then
%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
    <td><table border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span class="TempletItem" onClick="ClickBtn(this,'SysUserGroup')">��Ա��</span></td>
        </tr>
      </table></td>
  </tr>
  <%
end if
if JudgePopedomTF(Session("Name"),"P040400") then
%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="KeyWords" style="display:;"> 
    <td><table border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span class="TempletItem" onClick="ClickBtn(this,'SysUserList')">��Ա</span></td>
        </tr>
      </table></td>
  </tr>
  <%
end if
if JudgePopedomTF(Session("Name"),"P040501") then
%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="Source" style="display:;"> 
    <td><table border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td nowrap><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span class="TempletItem" onClick="ClickBtn(this,'SysParameter')">����ϵͳ����</span></td>
        </tr>
      </table></td>
  </tr>
  <%
end if
if JudgePopedomTF(Session("Name"),"P040502") then 
%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="Source" style="display:;"> 
    <td><table border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span class="TempletItem" onClick="ClickBtn(this,'DownLoadParameter')">����ϵͳ����</span></td>
        </tr>
      </table></td>
  </tr>
  <%
end if
if JudgePopedomTF(Session("Name"),"P040503") then 
	If HaveValueTF = True then
%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="Source" style="display:;">
    <td><table width="100" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="20"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td width="107"><span class="TempletItem" onclick="ClickBtn(this,'Mall_Config');" Type="Mall_Config">�̳ǲ�������</span></td>
        </tr>
      </table></td>
  </tr>
  <%End If%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="Source" style="display:;"> 
    <td><table border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span class="TempletItem" onClick="ClickBtn(this,'SysConstSet')">վ�㳣������</span></td>
        </tr>
      </table></td>
  </tr>
  <%
end if
if JudgePopedomTF(Session("Name"),"P070100") then 
%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="Source" style="display:;"> 
    <td><table border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span class="TempletItem" onClick="ClickBtn(this,'Folder')">����վ����</span></td>
        </tr>
      </table></td>
  </tr>
  <%
end if
if JudgePopedomTF(Session("Name"),"P070200") then 
%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="Source" style="display:;"> 
    <td><table border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span class="TempletItem" onClick="ClickBtn(this,'AdsList')">������</span></td>
        </tr>
      </table></td>
  </tr>
  <%
end if
if JudgePopedomTF(Session("Name"),"P070300") then 
%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="Source" style="display:;"> 
    <td><table border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span class="TempletItem" onClick="ClickBtn(this,'VoteList')">ͶƱ����</span></td>
        </tr>
      </table></td>
  </tr>
  <%
end if
if JudgePopedomTF(Session("Name"),"P070400") then 
%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="Source" style="display:;"> 
    <td><table border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span class="TempletItem" onClick="ClickBtn(this,'OrdinaryList1')">�ؼ��ֹ���</span></td>
        </tr>
      </table></td>
  </tr>
  <%
end if
if JudgePopedomTF(Session("Name"),"P070400") then 
%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="Source" style="display:;"> 
    <td><table height="19" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span class="TempletItem" onClick="ClickBtn(this,'OrdinaryList2')">��Դ����</span></td>
        </tr>
      </table></td>
  </tr>
  <%
end if
if JudgePopedomTF(Session("Name"),"P070400") then 
%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="Source" style="display:;"> 
    <td><table border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span class="TempletItem" onClick="ClickBtn(this,'OrdinaryList3')">���߹���</span></td>
        </tr>
      </table></td>
  </tr>
  <%
end if
if JudgePopedomTF(Session("Name"),"P070400") then 
%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="Source" style="display:;"> 
    <td><table border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span class="TempletItem" onClick="ClickBtn(this,'OrdinaryList4')">���α༭����</span></td>
        </tr>
      </table></td>
  </tr>
  <%
end if
if JudgePopedomTF(Session("Name"),"P070400") then 
%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="Source" style="display:;"> 
    <td><table border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span class="TempletItem" onClick="ClickBtn(this,'OrdinaryList5')">�ڲ����ӹ���</span></td>
        </tr>
      </table></td>
  </tr>
  <%
end if
if JudgePopedomTF(Session("Name"),"P070500") then 
%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="Source" style="display:;"> 
    <td><table border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span class="TempletItem" onClick="ClickBtn(this,'OrdinaryFriendLink')">�������ӹ���</span></td>
        </tr>
      </table></td>
  </tr>
  <%
end if
if JudgePopedomTF(Session("Name"),"P0406001") then 
%>
  <tr allparentid="0" parentid="0" classid="UserManage"> 
    <td><table border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span class="TempletItem" onClick="ClickBtn(this,'DataBase_Statistic')">����ͳ��</span></td>
        </tr>
      </table></td>
  </tr>
  <%
end if
if JudgePopedomTF(Session("Name"),"P040602") then 
%>
  <tr allparentid="UserManage" parentid="UserManage" classid="AdminGroup" style="display:;"> 
    <td><table border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span class="TempletItem" onClick="ClickBtn(this,'DataBase_Space')">�ռ�ռ��</span></td>
        </tr>
      </table></td>
  </tr>
  <%
end if
if (JudgePopedomTF(Session("Name"),"P040603")) OR (JudgePopedomTF(Session("Name"),"P040604")) then 
%>
  <tr allparentid="UserManage" parentid="UserManage" classid="Admin" style="display:;"> 
    <td><table border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span class="TempletItem" onClick="ClickBtn(this,'DataBase_Operate')">���ݿⱸ��/ѹ��</span></td>
        </tr>
      </table></td>
  </tr>
  <%
end if
if JudgePopedomTF(Session("Name"),"P040605") then 
%>
  <tr allparentid="UserManage" parentid="UserManage" classid="UserGroup" style="display:;"> 
    <td><table border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span class="TempletItem" onClick="ClickBtn(this,'DataBase_ExeCuteSql')">ִ��SQL�ű�</span></td>
        </tr>
      </table></td>
  </tr>
  <%
end if
if JudgePopedomTF(Session("Name"),"P040607") then 
%>
  <tr allparentid="UserManage" parentid="UserManage" classid="Users" style="display:;"> 
    <td><table border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span class="TempletItem" onClick="ClickBtn(this,'DataBase_LogManage')">��̨��־����</span></td>
        </tr>
      </table></td>
  </tr>
  <%
end if
%>
</table>
</body>
</html>
<script language="javascript">
var SelectedClassObj=null;
function ClickBtn(Obj,TypeStr)
{
	if (Obj!=SelectedClassObj)
	{
		Obj.className='TempletSelectItem';
		if (SelectedClassObj!=null) SelectedClassObj.className='TempletItem';
		SelectedClassObj=Obj;
	}
	top.GetEkMainObject().location=GetLocation(TypeStr,Obj);
}
function GetLocation(TypeStr,Obj)
{
	var LocationStr='';
	if (TypeStr.slice(0,6)=='FreeJS')
	{
		LocationStr='JS/FreeJSFileList.asp?JSID='+TypeStr.slice(6);
		return LocationStr;
	}
	switch (TypeStr)
	{
		case 'SysAdminGroup':
			LocationStr='System/SysAdminGroup.asp';
			break;
		case 'SysAdminList':
			LocationStr='System/SysAdminList.asp';
			break;
		case 'SysUserGroup':
			LocationStr='System/SysUserGroup.asp';
			break;
		case 'SysUserList':
			LocationStr='System/SysUserList.asp';
			break;
		case 'SysParameter':
			LocationStr='System/SysParameter.asp';
			break;
		case 'DownLoadParameter':
			LocationStr='System/DownLoadParameter.asp';
			break;
		case 'Mall_Config':
			LocationStr='Mall/Mall_Config.asp';
			break;
		case 'SysConstSet':
			LocationStr='System/SysConstSet.asp';
			break;
		case 'Folder':
			LocationStr='Recycle/Folder.asp';
			break;
		case 'AdsList':
			LocationStr='Ads/AdsList.asp';
			break;
		case 'VoteList':
			LocationStr='Vote/VoteList.asp';
			break;
		case 'OrdinaryList1':
			LocationStr='Info/OrdinaryList.asp?Type=1';
			break;
		case 'OrdinaryList2':
			LocationStr='Info/OrdinaryList.asp?Type=2';
			break;
		case 'OrdinaryList3':
			LocationStr='Info/OrdinaryList.asp?Type=3';
			break;
		case 'OrdinaryList4':
			LocationStr='Info/OrdinaryList.asp?Type=4';
			break;
		case 'OrdinaryList5':
			LocationStr='Info/OrdinaryList.asp?Type=5';
			break;
		case 'OrdinaryFriendLink':
			LocationStr='Info/OrdinaryFriendLink.asp';
			break;
		case 'DataBase_Statistic':
			LocationStr='System/DataBase_Statistic.asp';
			break;
		case 'DataBase_Space':
			LocationStr='System/DataBase_Space.asp';
			break;
		case 'DataBase_Operate':
			LocationStr='System/DataBase_Operate.asp';
			break;
		case 'DataBase_ExeCuteSql':
			LocationStr='System/DataBase_ExeCuteSql.asp';
			break;
		case 'DataBase_LogManage':
			LocationStr='System/DataBase_LogManage.asp';
			break;
	}
	return LocationStr;
}
</script>
<%
Set Conn = Nothing
%>
