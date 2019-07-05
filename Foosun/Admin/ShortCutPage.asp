<% Option Explicit %>
<!--#include file="../../Inc/Const.asp" -->
<!--#include file="../../Inc/Cls_DB.asp" -->
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
<title>快捷菜单</title>
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
          <td><span class="TempletItem" onClick="ClickBtn(this,'SysAdminGroup')">管理员组</span></td>
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
          <td><span class="TempletItem" onClick="ClickBtn(this,'SysAdminList')">管理员</span></td>
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
          <td><span class="TempletItem" onClick="ClickBtn(this,'SysUserGroup')">会员组</span></td>
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
          <td><span class="TempletItem" onClick="ClickBtn(this,'SysUserList')">会员</span></td>
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
          <td><span class="TempletItem" onClick="ClickBtn(this,'SysParameter')">新闻系统参数</span></td>
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
          <td><span class="TempletItem" onClick="ClickBtn(this,'DownLoadParameter')">下载系统参数</span></td>
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
          <td width="107"><span class="TempletItem" onclick="ClickBtn(this,'Mall_Config');" Type="Mall_Config">商城参数设置</span></td>
        </tr>
      </table></td>
  </tr>
  <%End If%>
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="Source" style="display:;"> 
    <td><table border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span class="TempletItem" onClick="ClickBtn(this,'SysConstSet')">站点常量设置</span></td>
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
          <td><span class="TempletItem" onClick="ClickBtn(this,'Folder')">回收站管理</span></td>
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
          <td><span class="TempletItem" onClick="ClickBtn(this,'AdsList')">广告管理</span></td>
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
          <td><span class="TempletItem" onClick="ClickBtn(this,'VoteList')">投票管理</span></td>
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
          <td><span class="TempletItem" onClick="ClickBtn(this,'OrdinaryList1')">关键字管理</span></td>
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
          <td><span class="TempletItem" onClick="ClickBtn(this,'OrdinaryList2')">来源管理</span></td>
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
          <td><span class="TempletItem" onClick="ClickBtn(this,'OrdinaryList3')">作者管理</span></td>
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
          <td><span class="TempletItem" onClick="ClickBtn(this,'OrdinaryList4')">责任编辑管理</span></td>
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
          <td><span class="TempletItem" onClick="ClickBtn(this,'OrdinaryList5')">内部链接管理</span></td>
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
          <td><span class="TempletItem" onClick="ClickBtn(this,'OrdinaryFriendLink')">友情链接管理</span></td>
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
          <td><span class="TempletItem" onClick="ClickBtn(this,'DataBase_Statistic')">数据统计</span></td>
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
          <td><span class="TempletItem" onClick="ClickBtn(this,'DataBase_Space')">空间占用</span></td>
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
          <td><span class="TempletItem" onClick="ClickBtn(this,'DataBase_Operate')">数据库备份/压缩</span></td>
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
          <td><span class="TempletItem" onClick="ClickBtn(this,'DataBase_ExeCuteSql')">执行SQL脚本</span></td>
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
          <td><span class="TempletItem" onClick="ClickBtn(this,'DataBase_LogManage')">后台日志管理</span></td>
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
