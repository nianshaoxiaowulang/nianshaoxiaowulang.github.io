<% Option Explicit %>
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Const.asp"-->
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
if Not JudgePopedomTF(Session("Name"),"P040602") then Call ReturnError1()
Dim fsoSpaceObj,sRootDir
If SysRootDir<>"" then
	sRootDir="/"& SysRootDir
Else
	sRootDir=""
End If
set fsoSpaceObj=Server.CreateObject(G_FS_FSO)

Dim SysPath,SysSpace,ShowSysSpace,GetSysSpace,SysPicSize
if SysRootDir = "" then
	SysPath=Server.mappath("/")
else
	SysPath=Server.mappath("/" & SysRootDir)
end if
if fsoSpaceObj.FolderExists(SysPath) then
	set GetSysSpace=fsoSpaceObj.GetFolder(SysPath)
	SysSpace=GetSysSpace.size
	if SysSpace=0  then
		ShowSysSpace=0
	else	
		SysSpace=SysSpace/1024/1024
		ShowSysSpace = formatnumber(SysSpace,6,-1)
	end if
else
	ShowSysSpace=0
end if
SysPicSize=ShowSysSpace*100
	
Dim AdminPath,AdminSpace,ShowAdminSpace,GetAdminSpace,AdminPicSize
AdminPath=Server.mappath(sRootDir&"/"&AdminDir)
if fsoSpaceObj.FolderExists(AdminPath) then
	set GetAdminSpace=fsoSpaceObj.GetFolder(AdminPath)
	AdminSpace=GetAdminSpace.size
	if AdminSpace=0 then
		ShowAdminSpace=0
	else
		AdminSpace=AdminSpace/1024/1024
		ShowAdminSpace = formatnumber(AdminSpace,6,-1)
	end if
else
    ShowAdminSpace=0
end if
AdminPicSize=ShowAdminSpace*100

Dim NewsPath,NewsSpace,ShowNewsSpace,GetNewsSpace,NewsPicSize
NewsPath=Server.mappath(sRootDir&"/"&ClassDir)
if fsoSpaceObj.FolderExists(NewsPath) then
	set GetNewsSpace=fsoSpaceObj.GetFolder(NewsPath)
	NewsSpace=GetNewsSpace.size
	if NewsSpace=0 then
		ShowNewsSpace=0
	else
		NewsSpace=NewsSpace/1024/1024
		ShowNewsSpace = formatnumber(NewsSpace,6,-1)
	end if
else
	ShowNewsSpace=0
end if
NewsPicSize=ShowNewsSpace*100

Dim ClassPath,ClassSpace,ShowClassSpace,GetClassSpace,ClassPicSize
ClassPath=Server.mappath(sRootDir&"/"&ClassDir)
if fsoSpaceObj.FolderExists(ClassPath) then
	set GetClassSpace=fsoSpaceObj.GetFolder(ClassPath)
	ClassSpace=GetClassSpace.size
	if ClassSpace=0 then
		ShowClassSpace=0
	else
		ClassSpace=ClassSpace/1024/1024
		ShowClassSpace = formatnumber(ClassSpace,6,-1)
	end if
else
	ShowClassSpace=0
end if
ClassPicSize=ShowClassSpace*100

Dim SpecialPath,SpecialSpace,ShowSpecialSpace,GetSpecialSpace,SpecialPicSize
SpecialPath=Server.mappath(sRootDir&"/"&ClassDir)
if fsoSpaceObj.FolderExists(SpecialPath) then
	set GetSpecialSpace=fsoSpaceObj.GetFolder(SpecialPath)
	SpecialSpace=GetSpecialSpace.size
	if SpecialSpace=0 then
		ShowSpecialSpace=0
	else
		SpecialSpace=SpecialSpace/1024/1024
		ShowSpecialSpace = formatnumber(SpecialSpace,6,-1)
	end if
else
	ShowSpecialSpace=0
end if
SpecialPicSize=ShowSpecialSpace*100

Dim UpfilesPath,UpfilesSpace,ShowUpfilesSpace,GetUpfilesSpace,UpfilesPicSize
UpfilesPath=Server.mappath(sRootDir&"/"&UpFiles)
if fsoSpaceObj.FolderExists(UpfilesPath) then
	set GetUpfilesSpace=fsoSpaceObj.GetFolder(UpfilesPath)
	UpfilesSpace=GetUpfilesSpace.size
	if UpfilesSpace=0 then
		ShowUpfilesSpace=0
	else
		UpfilesSpace=UpfilesSpace/1024/1024
		ShowUpfilesSpace = formatnumber(UpfilesSpace,6,-1)
	end if
else
	ShowUpfilesSpace=0
end if
UpfilesPicSize=ShowUpfilesSpace*100

Dim TempletPath,TempletSpace,ShowTempletSpace,GetTempletSpace,TempletPicSize
TempletPath=Server.mappath(sRootDir&"/"&TempletDir)
if fsoSpaceObj.FolderExists(TempletPath) then
	set GetTempletSpace=fsoSpaceObj.GetFolder(TempletPath)
	TempletSpace=GetTempletSpace.size
	if TempletSpace=0 then
		ShowTempletSpace=0
	else
		TempletSpace=TempletSpace/1024/1024
		ShowTempletSpace = formatnumber(TempletSpace,6,-1)
	end if
else
	ShowTempletSpace=0
end if
TempletPicSize=ShowTempletSpace*100
%>
<html>
<head>
<title>���ݿռ�ռ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script src="../JS/PublicJS.js" language="JavaScript"></script>
<body bgcolor="#FFFFFF" topmargin="2" leftmargin="2" oncontextmenu="return false;">
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="28" class="ButtonListLeft">
<div align="center"><strong>ϵͳ�ռ�ֲ����</strong></div></td>
  </tr>
</table>
<table border="0" align="center" cellpadding="0" cellspacing="0">
 <tr><td height="50">&nbsp;</td>
 </tr>
  <tr><td>
<table width="600" height="256" border="1" align="center" cellpadding="3" cellspacing="3" bordercolor="e6e6e6" >
<%if ShowSysSpace<>0 then %>
  <tr> 
    <td height="16" >ϵͳռ�ÿռ䣺<img src="../../Images/Visit/count.gif" width="<% =SysPicSize/300 %>" height=10><%=ShowSysSpace%>&nbsp;MB</td>
  </tr>
<% else %>
  <tr> 
    <td height="16" >ϵͳռ�ÿռ䣺<%=ShowSysSpace%>&nbsp;MB</td>
  </tr>
<% end if
if ShowAdminSpace<>0 then %>
  <tr> 
      <td height="16" >����ռ�ÿռ䣺<img src="../../Images/Visit/count.gif" width="<% =AdminPicsize/300 %>" height=10><%=ShowAdminSpace%>&nbsp;MB</td>
  </tr>
<% else %>
  <tr> 
      <td height="16" >����ռ�ÿռ䣺<%=ShowAdminSpace%>&nbsp;MB</td>
  </tr>
<% end if
if ShowNewsSpace<>0 then %>
  <tr> 
    <td  height="16" > ����ռ�ÿռ䣺<img src="../../Images/Visit/count.gif" width="<%=NewsPicSize/300 %>" height=10><%=ShowNewsSpace%>&nbsp;MB</td>
  </tr>
<%else%>
  <tr> 
    <td  height="16" > ����ռ�ÿռ䣺<%=ShowNewsSpace%>&nbsp;MB</td>
  </tr>
<%end if
if ShowClassSpace<>0 then %>
  <tr> 
    <td  height="16"  > ��Ŀռ�ÿռ䣺<img src="../../Images/Visit/count.gif" width="<%=ClassPicSize/300 %>" height=10><%=ShowClassSpace%>&nbsp;MB</td>
  </tr>
<%else%>
  <tr> 
    <td  height="16" > ��Ŀռ�ÿռ䣺<%=ShowClassSpace%>&nbsp;MB</td>
  </tr>
<%end if
if ShowSpecialSpace<>0 then %>
  <tr> 
    <td  height="16"  > ר��ռ�ÿռ䣺<img src="../../Images/Visit/count.gif" width="<%=SpecialPicSize/300 %>" height=10><%=ShowSpecialSpace%>&nbsp;MB</td>
  </tr>
<%else%>
  <tr> 
    <td  height="16" > ר��ռ�ÿռ䣺<%=ShowSpecialSpace%>&nbsp;MB</td>
  </tr>
<%end if
if ShowUpfilesSpace<>0 then %>
  <tr> 
    <td  height="16" > �ϴ�ռ�ÿռ䣺<img src="../../Images/Visit/count.gif" width="<%=UpfilesPicSize/300 %>" height=10><%=ShowUpfilesSpace%>&nbsp;MB</td>
  </tr>
<%else%>
  <tr> 
    <td  height="16" > �ϴ�ռ�ÿռ䣺<%=ShowUpfilesSpace%>&nbsp;MB</td>
  </tr>
<%end if
if ShowTempletSpace<>0 then %>
  <tr> 
    <td  height="16" > ģ��ռ�ÿռ䣺<img src="../../Images/Visit/count.gif" width="<%=TempletPicSize/300 %>" height=10><%=ShowTempletSpace%>&nbsp;MB</td>
  </tr>
<%else%>
  <tr> 
    <td  height="16" > ģ��ռ�ÿռ䣺<%=ShowTempletSpace%>&nbsp;MB</td>
  </tr>
<%end if%>
</table>
</td></tr></table>
</body>
</html>