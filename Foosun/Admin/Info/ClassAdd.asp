<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../Inc/Cls_Info.asp" -->
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
Dim DBC,Conn,sRootDir
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
if SysRootDir<>"" then sRootDir="/"+SysRootDir else sRootDir=""
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"" & Request("ParentID") & "") then Call ReturnError1()
if Not JudgePopedomTF(Session("Name"),"P010100") then Call ReturnError1()
Dim ParentID,Result,ClassID,ParentCName,RsParentObj,SaveFilePath,DoMain,SelectPathBtnDisabledStr,TempParentID,SelectPath
Dim CheckRootClassNumber  '������Ŀ��������ֹ��ѭ��
Dim RsMenuConfigObj,HaveValueTF
Set RsMenuConfigObj = Conn.execute("Select IsShop From FS_Config")
if RsMenuConfigObj("IsShop") = 1 then
	HaveValueTF = True
Else
	HaveValueTF = False
End if
Set RsMenuConfigObj = Nothing
Result = Request("Result")
if Result = "InClass" then
	Dim CClass,ReturnCheckInfo,ReturnValueArray
	Set CClass = New InfoClass
	CClass.TForm = Request.Form
	ReturnCheckInfo = CClass.AddAndModifyClass()
	ReturnValueArray = Split(ReturnCheckInfo,"||")
	if ReturnValueArray(0) = "Success" then
		%>
		<script>
			top.GetNavFoldersObject().location='../Menu_Folders.asp?Action=ContentTree&OpenClassIDList=<% = ReturnValueArray(1) %>';		
		</script>
		<%
	else
		%>
		<script>alert('<% = ReturnCheckInfo %>');history.back();</script>
		<%
	end if
ElseIf Result="OutClass" then
	Dim RsAddClass,StrClassID
	on error resume next
	Set RsAddClass=Server.CreateObject(G_FS_RS)
	RsAddClass.open "Select * from FS_NewsClass where 1=2",Conn,3,3
	RsAddClass.addNew
	StrClassID=GetRandomID18()
	RsAddClass("ClassID") = StrClassID
	RsAddClass("ClassCName")=NoCSSHackAdmin(request.Form("ClassName"),"��Ŀ����")
	RsAddClass("ClassLink")=NoCSSHackAdmin(Request.Form("ClassLink"),"���ӵ�ַ")
	RsAddClass("ClassEName")="OutClass"&StrClassID
	RsAddClass("ParentID")=0
	RsAddClass("ChildNum")=0
	RsAddClass("ClassTemp")="ClassTemplet"
	RsAddClass("Contribution")=0
	RsAddClass("DelFlag")=0
	RsAddClass("FileTime")=100
	RsAddClass("BrowPop") = 0
	RsAddClass("RedirectList") = Cint(Request("RedirectList"))
	if Request.Form("Orders") <> "" then
		if IsNumeric(Request.Form("Orders")) then RsAddClass("Orders") = Request.Form("Orders")
	end if
	if Request.Form("ShowTF") = "1" then
		RsAddClass("ShowTF") = 1
	else
		RsAddClass("ShowTF") = 0
	end if
	RsAddClass("AddTime") = Now
	RsAddClass("IsOutClass")=1
	RsAddClass.UpDate
	RsAddClass.Close
	Set RsAddClass=Nothing
	If err=0 then 
		%>
		<script>
			top.GetNavFoldersObject().location='../Menu_Folders.asp?Action=ContentTree&OpenClassIDList=<% = StrClassID %>';		
		</script>
		<%
	else
		%>
		<script>alert('<% = err.description %>');history.back();</script>
		<%
	end if
end if

CheckRootClassNumber = 30
SelectPathBtnDisabledStr = ""
ClassID = Request("ClassID")
ParentID = Request("ParentID")
SaveFilePath = Request("SaveFilePath")
if ParentID = "" Or ParentID = "0" then
	ParentID = "0"
	ParentCName = "ϵͳ����Ŀ"
	SelectPath = "/" & RemoveVirtualPath(ClassDir)
else
	Set RsParentObj = Conn.Execute("Select ClassCName,ParentID,DoMain,SaveFilePath from FS_NewsClass where ClassID='" & ParentID & "'")
	if RsParentObj.Eof then
		Set RsParentObj = Nothing
		Set Conn = Nothing
		Alert "����Ŀ�����ڣ������¼���"
		Response.End
	else
		Dim CheckRootClassIndex
		CheckRootClassIndex = 1
		ParentCName = RsParentObj("ClassCName")
		TempParentID = RsParentObj("ParentID")
		do while Not (TempParentID = "0")
			CheckRootClassIndex = CheckRootClassIndex + 1
			RsParentObj.Close
			Set RsParentObj = Nothing
			Set RsParentObj = Conn.Execute("Select ClassCName,ParentID,Domain,SaveFilePath from FS_NewsClass where ClassID='" & TempParentID & "'")
			if RsParentObj.Eof then
				Set RsParentObj = Nothing
				Alert "����Ŀ������"
				Response.End
			end if
			TempParentID = RsParentObj("ParentID")
			if CheckRootClassIndex > CheckRootClassNumber then TempParentID = "0" '��ֹ��ѭ��
		Loop
		DoMain = RsParentObj("DoMain")
		if (Not IsNull(DoMain)) And (DoMain <> "") then
			SelectPath = RsParentObj("SaveFilePath")
			SelectPathBtnDisabledStr = " disabled"
		else
			SelectPath = "/" & RemoveVirtualPath(ClassDir)
		end if
	end if
	Set RsParentObj = Nothing
end if
Dim DoMainDisabledStr
if ParentID <> "0" then
	DoMainDisabledStr = " disabled"
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��Ŀ���</title>
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script language="JavaScript" src="../../SysJS/PublicJS.js"></script>
<body scroll=no topmargin="2" leftmargin="2">
<div id="TempShowMenu" style="display=''">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
            <td width="35" align="center" alt="����" onClick="OK();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  	<td width=2 class="Gray">|</td>
		  	<td width=35 align="center" alt="����" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
            <td>&nbsp; <input name="Result" type="hidden" id="Result" value="InClass"> 
              <input name="ClassID" value="<% = ClassID %>" type="hidden" id="ClassID"> 
              <input name="ParentID" value="<% = ParentID %>" type="hidden" id="ParentID"></td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
</div>
<div id="InClass" style="display:none;">
<form action="" method="post" name="InClassForm">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999" >
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
            <td width="35" align="center" alt="����" onClick="InOK();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  	<td width=2 class="Gray">|</td>
		  	<td width=35 align="center" alt="����" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
            <td>&nbsp; <input name="Result" type="hidden" id="Result" value="InClass"> 
              <input name="ClassID" value="<% = ClassID %>" type="hidden" id="ClassID"> 
              <input name="ParentID" value="<% = ParentID %>" type="hidden" id="ParentID"></td>
        </tr>
      </table>
	  </td>
  </tr>
</table>

    <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#E3E3E3">
      <tr bgcolor="#FFFFFF"> 
        <td width="100" height="26"> <div align="center">��������</div></td>
        <td> <div align="center"> 
            <input value="<% = Request("ClassCName") %>" name="ClassCName" type="text" id="ClassCName" style="width:100%;">
          </div></td>
      </tr>
      <tr bgcolor="#FFFFFF"> 
        <td height="26"> <div align="center">Ӣ������</div></td>
        <td> <div align="center"> 
            <input value="<% = Request("ClassEName") %>" name="ClassEName" type="text" id="ClassEName" style="width:100%;">
          </div></td>
      </tr>
      <tr bgcolor="#FFFFFF"> 
        <td height="26" align="center"> �� �� Ŀ</td>
        <td> <input readonly value="<% = ParentCName %>" style="width:100%;" type="text" name="textfield3"> 
        </td>
      </tr>
      <tr bgcolor="#FFFFFF"> 
        <td height="26" align="center"> ��������</td>
        <td nowrap> <input <% = DoMainDisabledStr %> type="text" name="DoMain" style="width:100%;" value="<% = DoMain %>"></td>
      </tr>
      <tr bgcolor="#FFFFFF"> 
        <td height="26" align="center"> ���Ȩ��</td>
        <td nowrap> <select name="BrowPop" id="BrowPop" style="width:100%" onChange="CheckFileExtName(this);">
            <option value="" <%if Request("BrowPop")="" then Response.Write("selected")%>> 
            </option>
            <%
		Dim BrowPopObj
		set BrowPopObj = Conn.Execute("Select Name,PopLevel from FS_MemGroup order by PopLevel asc")
		while not BrowPopObj.eof
		%>
            <option value="<%=BrowPopObj("PopLevel")%>" <%if Request("BrowPop")<>"" and isnull(Request("BrowPop"))=false then if Cint(Request("BrowPop")) = Cint(BrowPopObj("PopLevel")) then Response.Write("selected") end if end if%>><%=BrowPopObj("Name")%></option>
            <%
			BrowPopObj.Movenext
		Wend
		BrowPopObj.Close
		Set BrowPopObj = Nothing
		%>
          </select></td>
      </tr>
      <tr bgcolor="#FFFFFF"> 
        <td height="26" align="center">�� չ ��</td>
        <td align="left" nowrap> <select name="FileExtName" style="width:100%;">
            <option value="htm">htm</option>
            <option value="html" selected>html</option>
            <option value="shtml">shtml</option>
            <option value="asp">asp</option>
          </select> </td>
      </tr>
      <tr bgcolor="#FFFFFF"> 
        <td height="26" align="center">��Ŀģ��</td>
        <td align="left" nowrap> <input name="ClassTemp" type="text" style="width:78%;" value="/<% = TempletDir %>/NewsClass/class.htm" readonly> 
          <input type="button" name="Submit" value="ѡ��ģ��" onClick="OpenWindowAndSetValue('../../FunPages/SelectFileFrame.asp?CurrPath=<%=sRootDir %>/<% = TempletDir %>',400,300,window,document.InClassForm.ClassTemp);document.InClassForm.ClassTemp.focus();"> 
        </td>
      </tr>
      <tr bgcolor="#FFFFFF"> 
        <td height="26" align="center">��������ģ��</td>
        <td align="left" nowrap> <input type="text" style="width:78%;" name="NewsTemp" value="/<% = TempletDir %>/NewsClass/news.htm" readonly> 
          <input type="button" name="Submit2" value="ѡ��ģ��" onClick="OpenWindowAndSetValue('../../FunPages/SelectFileFrame.asp?CurrPath=<%=sRootDir %>/<% = TempletDir %>',400,300,window,document.InClassForm.NewsTemp);document.InClassForm.NewsTemp.focus();"></td>
      </tr>
      <tr bgcolor="#FFFFFF"> 
        <td height="26" align="center"> ��������ģ��</td>
        <td align="left" nowrap> <input type="text" style="width:78%;" name="DownLoadTemp" value="/<% = TempletDir %>/NewsClass/DownLoad.htm" readonly> 
          <input type="button" name="Submit2" value="ѡ��ģ��" onClick="OpenWindowAndSetValue('../../FunPages/SelectFileFrame.asp?CurrPath=<%=sRootDir %>/<% = TempletDir %>',400,300,window,document.InClassForm.DownLoadTemp);document.InClassForm.DownLoadTemp.focus();"></td>
      </tr>
      <tr bgcolor="#FFFFFF" <%if HaveValueTF = False then response.Write("style=""display:none""")%>> 
        <td height="26" align="center"> ������Ʒģ��</td>
        <td align="left" nowrap> <input type="text" style="width:78%;" name="ProductTemp" value="/<% = TempletDir %>/Mall/Product.htm" readonly> 
          <input type="button" name="Submit2" value="ѡ��ģ��" onClick="OpenWindowAndSetValue('../../FunPages/SelectFileFrame.asp?CurrPath=<%=sRootDir %>/<% = TempletDir %>',400,300,window,document.InClassForm.ProductTemp);document.InClassForm.ProductTemp.focus();"></td>
      </tr>
      <tr bgcolor="#FFFFFF"> 
        <td height="26" align="center"> ����·��</td>
        <td align="left" nowrap> <input readonly value="/<% = ClassDir %>" type="text" style="width:78%;" name="SaveFilePath"> 
          <input type="button" <% =SelectPathBtnDisabledStr %> name="Submit" value="ѡ��·��" onClick="OpenWindowAndSetValue('../../FunPages/SelectPathFrame.asp?CurrPath=<%=sRootDir %>/<% = ClassDir %>',400,300,window,document.InClassForm.SaveFilePath);document.InClassForm.SaveFilePath.focus();"> 
        </td>
      </tr>
      <tr bgcolor="#FFFFFF"> 
        <td height="26"> <div align="center">�������</div></td>
        <td> <div align="left"> 
            <input readonly value="<%If Request("AddTime")="" then Response.Write(Now()) else Response.Write(Request("AddTime")) %>" type="text" style="width:78%;" name="AddTime">
            <input name="sdaf" type="button" id="sdaf" onClick="OpenWindowAndSetValue('../../FunPages/SelectDate.asp',280,120,window,document.InClassForm.AddTime);document.InClassForm.AddTime.focus();" value="ѡ������">
          </div></td>
      </tr>
      <tr bgcolor="#FFFFFF"> 
        <td height="26"> <div align="center">���Ӳ���</div></td>
        <td> <input name="Contribution" type="checkbox" id="Contribution2" value="1" <% if Request("Contribution") = 1 then Response.Write("checked") %>>
          ����Ͷ�� &nbsp;&nbsp;&nbsp;&nbsp; <input name="ShowTF" type="checkbox" id="ShowTF2" value="1" checked <% if Request("ShowTF") = 1 then Response.Write("checked") %>>
          ǰ̨��ʾ &nbsp;&nbsp;&nbsp;&nbsp;��Ŀ���� 
          <input name="Orders" type="text" size="5" maxlength="4">
          ���Ź鵵ʱ�䣨���֣� 
          <input name="FileTime" type="text" value="<% if Request("FileTime") = "" then Response.Write("100") else Response.Write(Request("FileTime"))%>" size="5" maxlength="3">
          �� ��
</td>
      <tr bgcolor="#FFFFFF">
        <td height="26"><div align="center">Ĭ��ת��</div></td>
        <td><select name="RedirectList" id="RedirectList" style="width:100%;">
            <option value="1" selected>�����б�</option>
            <option value="2">�����б�</option>
            <%if HaveValueTF = True then%>
            <option value="3">��Ʒ�б�</option>
            <%End if%>
          </select></td>
    </table>
</form>
</div>

<div  dwcopytype="CopyTableRow" id="OutClass" style="display:none;">

<form action="" method="post" name="OutClassForm">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
        	  <td width="35" align="center" alt="����" onClick="OutOK();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
			  <td width=2 class="Gray">|</td>
			  <td width=35 align="center" alt="����" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
            <td>&nbsp; <input name="Result" type="hidden" id="Result" value="OutClass"> 
            </td>
        </tr>
      </table>
	  </td>
  </tr>
</table>

    <table width="100%"  border="0" cellpadding="3" cellspacing="1" bgcolor="#E3E3E3">
      <tr bgcolor="#FFFFFF"> 
        <td width="110" height="26"> 
          <div align="center">��Ŀ����</div></td>
        <td> 
          <input name="ClassName" type="text" id="Name" style="width:100%" value=""></td>
    </tr>
      <tr bgcolor="#FFFFFF"> 
        <td height="26"> 
          <div align="center">���ӵ�ַ</div></td>
        <td> 
          <input name="ClassLink" type="text" id="Link" style="width:100%" value=""></td>
    </tr>
      <tr bgcolor="#FFFFFF"> 
        <td height="26"> 
          <div align="center">���Ӳ���</div></td>
        <td> 
          <input name="ShowTF" type="checkbox" id="ShowTF2" value="1" checked <% if Request("ShowTF") = 1 then Response.Write("checked") %>>
        ǰ̨��ʾ &nbsp;&nbsp;&nbsp;&nbsp;��Ŀ���� 
        <input name="Orders" type="text" size="5" maxlength="4">
      </td>
    </tr>
</table>
</form>
</div>
<table width="100%" height="26" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td colspan="5" height="2"></td>
	</tr>
	<tr bgcolor="#EEEEEE">
		
    <td height="26"> <table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="#999999">
        <tr>
          <td bgcolor="#EFEFEF">
<table width="100%" height="24" border="0" cellpadding="5" cellspacing="1">
        <tr>
				<td width="14%"  align="left" alt="�����ͨ��Ŀ" onClick="InClass.style.display='';OutClass.style.display='none';TempShowMenu.style.display='none';IsOutImg.src='../images/r.gif';IsInImg.src='../images/u.gif';" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut"><div align="center">�����ͨ��Ŀ</div></td>
			  	<td width="86%"><div align="right"><img id="IsInImg" src="../images/r.gif" width="20" height="20"></div></td>
			  </tr>
			</table></td>
        </tr>
      </table> </td>
	</tr>
</table>
<table width="100%" height="26" border="0" cellpadding="0" cellspacing="0">
	<tr>
	<td colspan="5" height="2"></td>
	</tr>
	<tr bgcolor="#EEEEEE">
		
    <td height="26"> <table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="#999999">
        <tr>
          <td bgcolor="#EFEFEF">
<table width="100%" height="24" border="0" cellpadding="5" cellspacing="1">
              <tr> 
                <td width="14%" align="left" alt="����ⲿ��Ŀ" onClick="InClass.style.display='none';OutClass.style.display='';TempShowMenu.style.display='none';IsOutImg.src='../images/u.gif';IsInImg.src='../images/r.gif';" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut"><div align="center">����ⲿ��Ŀ</div></td>
                <td width="86%"><div align="right"><img id="IsOutImg" src="../images/r.gif" width="20" height="20"></div></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
	</tr>
</table>

</body>
</html>
<%
Set Conn = Nothing
Sub Alert(InfoStr)
	%>
	<script language="JavaScript">
	alert('<% = InfoStr %>');
	history.back();
	window.close();
	</script>
	<%
End Sub
%>
<script language="JavaScript">
function InOK()
{
	if (document.InClassForm.ClassCName.value=='') {alert('��д��Ŀ��������');document.InClassForm.ClassCName.focus();return;}
	if (document.InClassForm.ClassEName.value=='') {alert('��д��ĿӢ������');document.InClassForm.ClassEName.focus();return;}
	if (CheckEnglishStr(document.InClassForm.ClassEName,'Ӣ������')==true)
	{
		document.InClassForm.submit();
	}
}
function OutOK()
{
	if (document.OutClassForm.ClassName.value=='') {alert('��д��Ŀ��������');document.OutClassForm.ClassName.focus();return;}
	if (document.OutClassForm.ClassLink.value=='') {alert('��д��Ŀ���ӵ�ַ');document.OutClassForm.ClassLink.focus();return;}
	document.OutClassForm.submit();
}
function OK()
{
	alert('����ѡ����Ҫ��ӵ���Ŀ������');
	return;
}
function CheckFileExtName(Obj)
{
	if (Obj.value!='')
	{
		for (var i=0;i<document.all.FileExtName.length;i++)
		{
			if (document.all.FileExtName.options(i).value=='asp') document.all.FileExtName.options(i).selected=true;
		}
		document.all.FileExtName.disabled=true;
	}
	else
	{
		document.all.FileExtName.disabled=false;
	}
}
</script>