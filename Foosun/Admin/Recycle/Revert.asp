<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<!--#include file="../Inc/Cls_JS.asp" -->
<%
'==============================================================================
'������ƣ���Ѷ��վ��Ϣ����ϵͳ
'��ǰ�汾��Foosun Content Manager System v3.1 
'���¸��£�2004.12
'==============================================================================
'��ҵע����ϵ��028-85098980-601,602 ����֧�֣�028-85098980-606��607,�ͻ�֧�֣�608
'��Ʒ��ѯQQ��159410,655071,66252421
'����֧��:���г���ʹ�����⣬�����ʵ�bbs.foosun.net���ǽ���ʱ�ش���
'���򿪷�����Ѷ������ & ��Ѷ���������
'��̳֧�֣���Ѷ������̳(http://bbs.foosun.net)
'�ٷ���վ��www.Foosun.net  ��ʾվ�㣺test.cooin.com    
'��վ����ר����www.cooin.com
'==============================================================================
'��Ѱ汾����������ҳ������Ȩ��Ϣ�������ϱ�վLOGO��������
'==============================================================================
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P070101") then Call ReturnError()
Dim TempSysRootDir
if SysRootDir = "" then
	TempSysRootDir = ""
else
	TempSysRootDir = "/" & SysRootDir
end if

Dim OperateType,ID
Dim Sql,RsTempObj,PromptInfo
OperateType = Request("OperateType")
ID = Request("ID")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��Ϣ��ԭ</title>
</head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet">
<body scrolling=no>
<table width="100%" border="0" cellspacing="5" cellpadding="0">
  <tr> 
    <td width="29%" height="30" align="center"><img src="../../Images/Question.gif" width="39" height="37"> 
    </td>
    <td width="71%" align="center"><div align="left">
        <%
if OperateType = "Class" then
	Sql = "Select ClassCName,ClassID,ParentID from FS_NewsClass where ClassID='" & ID & "'"
	PromptInfo = "��Ŀ"
elseif OperateType = "News" then
	Sql = "Select Title,NewsID,ClassID from FS_News where NewsID='" & ID & "'"
	PromptInfo = "����"
else
	Sql = ""
end if
if Sql <> "" then
	Set RsTempObj = Conn.Execute(Sql)
	if Not RsTempObj.Eof then
		Dim ResponseStr
		if PromptInfo = "��Ŀ" then
			if RsTempObj("ParentID") <> "0" then
				Sql = "Select ClassID,DelFlag from FS_NewsClass Where ClassID='" & RsTempObj("ParentID") & "'"
				RsTempObj.Close
				Set RsTempObj = Conn.Execute(Sql)
				if RsTempObj.Eof then
					ResponseStr = "��Ŀ�ĸ���Ŀ�Ѿ���ɾ������ѡ��Ŀ����Ŀ<select name=""ParentID""><option value=""0"">����Ŀ</option>" & ClassList & "</select>"
				else
					if RsTempObj("DelFlag") = 1 then
						ResponseStr = "��Ŀ�ĸ���Ŀ�ڻ���վ����ѡ��Ŀ����Ŀ<select name=""ParentID""><option value=""0"">����Ŀ</option>" & ClassList & "</select>"
					else
						ResponseStr = "ȷ��Ҫ��ԭ����Ŀ��"
					end if
				end if
			else
				ResponseStr = "ȷ��Ҫ��ԭ����Ŀ��"
			end if
			Response.Write(ResponseStr)
		else
			Sql = "Select ClassID,DelFlag from FS_NewsClass where ClassID='" & RsTempObj("ClassID") & "'"
			Set RsTempObj = Conn.Execute(Sql)
			if RsTempObj.Eof then
				ResponseStr = "������Ŀ�Ѿ���ɾ������ѡ��Ŀ����Ŀ<select name=""ParentID"">" & ClassList & "</select>"
			else
				if RsTempObj("DelFlag") = 1 then
					ResponseStr = "������Ŀ�ڻ���վ����ѡ��Ŀ����Ŀ<select name=""ParentID"">" & ClassList & "</select>"
				else
					ResponseStr = "ȷ��Ҫ��ԭ��������"
				end if
			end if
			Response.Write(ResponseStr)
		end if
	else
%>
          <script language="JavaScript">
alert('��<% = PromptInfo %>�Ѿ���ɾ��');
window.close();
        </script>
        <%
	end if
else
%>
          <script language="JavaScript">
alert('�������ݴ���');
window.close();
        </script>
        <%
end if
%>
    </div></td>
  </tr>
  <tr> 
    <td height="30" colspan="2"><div align="center"> 
        <input onClick="SubmitFun();" type="button" name="Submit" value=" ȷ �� ">
        <input type="button" onClick="window.close();" name="Submit2" value=" ȡ �� ">
      </div></td>
  </tr>
</table>
</body>
</html>
<%
Dim Result
Result = Request("Result")
if Result = "Submit" then
	Dim ParentID
	ParentID = Request("ParentID")
	On Error Resume Next
	if OperateType = "Class" then
		if ParentID = "" then
			Sql = "Update FS_NewsClass Set DelFlag=0 where ClassID='" & ID & "'"
		else
			Sql = "Update FS_NewsClass Set DelFlag=0,ParentID='" & ParentID & "' where ClassID='" & ID & "'"
		end if
		Conn.Execute(Sql)
		Sql = "UpDate FS_News Set DelTF=0 where ClassID='" & ID & "'"
		Conn.Execute(Sql)
		Conn.Execute("Update FS_FreeJsFile set DelFlag=0 where ClassID='"&ID&"'")
		'-----------���������������JS�ļ�---------------------
		Dim RevertFreeJsObj,TempClassIDStrs,FreeJsObj
		TempClassIDStrs = ""
		Set RevertFreeJsObj = Conn.Execute("Select distinct JSName from FS_FreeJsFile where ClassID='" & ID & "'")
		Do while Not RevertFreeJsObj.eof
			Set FreeJsObj = Conn.Execute("Select EName,Manner from FS_FreeJS where EName='"&RevertFreeJsObj("JSName")&"'")
		    If Not FreeJsObj.eof then
			Select case FreeJsObj("Manner")
			 case "1"   WCssA FreeJsObj("EName"),True
			 case "2"   WCssB FreeJsObj("EName"),True
			 case "3"   WCssC FreeJsObj("EName"),True
			 case "4"   WCssD FreeJsObj("EName"),True
			 case "5"   WCssE FreeJsObj("EName"),True
			 case "6"   PCssA FreeJsObj("EName"),True
			 case "7"   PCssB FreeJsObj("EName"),True
			 case "8"   PCssC FreeJsObj("EName"),True
			 case "9"   PCssD FreeJsObj("EName"),True
			 case "10"   PCssE FreeJsObj("EName"),True
			 case "11"   PCssF FreeJsObj("EName"),True
			 case "12"   PCssG FreeJsObj("EName"),True
			 case "13"   PCssH FreeJsObj("EName"),True
			 case "14"   PCssI FreeJsObj("EName"),True
			 case "15"   PCssJ FreeJsObj("EName"),True
			 case "16"   PCssK FreeJsObj("EName"),True
			 case "17"   PCssL FreeJsObj("EName"),True
		   End Select
		   End If
		   FreeJsObj.Close
		   Set FreeJsObj = Nothing
		
			RevertFreeJsObj.MoveNext
		Loop
		RevertFreeJsObj.Close
		Set RevertFreeJsObj = Nothing
		'------------------------------------------------------
		if Err.Number = 0 then
			Response.Write("<script>window.close();</script>")
		else
			Alert "��ԭʧ��"
		end if
	elseif OperateType = "News" then
		if ParentID = "" then
			Sql = "Update FS_News Set DelTF=0 where NewsID='" & ID & "'"
		else
			Sql = "Update FS_News Set DelTF=0,ClassID='" & ParentID & "' where NewsID='" & ID & "'"
		end if
		Conn.Execute(Sql)
		Conn.Execute("Update FS_FreeJsFile set DelFlag=0 where FileName=(Select FileName from FS_News where NewsID='" & ID & "')")
		'-------------���������������JS�ļ�-------------
		TempClassIDStrs = ""
		Set RevertFreeJsObj = Conn.Execute("Select distinct JSName from FS_FreeJsFile where FileName=(Select FileName from FS_News where NewsID='" & ID & "')")
		Do while Not RevertFreeJsObj.eof
			Set FreeJsObj = Conn.Execute("Select EName,Manner from FS_FreeJS where EName='"&RevertFreeJsObj("JSName")&"'")
		    If Not FreeJsObj.eof then
				Select case FreeJsObj("Manner")
				 case "1"   WCssA FreeJsObj("EName"),True
				 case "2"   WCssB FreeJsObj("EName"),True
				 case "3"   WCssC FreeJsObj("EName"),True
				 case "4"   WCssD FreeJsObj("EName"),True
				 case "5"   WCssE FreeJsObj("EName"),True
				 case "6"   PCssA FreeJsObj("EName"),True
				 case "7"   PCssB FreeJsObj("EName"),True
				 case "8"   PCssC FreeJsObj("EName"),True
				 case "9"   PCssD FreeJsObj("EName"),True
				 case "10"   PCssE FreeJsObj("EName"),True
				 case "11"   PCssF FreeJsObj("EName"),True
				 case "12"   PCssG FreeJsObj("EName"),True
				 case "13"   PCssH FreeJsObj("EName"),True
				 case "14"   PCssI FreeJsObj("EName"),True
				 case "15"   PCssJ FreeJsObj("EName"),True
				 case "16"   PCssK FreeJsObj("EName"),True
				 case "17"   PCssL FreeJsObj("EName"),True
			   End Select
		   End If
		   FreeJsObj.Close
		   Set FreeJsObj = Nothing
		   RevertFreeJsObj.MoveNext
		Loop
		RevertFreeJsObj.Close
		Set RevertFreeJsObj = Nothing
		'------------------------------------------------------------
		if Err.Number = 0 then
			Response.Write("<script>window.close();</script>")
			Response.End
		else
			Alert "��ԭʧ��"
		end if
	else
		Alert "�������ݴ���"
	end if
end if
Set RsTempObj = Nothing
Set Conn = Nothing
%>
<script language="JavaScript">
function SubmitFun()
{
	var SelectParentID='';
	for (var i=0;i<document.all.length;i++)
	{
		if (document.all(i).name=='ParentID') SelectParentID=document.all(i).value;
	}
	window.location='?OperateType=<% = OperateType%>&ID=<% = ID %>&Result=Submit&ParentID='+SelectParentID;
}
</script>
<%
Function Alert(InfoStr)
%>
<script language="JavaScript">
alert('<% = InfoStr %>');
dialogArguments.location.reload();
window.close();
</script>
<%
End Function
Function ClassList()
	Dim ClassListObj
	Set ClassListObj = Conn.Execute("select * from FS_newsclass where DelFlag=0")
	do while Not ClassListObj.Eof
		ClassList = ClassList & "<option value="&ClassListObj("ClassID")&"" & ">" & ClassListObj("ClassCName") & "</option>" & Chr(13) & Chr(10)
		ClassListObj.MoveNext	
	loop
	ClassListObj.Close
	Set ClassListObj = Nothing
End Function
%>