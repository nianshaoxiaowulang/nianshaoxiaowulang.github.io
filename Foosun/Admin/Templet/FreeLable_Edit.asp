<% Option Explicit %>
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<% 
Dim  DBC,Conn,TempClassListStr,TempListStr
Set  DBC = New DataBaseClass
Set  Conn = DBC.OpenConnection()
Set  DBC = Nothing
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
'============================================================================================================
%>
<!--#include file="../../../Inc/Session.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P031300") then Call ReturnError1()
'Ȩ���ж�
Dim SqlStr,Rs
Dim Action,IFreeLableID,ISql,IName,IFieldsCName,IStyleContent,IStartFlag,IEndFlag,IDescription
IFreeLableID = Replace(Request("FreeLableID"),"'","")
Action = Request("Action")
If Action <> "Submit" Then	'�ж����û���ȷ�������´�ҳ��
	If IFreeLableID <> "" Then	'�ж����½����Ǳ༭���ɱ�ǩ
		if Not JudgePopedomTF(Session("Name"),"P031302") then Call ReturnError1()
		SqlStr = "Select * from FS_FreeLable where FreeLableID = '"&IFreeLableID&"'"
		Set Rs = conn.Execute(SqlStr)
		If not Rs.eof Then
			ISql = Replace(Rs("Sql"),"*|*","'")
			IName = Rs("Name")
			IFieldsCName = Rs("FieldsCName")
			IStyleContent = Replace(Rs("StyleContent"),"*|*","'")
			IStartFlag = Rs("StartFlag")
			IEndFlag = Rs("EndFlag")
			IDescription = Rs("Description")
		else
			Response.Write("<script>alert('��Ч�����ɱ�ǩ���')</script>")
			Response.End
		End if
	Else
		if Not JudgePopedomTF(Session("Name"),"P031301") then Call ReturnError1()
		IStartFlag = "<!--��ǩ��ʼ-->"
		IEndFlag = "<!--��ǩ����-->"
	End if
Else
	IFreeLableID = Request.Form("FreeLableID")
	ISql = Request.Form("SqlPreview")
	IName = NoCSSHackAdmin(Request.Form("Name"),"��ǩ����")
	IFieldsCName = Request.Form("FieldsCName")
	IStyleContent = Request.Form("StyleContent")
	IStartFlag = Request.Form("StartFlag")
	IEndFlag = Request.Form("EndFlag")
	IDescription = Request.Form("Description")
End if

'�������û�ѡ������ݱ�,�����γ�(FS_News,FS_NewsClass)����(FS_Download,FS_NewsClass)�����
Dim EndIndexOfLastSearch,TablesStr,TableArray,TempArray
If ISql <> "" Then
	SqlStr = ISql
	If InStr(SqlStr," from") <> 0 Then
		EndIndexOfLastSearch = InStr(ISql," from ") + 6
		SqlStr = Mid(SqlStr,EndIndexOfLastSearch)
	Else
		Response.Write("<script>alert('��Ч��Sql���')</script>")
		Response.End
	End if
	if InStr(SqlStr," where ") <> 0 Then
		EndIndexOfLastSearch = InStr(SqlStr," where ")
		TablesStr = Mid(SqlStr,1,EndIndexOfLastSearch-1)
	else
		if InStr(SqlStr," Order by") <> 0 Then
			EndIndexOfLastSearch = InStr(SqlStr," Order by ")
			TablesStr = Mid(SqlStr,1,EndIndexOfLastSearch-1)
		else
			TablesStr = SqlStr
		End if
	End if
	If TablesStr <> "" Then
		TempArray = split(TablesStr,",")
		If UBound(TempArray) = 0 Then
			if TempArray(0) = "FS_NewsClass" Then
				TableArray = Array("FS_News","FS_NewsClass")
			else
				TableArray = Array(TempArray(0),"FS_NewsClass")
			End if
		else
			TableArray = Array(TempArray(0),TempArray(1))
		End if
	Else
		Response.Write("<script>alert('��Ч��Sqddl���')</script>")
		Response.End
	End if
Else
	TableArray = Array("FS_News","FS_NewsClass")
End if
'�������ݱ����
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script language="javascript" src="../../SysJs/PublicJS.js"></script>
<title>�����ǩ</title>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
</head>
<%
'����Ǳ༭��ǩ,����ó�ʼ����������SQL���,���ú����л���
   If Action = "Submit" Or IFreeLableID <> "" Then
%>
<body leftmargin="2" topmargin="2" onload="InitialForm()">
<% Else %>
<body leftmargin="2" topmargin="2">
<% End if %>
<form action="" method="post" name="FreeLableForm">
<table width="791" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td width="787" height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=44 id="PreiorStepButton" style="display:none"align="center" alt="��һ��" onClick="PreiorStep();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">��һ��</td>
		  <td width=6 id="PreiorStepButtonX" style="display:none" class="Gray">|</td>
		  <td width=44 id="NextStepButton" align="center" alt="��һ��" onClick="NextStep();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">��һ��</td>
		  <td width=6 id="NextStepButtonX" class="Gray">|</td>
		  <td width=34 id="SaveButton" style="display:none" align="center" alt="����" onClick="DoSubmit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  <td width=6 id="SaveButtonX" style="display:none" class="Gray">|</td>
		  <td width=34 align="center" alt="����" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  <td width="612">&nbsp;
		 	<input name="Action" value="Submit" type="hidden">
			<input name="FieldsCName" type="hidden">
			<input name="FreeLableID" type="hidden" value="<%=IFreeLableID%>">
		 </td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
<div id="Step_CreateSql">
<table width="791" border="0" cellpadding="0" cellspacing="1" class="tabbgcolor">
	<tr bgcolor="#ffffff"> 
	    <td height="46"> 
          <div align="center"> 
		  ���ű�FS_News���ֶ� </div></td>
	  <td>&nbsp;</td>
	  <td><div align="center">��Ŀ��(FS_NewsClass)�ֶ�</div></td>
	</tr>
	<tr bgcolor="#FFFFFF" valign="bottom"> 
	  <td height="20"><table width=96% height=100% border="0" cellpadding="0" cellspacing="0">
		  <tr align=center>
			<td width="33%" class="ButtonListLeft">�ֶ�����</td>
			<td width="10%" class="ButtonList">����</td>
			<td width="19%" class="ButtonList">����</td>
			<td width="15%" class="ButtonList" id="FieldsList_0_AddExpContainer">����</td>
			<td width="23%" class="ButtonList">����ֵ</td>
		  </tr>
		</table></td>
	  <td>&nbsp;</td>
	  <td><table width=96% height=100% border="0" cellpadding="0" cellspacing="0">
		  <tr align=center>
			<td width="33%" class="ButtonListLeft">�ֶ�����</td>
			<td width="10%" class="ButtonList">����</td>
			<td width="19%" class="ButtonList">����</td>
			<td width="15%" class="ButtonList" id="FieldsList_1_AddExpContainer">����</td>
			<td width="23%" class="ButtonList">����ֵ</td>
		  </tr>
		</table></td>
	</tr>
	<tr bgcolor="#FFFFFF"> 
	  <td width="355" height="280"> <iframe scrolling="auto" src="FreeLable_FieldsList.asp?TableName=<%=TableArray(0)%>" style="width:100%;height:100%" name="FieldsListFrame_0"></iframe> 
	  </td>
	  <td width="81"><table width="100%" border="0" cellspacing="0" cellpadding="0">
		  <tr> 
			<td height="30" align="center">
				<input name="Submitddd" type="button" onClick="CleanSelectedOfTable(1);" class="Anbut1" id="Submitddd" value="�������ű�">
			 </td>
		  </tr>
		  <tr> 
			<td height="30" align="center">
				<input name="Submit4" onClick="CleanSelectedOfTable(2);" type="button" class="Anbut1" value="������Ŀ��">
			</td>
		  </tr>
		  <tr> 
			<td height="30" align=center><input name="Submit4" onClick="document.FieldsListFrame_0.ClickTab(-1);document.FieldsListFrame_1.ClickTab(-1);" type="button" class="Anbut1" value="����"></td>
		  </tr>
		  <tr> 
			<td height="30">&nbsp;</td>
		  </tr>
		  <tr> 
			<td height="30" align="center"> <input name="Submit6" onClick="SetSearchArraySequence();" type="button" class="Anbut1" value="����������"> 
			</td>
		  </tr>
		</table></td>
	  <td width="351"> <iframe scrolling="auto" src="FreeLable_FieldsList.asp?TableName=<%=TableArray(1)%>" style="width:100%;height:100%" name="FieldsListFrame_1"></iframe> 
	  </td>
	</tr>
	<tr id="SaveBuildInfoDescription"> 
	  <td colspan="3" bgcolor="#FFFFFF" align="center">SQLԤ��</td>
	</tr>
	<tr id="SaveBuildInfoDescription"> 
	  <td colspan="3" bgcolor="#FFFFFF"><textarea name="SqlPreview" rows="5" style="width:790" readonly><%=ISql%></textarea></td>
	</tr>
	<tr bgcolor="#FF0000" id="SaveBuildInfoClue"> 
	    <td height="29" colspan="3" bgcolor="#FFFFFF">
		<font color="#FF0000">ע���ж����ֶ�,�Ƿ�������š�ɾ����ǵȵ�ֵΪ1���棩��0���٣�����PicNewsTF=1��ʾ��ͼƬ����</font>
		</td>
	</tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
        <td align="center" onClick="if(document.all.ClassList.style.display==''){document.all.ClassList.style.display = 'none';this.innerHTML='<font size=+2 color=#FF0000>[�鿴��Ŀ���ձ�]</font>';}else{document.all.ClassList.style.display = '';this.innerHTML='<font color=blue><b>[������Ŀ���ձ�]</b></font>'}" style="cursor:hand;"><font color="blue"><b>[�鿴��Ŀ���ձ�]</b></font></td>
  </tr>
  <tr id="ClassList" style="display:none">
    <td><iframe scrolling="auto" src="FreeLable_ClassList.asp" style="width:98%;height:100%"></iframe></td>
  </tr>
</table>
</div>
<div id="Step_SetStyle" style="display:none">
    <table width="790" border="0" cellpadding="4" cellspacing="1" class="tabbgcolor">
      <tr bgcolor="#ffffff"> 
        <td height=36 colspan=3>��ǩ����: 
          <%   If IFreeLableID <> "" Then %> <input type="text" name="Name" value="<%=IName%>" readonly> 
          <%   Else %> <input type="text" name="Name" value="<%=IName%>"> 
          <%   End if %> &nbsp;&nbsp;&nbsp;&nbsp; ����: 
          <input type="text" name="Description" value="<%=IDescription%>" style="width:40%">
          &nbsp;&nbsp;&nbsp;&nbsp; <input name="button" type="button" onClick="SelectStyle();" value="ѡ����ʽ"> 
          <input name="button" type="button" onClick="Refresh();" value="ˢ��"> 
          <input name="button" type="button" onClick="PreView();" value="Ԥ��"></td>
      </tr>
      <tr bgcolor="#ffffff"> 
        <td height=36 colspan="3"> 
          <!--��ǩ��ʼ��־: -->
          <input type="hidden" name="StartFlag" value="<%=IStartFlag%>"> 
          <!--��ǩ������־: -->
          <input type="hidden" name="EndFlag" value="<%=IEndFlag%>">
          ��Ч�ֶ�: 
          <select name="ValidFields" style="width:15%">
          </select> <input name="button2" type="button" onClick="InsertStr(document.all.ValidFields.value)" value="����">
          &nbsp;&nbsp;&nbsp;&nbsp; Ԥ����: 
          <select id="PreDefine" style="width:12%">
            <option value="[#Url#]">[#����·��#]</option>
            <option value="[#ClassUrl#]">[#��Ŀ·��#]</option>
            <option value="[#PicUrl#]">[#ͼƬ·��#]</option>
          </select> <input type="button" id = "PreDefineButton" onClick="InsertStr(document.all.PreDefine.value)" value="����">
          &nbsp;&nbsp;&nbsp;&nbsp; ������ʽ(�������ֶΣ�: 
          <select id="DateStyle" style="width:12%">
            <option value="($yyyy-mm-dd$)">yyyy-mm-dd</option>
            <option value="($yyyy.mm.dd$)">yyyy.mm.dd</option>
            <option value="($yyyy/mm/dd$)">yyyy/mm/dd</option>
            <option value="($mm/dd/yyyy$)">mm/dd/yyyy</option>
            <option value="($dd/mm/yyyy$)">dd/mm/yyyy</option>
            <option value="($mm-dd-yyyy$)">mm-dd-yyyy</option>
            <option value="($mm.dd.yyyy$)">mm.dd.yyyy</option>
            <option value="($mm-dd$)">mm-dd</option>
            <option value="($mm/dd$)">mm/dd</option>
            <option value="($mm.dd$)">mm.dd</option>
            <option value="($mm��dd��$)">mm��dd��</option>
            <option value="($dd��hhʱ$)">dd��hhʱ</option>
            <option value="($dd��hh��$)">dd��hh��</option>
            <option value="($hhʱmm��$)">hhʱmm��</option>
            <option value="($hh:mm$)">hh:mm</option>
            <option value="($yyyy��mm��dd��$)">yyyy��mm��dd��</option>
          </select> <input type="button" id = "DateStyleButton" onClick="InsertStr(document.all.DateStyle.value)" value="����"> 
        </td>
      </tr>
      <tr bgcolor="#ffffff"> 
        <iframe id="StyleContainer" style="display:none"></iframe>
        <td width="46%" height="373" valign="top"><iframe id="Editer_HTML" style="width:100%;height:100%"></iframe></td>
        <td valign="top" colspan="2"><iframe id="Editer_Code" style="width:100%;height:100%" onfocus = "Target='Editer_Code';"></iframe></td>
        <input type="hidden" name="StyleContent" value="<%=Replace(Replace(Replace(IStyleContent,"<","%3C"),">","%3E"),"""","%22")%>">
      </tr>
      <tr bgcolor="#ffffff"> 
        <td colspan="3"> <font color="#FF0000">ע��ѭ������{#...#}����ѭ������{*n...*}(n>0)�����¼��š��ֶ�[*...*]������(#...#)��(#Left([*News.Title*],20)#)��������ʽ($...$)��ϵͳԤ����[#...#]��Ԥ������[#����URL#]������ID��[#��ĿURL#]����ĿID��[#ͼƬURL#]��ͼƬ·��</font></td>
      </tr>
      <tr bgcolor="#ffffff">
        <td colspan="3">&nbsp;</td>
      </tr>
    </table>
</div>
</form>
<form action="FreeLable_PreView.asp" method="post" name="PreviewForm" target="_blank">
	<input name="SqlStr" type="hidden">
	<input name="StyleContent" type="hidden">
</form>
</body>
</html>
<script language="JavaScript">
var SysRootDir = "<%=SysRootDir%>";
var StyleFiles = "<%=StyleFiles%>";
var TableFieldArray=new Array();
var TableFieldsCNameArray=new Array();
var ExpressionArray=new Array();
var OrderSearchArray=new Array();
var TablesArray = new Array("<%=TableArray(0)%>","<%=TableArray(1)%>");
var QueryNum = "10";
var Target="Editer_Code";
Editer_Code.document.designMode="on";
Editer_HTML.document.designMode="off";
<%
	If Action = "Submit" Or IFreeLableID <> "" Then
%>
//����SQL���,��ʼ�����л���
function InitialForm()
{
	var SqlStr = FreeLableForm.SqlPreview.value;
	var TableFieldNameStr = "<%=IFieldsCName%>";
	var TableFieldStr = "";
	var ExpressionStr = "";
	var OrderSearchStr = "";
	var TablesStr = "",TempTablesArray;
	var EndIndexOfLastSearch = -1;
	var FirstIndexOfThisSearch = 0
	var i,j,StyleContent;
	if(SqlStr.indexOf(" Order by ") != -1)
	{
		OrderSearchStr = SqlStr.substr(SqlStr.indexOf(" Order by ")+10);
		SqlStr = SqlStr.substr(0,SqlStr.indexOf(" Order by "));
	}
	if(SqlStr.indexOf(" where ") != -1)
	{
		ExpressionStr = SqlStr.substr(SqlStr.indexOf(" where ")+7);
		SqlStr = SqlStr.substr(0,SqlStr.indexOf(" where "));
	}
	if(SqlStr.indexOf(" from ") != -1)
	{
		TablesStr = SqlStr.substr(SqlStr.indexOf(" from ")+6);
		SqlStr = SqlStr.substr(0,SqlStr.indexOf(" from "));
	}
	else
		return;
	SqlStr = SqlStr.substr(7);
	if(SqlStr.indexOf("Top ") != -1)
	{
		SqlStr = SqlStr.substr(4);
		QueryNum = SqlStr.substr(0,SqlStr.indexOf(" "));
		TableFieldStr = SqlStr.substr(SqlStr.indexOf(" ")+1);
	}
	else
	{
		QueryNum = "";
		TableFieldStr = SqlStr;
	}

	TempTablesArray = TablesStr.split(",");
	TableFieldArray = TableFieldStr.replace(/\ /g,"").split(",");
	if(ExpressionStr != "")
		ExpressionArray = ExpressionStr.split(" and ");
	if(OrderSearchStr != "")
		OrderSearchArray = OrderSearchStr.split(",");
	TableFieldsCNameArray = TableFieldNameStr.split(",");
	if(TempTablesArray.length == 1)
	{
		for(i=0;i<TableFieldArray.length;i++)
		{
			TableFieldArray[i] = TempTablesArray[0]+"."+TableFieldArray[i];
		}
		for(i=0;i<ExpressionArray.length;i++)
		{
			ExpressionArray[i] = TempTablesArray[0]+"."+ExpressionArray[i];
		}
		for(i=0;i<OrderSearchArray.length;i++)
			OrderSearchArray[i] = TempTablesArray[0]+"."+OrderSearchArray[i];
	}
	for(i=0;i<ExpressionArray.length;i++)
	{
		if(ExpressionArray[i].indexOf("=") != -1)
		{
			for(j=0;j<TablesArray.length;j++)
				if(ExpressionArray[i].indexOf(TablesArray[j]+".",ExpressionArray[i].indexOf("=")) != -1)
				{
					ExpressionArray[i] = "";
					break;
				}
		}
	}
	FieldsListFrame_0.Initial(TableFieldArray,ExpressionArray,OrderSearchArray);
	FieldsListFrame_1.Initial(TableFieldArray,ExpressionArray,OrderSearchArray);

	StyleContent = FreeLableForm.StyleContent.value.replace(/\%3C/g,"<").replace(/\%3E/g,">").replace(/\%22/g,"\"");
	FreeLableForm.StyleContent.value = "";
	Editer_HTML.document.body.innerHTML = StyleContent;
	Editer_Code.document.body.innerText = StyleContent;
	
	ShowSql();
}
<%
	End if
%>
function NextStep()
{
	if(!CanNextToStyle())
	{
		alert("û��ѡ���ֶΣ�");
		return;
	}
	document.all.PreiorStepButton.style.display='';
	document.all.PreiorStepButtonX.style.display='';
	document.all.NextStepButton.style.display='none';
	document.all.NextStepButtonX.style.display='none';
	document.all.SaveButton.style.display='';
	document.all.SaveButtonX.style.display='';
	document.all.Step_CreateSql.style.display = 'none';
	document.all.Step_SetStyle.style.display = '';
	RefreshValidFields();
}
function PreiorStep()
{
	document.all.PreiorStepButton.style.display='none';
	document.all.PreiorStepButtonX.style.display='none';
	document.all.NextStepButton.style.display='';
	document.all.NextStepButtonX.style.display='';
	document.all.SaveButton.style.display='none';
	document.all.SaveButtonX.style.display='none';
	document.all.Step_CreateSql.style.display = '';
	document.all.Step_SetStyle.style.display = 'none';
}
//Ԥ����ǩ��ʽ
function PreView()
{
	var StyleContent = "";
	switch (Target)
	{
		case "Editer_HTML":
			StyleContent = Editer_HTML.document.body.innerHTML;
			break;
		case "Editer_Code":
			StyleContent = Editer_Code.document.body.innerText;
			break;
		default:
			StyleContent = Editer_Code.document.body.innerText;
			break;
	}
	PreviewForm.SqlStr.value = FreeLableForm.SqlPreview.value;
	PreviewForm.StyleContent.value = StyleContent;
	PreviewForm.submit();
}
//���ɽ��������
function DoSubmit()
{
	var i,FieldsCName;
	FieldsCName = "";
	switch (Target)
	{
		case "Editer_HTML":
			FreeLableForm.StyleContent.value = Editer_HTML.document.body.innerHTML;
			break;
		case "Editer_Code":
			FreeLableForm.StyleContent.value = Editer_Code.document.body.innerText;
			break;
		default:
			FreeLableForm.StyleContent.value = Editer_Code.document.body.innerText;
			break;
	}

	for(i=0;i<TableFieldArray.length;i++)
		if(TableFieldArray[i] != "")
		{
			if(FieldsCName == "")
				FieldsCName = TableFieldsCNameArray[i];
			else
				FieldsCName = FieldsCName + "," + TableFieldsCNameArray[i];
		}
	FreeLableForm.FieldsCName.value = FieldsCName;
	FreeLableForm.submit();
}
//�ڱ༭��ʽ����ʱͬ�������༭����
function Refresh()
{
	if(Target != "")
	{
		switch (Target)
		{
			case "Editer_HTML":
				Editer_Code.document.body.innerText = Editer_HTML.document.body.innerHTML;
				break;
			case "Editer_Code":
				Editer_HTML.document.body.innerHTML = Editer_Code.document.body.innerText;
				break;
		}
	}

}
//ѡ���û��༭�õ���ʽ�ļ�,�����ش���StyleContainer��������ʽ�ļ�,�����ü�⺯��
function SelectStyle()
{
	var CurrPath,ReturnValue,XmlObj,StyleStr,StreamObj,XmlhttpObjName,StreamObjName
	if(SysRootDir!="")
		CurrPath = "/" + SysRootDir + "/" + StyleFiles;
	else
		CurrPath = "/" + StyleFiles;
	ReturnValue = OpenWindow("../../funpages/frame.asp?FileName=SelectStyleFrame.asp&Pagetitle=ѡ����ʽ&CurrPath="+CurrPath,600,350,window);
	if(ReturnValue == "")
		return;
	if(SysRootDir!="")
		ReturnValue ="/" + SysRootDir + ReturnValue;	
	StyleContainer.document.location.href = ReturnValue;
	CheckStyleContainer();
}
//�����ʽ�ļ��Ƿ�����,�����������װ�������༭����
function CheckStyleContainer()
{
	var StyleStr;
	if(StyleContainer.document.body.innerHTML != "")
	{
			
		StyleStr = StyleContainer.document.body.innerHTML;
		StyleContainer.document.body.innerHTML = "";
		if(StyleStr.toLowerCase().indexOf("<body") != -1)
			StyleStr = StyleStr.substr(StyleStr.indexOf(">",StyleStr.toLowerCase().indexOf("<body"))+1);
		if(StyleStr.toLowerCase().indexOf("</body>") != -1)
			StyleStr = StyleStr.substring(0,StyleStr.toLowerCase().indexOf("</body>"));

		Editer_HTML.document.body.innerHTML = StyleStr;
		Editer_Code.document.body.innerText = StyleStr;
	}
	else
		setTimeout("CheckStyleContainer();",300);
}
//������SQL���ת����һ��ʱ,���±༭����
function RefreshValidFields()
{
	var i,TableFlag = 0;
	var ValidFieldsObj = FreeLableForm.ValidFields;
	var OptionObj;
	var ValidFieldsLength = ValidFieldsObj.length;
	var DateStyle_Flag = false;
	for(i=0;i<ValidFieldsLength;i++)
	{
		ValidFieldsObj.remove(0);
	}
	for (i=0;i<TableFieldArray.length;i++)
	{
		TempArray=TableFieldArray[i].split('.');
		if (TempArray.length>=2)
		{
			if (TableFlag!==3)
			{
				switch (TableFlag)
				{
					case 0:
						if (TempArray[0]=='FS_News') TableFlag=1;
						if (TempArray[0]=='FS_NewsClass') TableFlag=2;
						if (TempArray[0]=='FS_Download') TableFlag=4;
						break;
					case 1:
						if (TempArray[0]=='FS_NewsClass') TableFlag=3;
						break;
					case 2:
						if (TempArray[0]=='FS_News') TableFlag=3;
						if (TempArray[0]=='FS_Download') TableFlag=5;
						break;
					case 4:
						if (TempArray[0]=='FS_NewsClass') TableFlag=5;
						break;
				}
			}
		}
	}
	for(i=0;i<TableFieldArray.length;i++)
	{
		if(TableFieldArray[i] != "")
		{
			if(TableFieldArray[i].toLowerCase().indexOf("adddate") != -1)
				DateStyle_Flag = true;
			OptionObj = document.createElement("option");
			if(TableFlag == 1 || TableFlag == 2 || TableFlag == 4)
			{
				OptionObj.setAttribute("value","[*"+TableFieldArray[i].substr(TableFieldArray[i].indexOf(".")+1)+"*]");
				OptionObj.setAttribute("innerText",TableFieldsCNameArray[i].substr(TableFieldsCNameArray[i].indexOf(".")+1));
			}
			else
			{
				OptionObj.setAttribute("value","[*"+TableFieldArray[i]+"*]");
				OptionObj.setAttribute("innerText",TableFieldsCNameArray[i]);
			}
			ValidFieldsObj.appendChild(OptionObj)
		}
	}
	document.all.DateStyleButton.disabled = !DateStyle_Flag
}
//��������(�ֶδ��롢ϵͳ������롢������ʽ���룩���༭����
function InsertStr(Str)
{
	if(Target != "")
	{
		switch (Target)
		{
			case "Editer_HTML":
				Editer_HTML.focus();
				if (Editer_HTML.document.selection.type.toLowerCase() != "none")
				{
					Editer_HTML.document.selection.clear() ;
				}
				Editer_HTML.document.selection.createRange().pasteHTML(Str) ; 
				Editer_Code.document.body.innerText = Editer_HTML.document.body.innerHTML;
				break;
			case "Editer_Code":
				Editer_Code.focus();
				if (Editer_Code.document.selection.type.toLowerCase() != "none")
				{
					Editer_Code.document.selection.clear() ;
				}
				Editer_Code.document.selection.createRange().pasteHTML(Str) ; 
				Editer_HTML.document.body.innerHTML = Editer_Code.document.body.innerText;
				break;
		}
	}
}
//������Ӹ�ֵ���������ӣ��ú������ֶ��б��Ӵ������
function SetAddExpContainerHTML(TableName,HTMLStr)
{
	var i;
	for(i=0;i<TablesArray.length;i++)
	{
		TableName = TableName + ".";
		if(TableName.indexOf(TablesArray[i]+".") != -1)
			document.all("FieldsList_"+i+"_AddExpContainer").innerHTML = HTMLStr;
	}
}
//�����ֶ��б��Ӵ������Ӹ�ֵ��������
function AddExpression(TableName,index)
{
	var i;
	for(i=0;i<TablesArray.length;i++)
	{
		TableName = TableName + ".";
		if(TableName.indexOf(TablesArray[i]+".") != -1)
			eval("FieldsListFrame_"+i+".AddExpression("+index+");")
	}
}
//���¸�ֵ�������ʽ���飬���ֶ��б��Ӵ������
function SetExpressionToArray(ExpArray,TableName)
{
	var i;
	if(TableName != "")
	{
		for(i=0;i<ExpressionArray.length;i++)
			if(ExpressionArray[i] != "" && ExpressionArray[i].indexOf(TableName+".") == -1)
				ExpArray[ExpArray.length] = ExpressionArray[i];
		
	}
	ExpressionArray = null;
	ExpressionArray = ExpArray;
	ShowSql();
}
//����������������飬������ֻ�����ֶ���Ϊɾ����������
function SetOrderToArray(OrderStr)
{
	var Operation = "Del",TempStr = OrderStr,flag=false;
	if(OrderStr == "") return;
	if(OrderStr.indexOf(" Asc") != -1)
	{
		Operation = "Asc";
		TempStr = OrderStr.substring(0,OrderStr.indexOf(" Asc"));
	}
	else if(OrderStr.indexOf(" Desc") != -1)
	{
		Operation = "Desc";
		TempStr = OrderStr.substring(0,OrderStr.indexOf(" Desc"));
	}
	for(i=0;i<OrderSearchArray.length;i++)
	{
		if(OrderSearchArray[i].indexOf(TempStr) !== -1)
			switch(Operation)
			{
				case "Del":
					OrderSearchArray[i] = "";
					flag = true;
					break;
				case "Asc":
					OrderSearchArray[i] = TempStr;
					flag = true;
					break;
				case "Desc":
					OrderSearchArray[i] = OrderStr;
					flag = true;
					break;
			}
	}
	if(flag == false)
		switch (Operation)
		{
			case "Asc":
				OrderSearchArray[OrderSearchArray.length] = TempStr;
				break;
			case "Desc":
				OrderSearchArray[OrderSearchArray.length] = OrderStr;
				break;
		}
	ShowSql();
}
//���ĳ���ֶ��б��Ӵ�����ѡ����ֶ�
function CleanSelectedOfTable(table)
{
	var i;
	switch(table)
	{
		case 1:
			document.FieldsListFrame_0.CleanSelected();
			for(i=0;i<TableFieldArray.length;i++)
				if(TableFieldArray[i].indexOf(TablesArray[0]+".") > -1)
				{
					TableFieldArray[i]='';
					TableFieldsCNameArray[i]='';
				}
			ShowSql();
			break;
		case 2:
			document.FieldsListFrame_1.CleanSelected();
			for(i=0;i<TableFieldArray.length;i++)
				if(TableFieldArray[i].indexOf(TablesArray[1]+".") > -1)
				{
					TableFieldArray[i]='';
					TableFieldsCNameArray[i]='';
				}
			ShowSql();
			break;
	}
}
//����ѡ���ֶΡ�������Ĵ���Ͳ�ѯ������
function SetSearchArraySequence()
{
	var OrderStr,FieldsStr,FieldsCNameStr,i,Url,TempArray
	OrderStr = "";
	FieldsStr = "";
	for(i=0;i<OrderSearchArray.length;i++)
		if(OrderSearchArray[i] !="")
			if(OrderStr == "")
				OrderStr = OrderSearchArray[i];
			else
				OrderStr = OrderStr+","+OrderSearchArray[i];
	for(i=0;i<TableFieldArray.length;i++)
		if(TableFieldArray[i] != '')
			if(FieldsStr == "")
			{
				FieldsStr = TableFieldArray[i];
				FieldsCNameStr = TableFieldsCNameArray[i];
			}
			else
			{
				FieldsStr = FieldsStr+","+TableFieldArray[i];
				FieldsCNameStr = FieldsCNameStr + ","+TableFieldsCNameArray[i];
			}

	Url = "FreeLable_SetSequence.asp?SqlStr="+QueryNum+"*"+OrderStr+"*"+FieldsStr+"*"+FieldsCNameStr;
	var ReturnValue = OpenWindow(Url,400,300,window);
	if(ReturnValue == "") return;
	OrderSearchArray = null;
	TableFieldArray = null;
	TableFieldsCNameArray = null;
	TempArray = ReturnValue.split("*");
	QueryNum = TempArray[0];
	OrderSearchArray = TempArray[1].split(",");
	TableFieldArray = TempArray[2].split(",");
	TableFieldsCNameArray = TempArray[3].split(",");
	ShowSql();
}
//�ж������ֶα�ѡ��,�Ƿ��ܽ���һ��
function CanNextToStyle()
{
	var i,flag;
	flag = false;
	for(i=0;i<TableFieldArray.length;i++)
		if(TableFieldArray[i] != "")
		{
			flag = true;
			break;
		}
	return flag;
}
//��ѡ���ֶΡ��������ʽ�����������м����ַ���
function SearchStrInSql(Flag,Str)
{
	var ReturnVal=-1,i=0;
	switch (Flag)
	{
		case 0:  //�ֶ�
			for (i=0;i<TableFieldArray.length;i++)
			{
				if (TableFieldArray[i].indexOf(Str) != -1)
				{
					ReturnVal=i;
					break;
				}
			}
			break;
		case 1:  //���ʽ
			for (i=0;i<ExpressionArray.length;i++) 
			{ 
				if (ExpressionArray[i].indexOf(Str) != -1) 
				{ 
					ReturnVal=i;
					break;
				}
			}
			break;
		case 2:  //�Ŵ�
			for (i=0;i<OrderSearchArray.length;i++) 
			{ 
				if (OrderSearchArray[i].indexOf(Str) != -1) 
				{ 
					ReturnVal=i;
					break;
				}
			}
			break;

		default :  //��������ע
			break;
	}
	return ReturnVal;
} 
//���ѡ���ֶμ����������Ƶ���Ӧ����
function AddField(FieldStr,FieldCNameStr)
{ 
	var ArrayIndex,i;
	ArrayIndex = -1;
	ArrayIndex = SearchStrInSql(0,FieldStr)
	if(ArrayIndex != -1) return;
	for (var i=0;i<TableFieldArray.length;i++)
	{
		if (TableFieldArray[i]=='') ArrayIndex=i;
		break;
	}
	if (ArrayIndex==-1)
	{
		i = TableFieldArray.length;
		TableFieldArray[i]=FieldStr;
		TableFieldsCNameArray[i]=FieldCNameStr;
	}
	else
	{
		TableFieldArray[ArrayIndex]=FieldStr;
		TableFieldsCNameArray[ArrayIndex]=FieldCNameStr;
	}
	ShowSql();	
}
//ɾ��ѡ���ֶ�
function RemoveField(fields)
{
	var ArrayIndex=-1;
	ArrayIndex=SearchStrInSql(0,fields);
	if (ArrayIndex>=0) TableFieldArray[ArrayIndex]='';
	ShowSql();
}
//�Ӹ�����������ȡ��������SQL��䲢��ʾ
function ShowSql()
{
	var OrderStr="";
	var SqlPreviewObj=document.FreeLableForm.SqlPreview;
	var FieldSql='',i=0,SearchSql='',Sql,TableStr='',TableFlag=0;TempArray='';
	for (i=0;i<TableFieldArray.length;i++)
	{
		TempArray=TableFieldArray[i].split('.');
		if (TempArray.length>=2)
		{
			if (TableFlag!==3)
			{
				switch (TableFlag)
				{
					case 0:
						if (TempArray[0]=='FS_News') TableFlag=1;
						if (TempArray[0]=='FS_NewsClass') TableFlag=2;
						if (TempArray[0]=='FS_Download') TableFlag=4;
						break;
					case 1:
						if (TempArray[0]=='FS_NewsClass') TableFlag=3;
						break;
					case 2:
						if (TempArray[0]=='FS_News') TableFlag=3;
						if (TempArray[0]=='FS_Download') TableFlag=5;
						break;
					case 4:
						if (TempArray[0]=='FS_NewsClass') TableFlag=5;
						break;
				}
			}
		}
	}
	for (i=0;i<TableFieldArray.length;i++)
	{
		if(TableFieldArray[i] != "")
			if (FieldSql=='') FieldSql=TableFieldArray[i];
			else FieldSql=FieldSql+','+TableFieldArray[i];
	}
	for (i=0;i<ExpressionArray.length && TableFlag != 0;i++)
	{
		if(ExpressionArray[i] != "")
		{
			switch(TableFlag)
			{
				case 1:
				{
					if(ExpressionArray[i].indexOf("FS_News.") == -1)
						continue;
					else
						break;
				}
				case 2:
				{
					if(ExpressionArray[i].indexOf("FS_NewsClass.") == -1)
						continue;
					else
						break;
				}
				case 3:
				{
					if(ExpressionArray[i].indexOf("FS_NewsClass.") == -1 && ExpressionArray[i].indexOf("FS_News.") == -1)
						continue;
					else
						break;
				}
				case 4:
				{
					if(ExpressionArray[i].indexOf("FS_Download.") == -1)
						continue;
					else
						break;
				}
				case 5:
				{
					if(ExpressionArray[i].indexOf("FS_NewsClass.") == -1 && ExpressionArray[i].indexOf("FS_Download.") == -1)
						continue;
					else
						break;
				}
			}
			if (SearchSql=='') SearchSql=ExpressionArray[i];
			else SearchSql=SearchSql+' and '+ExpressionArray[i];
		}
	}
	switch (TableFlag)
	{
		case 0:
			TableStr='';
			break;
		case 1:
			TableStr='FS_News';
			break;
		case 2:
			TableStr='FS_NewsClass';
			break;
		case 3:
			TableStr='FS_News,FS_NewsClass';
			if(SearchSql != "")
				SearchSql = SearchSql +" and FS_News.ClassID = FS_NewsClass.ClassID";
			else
				SearchSql = "FS_News.ClassID = FS_NewsClass.ClassID";
			break;
		case 4:
			TableStr='FS_Download';
			break;
		case 5:
			TableStr='FS_Download,FS_NewsClass';
			if(SearchSql != "")
				SearchSql = SearchSql +" and FS_Download.ClassID = FS_NewsClass.ClassID";
			else
				SearchSql = "FS_Download.ClassID = FS_NewsClass.ClassID";
			break;
			break;
		default :
			TableStr='';
			break;
	}
	
	for (i=0;i<OrderSearchArray.length && TableFlag != 0;i++)
	{
		if(OrderSearchArray[i] != "")
		{
			switch(TableFlag)
			{
				case 1:
				{
					if(OrderSearchArray[i].indexOf("FS_News.") == -1)
						continue;
					else
						break;
				}
				case 2:
				{
					if(OrderSearchArray[i].indexOf("FS_NewsClass.") == -1)
						continue;
					else
						break;
				}
				case 4:
				{
					if(OrderSearchArray[i].indexOf("FS_Download.") == -1)
						continue;
					else
						break;
				}
			}
			if (OrderStr=='') OrderStr=OrderSearchArray[i];
			else OrderStr=OrderStr+','+OrderSearchArray[i];
		}
	}
	
	if(IsNumeric(QueryNum) && TableFlag!=0) FieldSql = "Top "+QueryNum+" "+FieldSql;
	
	if (FieldSql=='') Sql='Select';
	else Sql='Select '+FieldSql+' from '+TableStr;
	if (SearchSql!='')
		Sql=Sql+' where '+SearchSql;
	if(OrderStr != "")
		Sql=Sql+' Order by '+OrderStr;
	switch(TableFlag)
	{
		case 1:
			Sql = Sql.replace(/FS_News\./g,"");
			break;
		case 2:
			Sql = Sql.replace(/FS_NewsClass\./g,"");
			break;
	}
	SqlPreviewObj.value=Sql;
}
//�ж��ַ����Ƿ�Ϊ������
function IsNumeric(Str)
{
	var i,NumericStr="0123456789";
	if(Str.substr(0,1) == "0" || Str=="") return false;
	for(i=0;i<Str.length;i++)
		if(NumericStr.indexOf(Str.substr(i,1)) == -1)
			return false;
	return true;
}

</script>
<%
If Action = "Submit" Then
'�½���������ɱ�ǩ�����ݿ�
	IFreeLableID = Replace(Request("FreeLableID"),"'","")
	ISql = Replace(Request("SqlPreview"),"'","*|*")
	IName = Replace(Request("Name"),"'","")
	IFieldsCName = Replace(Request("FieldsCName"),"'","")
	IStyleContent = Replace(Request("StyleContent"),"'","*|*")
	IStartFlag = Replace(Request("StartFlag"),"'","")
	IEndFlag = Replace(Request("EndFlag"),"'","")
	IDescription = Replace(Request("Description"),"'","")

	If ISql = "" Then
		Response.Write("<script>alert('Sql���Ϊ��')</script>")
		Response.End
	Elseif InStr(LCase(ISql),"select") = 0 Or InStr(LCase(ISql),"insert") <> 0 Or InStr(LCase(ISql),"update") <> 0 Or InStr(LCase(ISql),"drop") <> 0 Then
		Response.Write("<script>alert('Sql�������Ƿ�����')</script>")
		Response.End
	End if
	If IName = "" Then
		Response.Write("<script>alert('��ǩ����Ϊ��')</script>")
		Response.End
	End if
	If IFieldsCName = "" Then
		Response.Write("<script>alert('û��ѡ���ֶ�')</script>")
		Response.End
	End if
	If IStyleContent = "" Then
		Response.Write("<script>alert('û�����ɱ�ǩ��ʽ����')</script>")
		Response.End
	End if
	If IStartFlag = "" Then
		Response.Write("<script>alert('��ǩ��ʼ��־Ϊ��')</script>")
		Response.End
	End if
	If IEndFlag = "" Then
		Response.Write("<script>alert('��ǩ������־Ϊ��')</script>")
		Response.End
	End if
	
	If IFreeLableID <> "" Then
		SqlStr = "select * from FS_freelable where FreeLableID = '"&IFreeLableID&"'"
		Set Rs = Server.CreateObject(G_FS_RS)
		Rs.open SqlStr,conn,3,3
		If Rs.eof Then
			Response.Write("<script>alert('��Ч�����ɱ�ǩ���')</script>")
			Response.End
		Else
			Rs("Sql") = ISql
			Rs("Name") = IName
			Rs("FieldsCName") = IFieldsCName
			Rs("StyleContent") = IStyleContent
			Rs("StartFlag") = IStartFlag
			Rs("EndFlag") = IEndFlag
			Rs("Description") = IDescription
			Rs.Update
			
			Rs.Close
			Set Rs = nothing
		End if
		Response.Write("<script> location = 'Templet_FreeLable.asp';</script>")
	Else
		SqlStr = "select name from FS_freelable where name='"&IName&"'"
		Set Rs = conn.Execute(SqlStr)
		if not Rs.eof Then
			Response.Write("<script>alert('��ǩ�����Ѿ�����')</script>")
			Response.End
		End if
		SqlStr = "select * from FS_freelable where 1=0"
		Set Rs = Server.CreateObject(G_FS_RS)
		Rs.open SqlStr,conn,3,3
		IFreeLableID  = GetRandomID18()
		Rs.Addnew
		Rs("FreeLableID") = IFreeLableID
		Rs("Sql") = ISql
		Rs("Name") = IName
		Rs("FieldsCName") = IFieldsCName
		Rs("StyleContent") = IStyleContent
		Rs("StartFlag") = IStartFlag
		Rs("EndFlag") = IEndFlag
		Rs("Description") = IDescription
		Rs("AddTime") = now
		Rs.Update
		Rs.Close
		Set Rs = nothing
		Response.Write("<script>if(confirm(""���ɱ�ǩ��ӳɹ�,�Ƿ�������?"")==false) location = 'Templet_FreeLable.asp'; else {window.location='?';} </script>")
	End if
End if
Set Conn = Nothing
%>