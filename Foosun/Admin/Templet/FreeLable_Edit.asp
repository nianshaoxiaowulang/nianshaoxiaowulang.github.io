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
'============================================================================================================
%>
<!--#include file="../../../Inc/Session.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P031300") then Call ReturnError1()
'权限判断
Dim SqlStr,Rs
Dim Action,IFreeLableID,ISql,IName,IFieldsCName,IStyleContent,IStartFlag,IEndFlag,IDescription
IFreeLableID = Replace(Request("FreeLableID"),"'","")
Action = Request("Action")
If Action <> "Submit" Then	'判断是用户已确定还是新打开页面
	If IFreeLableID <> "" Then	'判断是新建还是编辑自由标签
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
			Response.Write("<script>alert('无效的自由标签编号')</script>")
			Response.End
		End if
	Else
		if Not JudgePopedomTF(Session("Name"),"P031301") then Call ReturnError1()
		IStartFlag = "<!--标签开始-->"
		IEndFlag = "<!--标签结束-->"
	End if
Else
	IFreeLableID = Request.Form("FreeLableID")
	ISql = Request.Form("SqlPreview")
	IName = NoCSSHackAdmin(Request.Form("Name"),"标签名称")
	IFieldsCName = Request.Form("FieldsCName")
	IStyleContent = Request.Form("StyleContent")
	IStartFlag = Request.Form("StartFlag")
	IEndFlag = Request.Form("EndFlag")
	IDescription = Request.Form("Description")
End if

'解析出用户选择的数据表,最终形成(FS_News,FS_NewsClass)或者(FS_Download,FS_NewsClass)的组合
Dim EndIndexOfLastSearch,TablesStr,TableArray,TempArray
If ISql <> "" Then
	SqlStr = ISql
	If InStr(SqlStr," from") <> 0 Then
		EndIndexOfLastSearch = InStr(ISql," from ") + 6
		SqlStr = Mid(SqlStr,EndIndexOfLastSearch)
	Else
		Response.Write("<script>alert('无效的Sql语句')</script>")
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
		Response.Write("<script>alert('无效的Sqddl语句')</script>")
		Response.End
	End if
Else
	TableArray = Array("FS_News","FS_NewsClass")
End if
'解析数据表结束
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script language="javascript" src="../../SysJs/PublicJS.js"></script>
<title>定义标签</title>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
</head>
<%
'如果是编辑标签,则调用初始化函数解析SQL语句,设置好运行环境
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
          <td width=44 id="PreiorStepButton" style="display:none"align="center" alt="上一步" onClick="PreiorStep();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">上一步</td>
		  <td width=6 id="PreiorStepButtonX" style="display:none" class="Gray">|</td>
		  <td width=44 id="NextStepButton" align="center" alt="下一步" onClick="NextStep();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">下一步</td>
		  <td width=6 id="NextStepButtonX" class="Gray">|</td>
		  <td width=34 id="SaveButton" style="display:none" align="center" alt="保存" onClick="DoSubmit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">保存</td>
		  <td width=6 id="SaveButtonX" style="display:none" class="Gray">|</td>
		  <td width=34 align="center" alt="后退" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
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
		  新闻表（FS_News）字段 </div></td>
	  <td>&nbsp;</td>
	  <td><div align="center">栏目表(FS_NewsClass)字段</div></td>
	</tr>
	<tr bgcolor="#FFFFFF" valign="bottom"> 
	  <td height="20"><table width=96% height=100% border="0" cellpadding="0" cellspacing="0">
		  <tr align=center>
			<td width="33%" class="ButtonListLeft">字段名称</td>
			<td width="10%" class="ButtonList">类型</td>
			<td width="19%" class="ButtonList">排序</td>
			<td width="15%" class="ButtonList" id="FieldsList_0_AddExpContainer">操作</td>
			<td width="23%" class="ButtonList">条件值</td>
		  </tr>
		</table></td>
	  <td>&nbsp;</td>
	  <td><table width=96% height=100% border="0" cellpadding="0" cellspacing="0">
		  <tr align=center>
			<td width="33%" class="ButtonListLeft">字段名称</td>
			<td width="10%" class="ButtonList">类型</td>
			<td width="19%" class="ButtonList">排序</td>
			<td width="15%" class="ButtonList" id="FieldsList_1_AddExpContainer">操作</td>
			<td width="23%" class="ButtonList">条件值</td>
		  </tr>
		</table></td>
	</tr>
	<tr bgcolor="#FFFFFF"> 
	  <td width="355" height="280"> <iframe scrolling="auto" src="FreeLable_FieldsList.asp?TableName=<%=TableArray(0)%>" style="width:100%;height:100%" name="FieldsListFrame_0"></iframe> 
	  </td>
	  <td width="81"><table width="100%" border="0" cellspacing="0" cellpadding="0">
		  <tr> 
			<td height="30" align="center">
				<input name="Submitddd" type="button" onClick="CleanSelectedOfTable(1);" class="Anbut1" id="Submitddd" value="消除新闻表">
			 </td>
		  </tr>
		  <tr> 
			<td height="30" align="center">
				<input name="Submit4" onClick="CleanSelectedOfTable(2);" type="button" class="Anbut1" value="消除栏目表">
			</td>
		  </tr>
		  <tr> 
			<td height="30" align=center><input name="Submit4" onClick="document.FieldsListFrame_0.ClickTab(-1);document.FieldsListFrame_1.ClickTab(-1);" type="button" class="Anbut1" value="生成"></td>
		  </tr>
		  <tr> 
			<td height="30">&nbsp;</td>
		  </tr>
		  <tr> 
			<td height="30" align="center"> <input name="Submit6" onClick="SetSearchArraySequence();" type="button" class="Anbut1" value="次序与数量"> 
			</td>
		  </tr>
		</table></td>
	  <td width="351"> <iframe scrolling="auto" src="FreeLable_FieldsList.asp?TableName=<%=TableArray(1)%>" style="width:100%;height:100%" name="FieldsListFrame_1"></iframe> 
	  </td>
	</tr>
	<tr id="SaveBuildInfoDescription"> 
	  <td colspan="3" bgcolor="#FFFFFF" align="center">SQL预览</td>
	</tr>
	<tr id="SaveBuildInfoDescription"> 
	  <td colspan="3" bgcolor="#FFFFFF"><textarea name="SqlPreview" rows="5" style="width:790" readonly><%=ISql%></textarea></td>
	</tr>
	<tr bgcolor="#FF0000" id="SaveBuildInfoClue"> 
	    <td height="29" colspan="3" bgcolor="#FFFFFF">
		<font color="#FF0000">注：判断性字段,是否标题新闻、删除标记等的值为1（真）或0（假），如PicNewsTF=1表示是图片新闻</font>
		</td>
	</tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
        <td align="center" onClick="if(document.all.ClassList.style.display==''){document.all.ClassList.style.display = 'none';this.innerHTML='<font size=+2 color=#FF0000>[查看栏目对照表]</font>';}else{document.all.ClassList.style.display = '';this.innerHTML='<font color=blue><b>[隐藏栏目对照表]</b></font>'}" style="cursor:hand;"><font color="blue"><b>[查看栏目对照表]</b></font></td>
  </tr>
  <tr id="ClassList" style="display:none">
    <td><iframe scrolling="auto" src="FreeLable_ClassList.asp" style="width:98%;height:100%"></iframe></td>
  </tr>
</table>
</div>
<div id="Step_SetStyle" style="display:none">
    <table width="790" border="0" cellpadding="4" cellspacing="1" class="tabbgcolor">
      <tr bgcolor="#ffffff"> 
        <td height=36 colspan=3>标签名称: 
          <%   If IFreeLableID <> "" Then %> <input type="text" name="Name" value="<%=IName%>" readonly> 
          <%   Else %> <input type="text" name="Name" value="<%=IName%>"> 
          <%   End if %> &nbsp;&nbsp;&nbsp;&nbsp; 描述: 
          <input type="text" name="Description" value="<%=IDescription%>" style="width:40%">
          &nbsp;&nbsp;&nbsp;&nbsp; <input name="button" type="button" onClick="SelectStyle();" value="选择样式"> 
          <input name="button" type="button" onClick="Refresh();" value="刷新"> 
          <input name="button" type="button" onClick="PreView();" value="预览"></td>
      </tr>
      <tr bgcolor="#ffffff"> 
        <td height=36 colspan="3"> 
          <!--标签开始标志: -->
          <input type="hidden" name="StartFlag" value="<%=IStartFlag%>"> 
          <!--标签结束标志: -->
          <input type="hidden" name="EndFlag" value="<%=IEndFlag%>">
          有效字段: 
          <select name="ValidFields" style="width:15%">
          </select> <input name="button2" type="button" onClick="InsertStr(document.all.ValidFields.value)" value="插入">
          &nbsp;&nbsp;&nbsp;&nbsp; 预定义: 
          <select id="PreDefine" style="width:12%">
            <option value="[#Url#]">[#新闻路径#]</option>
            <option value="[#ClassUrl#]">[#栏目路径#]</option>
            <option value="[#PicUrl#]">[#图片路径#]</option>
          </select> <input type="button" id = "PreDefineButton" onClick="InsertStr(document.all.PreDefine.value)" value="插入">
          &nbsp;&nbsp;&nbsp;&nbsp; 日期样式(须日期字段）: 
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
            <option value="($mm月dd日$)">mm月dd日</option>
            <option value="($dd日hh时$)">dd日hh时</option>
            <option value="($dd日hh点$)">dd日hh点</option>
            <option value="($hh时mm分$)">hh时mm分</option>
            <option value="($hh:mm$)">hh:mm</option>
            <option value="($yyyy年mm月dd日$)">yyyy年mm月dd日</option>
          </select> <input type="button" id = "DateStyleButton" onClick="InsertStr(document.all.DateStyle.value)" value="插入"> 
        </td>
      </tr>
      <tr bgcolor="#ffffff"> 
        <iframe id="StyleContainer" style="display:none"></iframe>
        <td width="46%" height="373" valign="top"><iframe id="Editer_HTML" style="width:100%;height:100%"></iframe></td>
        <td valign="top" colspan="2"><iframe id="Editer_Code" style="width:100%;height:100%" onfocus = "Target='Editer_Code';"></iframe></td>
        <input type="hidden" name="StyleContent" value="<%=Replace(Replace(Replace(IStyleContent,"<","%3C"),">","%3E"),"""","%22")%>">
      </tr>
      <tr bgcolor="#ffffff"> 
        <td colspan="3"> <font color="#FF0000">注：循环内容{#...#}、不循环内容{*n...*}(n>0)代表记录序号、字段[*...*]、函数(#...#)如(#Left([*News.Title*],20)#)、日期样式($...$)、系统预定义[#...#]，预定义中[#新闻URL#]需新闻ID、[#栏目URL#]需栏目ID、[#图片URL#]需图片路径</font></td>
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
//解析SQL语句,初始化运行环境
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
		alert("没有选择字段！");
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
//预览标签样式
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
//生成结果并保存
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
//在编辑样式代码时同步两个编辑窗口
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
//选择用户编辑好的样式文件,在隐藏窗体StyleContainer中下载样式文件,并调用检测函数
function SelectStyle()
{
	var CurrPath,ReturnValue,XmlObj,StyleStr,StreamObj,XmlhttpObjName,StreamObjName
	if(SysRootDir!="")
		CurrPath = "/" + SysRootDir + "/" + StyleFiles;
	else
		CurrPath = "/" + StyleFiles;
	ReturnValue = OpenWindow("../../funpages/frame.asp?FileName=SelectStyleFrame.asp&Pagetitle=选择样式&CurrPath="+CurrPath,600,350,window);
	if(ReturnValue == "")
		return;
	if(SysRootDir!="")
		ReturnValue ="/" + SysRootDir + ReturnValue;	
	StyleContainer.document.location.href = ReturnValue;
	CheckStyleContainer();
}
//检测样式文件是否下载,如果已下载则装入两个编辑窗口
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
//从生成SQL语句转入下一步时,更新编辑区域
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
//插入内容(字段代码、系统定义代码、日期样式代码）到编辑窗口
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
//更新添加赋值操作超链接，该函数由字段列表子窗体调用
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
//调用字段列表子窗体的添加赋值操作函数
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
//更新赋值操作表达式数组，由字段列表子窗体调用
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
//保存排序项到排序数组，排序项只包含字段名为删除该排序项
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
//清除某个字段列表子窗体中选择的字段
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
//设置选中字段、排序项的次序和查询的数量
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
//判断有无字段被选择,是否能进入一步
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
//在选中字段、操作表达式、排序数组中检索字符串
function SearchStrInSql(Flag,Str)
{
	var ReturnVal=-1,i=0;
	switch (Flag)
	{
		case 0:  //字段
			for (i=0;i<TableFieldArray.length;i++)
			{
				if (TableFieldArray[i].indexOf(Str) != -1)
				{
					ReturnVal=i;
					break;
				}
			}
			break;
		case 1:  //表达式
			for (i=0;i<ExpressionArray.length;i++) 
			{ 
				if (ExpressionArray[i].indexOf(Str) != -1) 
				{ 
					ReturnVal=i;
					break;
				}
			}
			break;
		case 2:  //排次
			for (i=0;i<OrderSearchArray.length;i++) 
			{ 
				if (OrderSearchArray[i].indexOf(Str) != -1) 
				{ 
					ReturnVal=i;
					break;
				}
			}
			break;

		default :  //其他，备注
			break;
	}
	return ReturnVal;
} 
//添加选中字段及其中文名称到对应数组
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
//删除选中字段
function RemoveField(fields)
{
	var ArrayIndex=-1;
	ArrayIndex=SearchStrInSql(0,fields);
	if (ArrayIndex>=0) TableFieldArray[ArrayIndex]='';
	ShowSql();
}
//从各个数组中提取数据生成SQL语句并显示
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
//判断字符串是否为正整数
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
'新建或更新自由标签到数据库
	IFreeLableID = Replace(Request("FreeLableID"),"'","")
	ISql = Replace(Request("SqlPreview"),"'","*|*")
	IName = Replace(Request("Name"),"'","")
	IFieldsCName = Replace(Request("FieldsCName"),"'","")
	IStyleContent = Replace(Request("StyleContent"),"'","*|*")
	IStartFlag = Replace(Request("StartFlag"),"'","")
	IEndFlag = Replace(Request("EndFlag"),"'","")
	IDescription = Replace(Request("Description"),"'","")

	If ISql = "" Then
		Response.Write("<script>alert('Sql语句为空')</script>")
		Response.End
	Elseif InStr(LCase(ISql),"select") = 0 Or InStr(LCase(ISql),"insert") <> 0 Or InStr(LCase(ISql),"update") <> 0 Or InStr(LCase(ISql),"drop") <> 0 Then
		Response.Write("<script>alert('Sql语句包含非法操作')</script>")
		Response.End
	End if
	If IName = "" Then
		Response.Write("<script>alert('标签名称为空')</script>")
		Response.End
	End if
	If IFieldsCName = "" Then
		Response.Write("<script>alert('没有选择字段')</script>")
		Response.End
	End if
	If IStyleContent = "" Then
		Response.Write("<script>alert('没有自由标签样式代码')</script>")
		Response.End
	End if
	If IStartFlag = "" Then
		Response.Write("<script>alert('标签开始标志为空')</script>")
		Response.End
	End if
	If IEndFlag = "" Then
		Response.Write("<script>alert('标签结束标志为空')</script>")
		Response.End
	End if
	
	If IFreeLableID <> "" Then
		SqlStr = "select * from FS_freelable where FreeLableID = '"&IFreeLableID&"'"
		Set Rs = Server.CreateObject(G_FS_RS)
		Rs.open SqlStr,conn,3,3
		If Rs.eof Then
			Response.Write("<script>alert('无效的自由标签编号')</script>")
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
			Response.Write("<script>alert('标签名称已经存在')</script>")
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
		Response.Write("<script>if(confirm(""自由标签添加成功,是否继续添加?"")==false) location = 'Templet_FreeLable.asp'; else {window.location='?';} </script>")
	End if
End if
Set Conn = Nothing
%>