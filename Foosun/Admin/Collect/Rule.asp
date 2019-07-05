<% Option Explicit %>
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="inc/Config.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
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
Dim DBC,Conn,CollectConn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
DBC.ConnStr = CollectDBConnectionStr
Set CollectConn = DBC.OpenConnection()
Set DBC = Nothing
'判断权限
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P080200") then Call ReturnError1()
'判断权限结束
if Request("action") = "Del" then
	if Not JudgePopedomTF(Session("Name"),"P080203") then Call ReturnError1()
	if Request("id") <> "" then CollectConn.Execute("Delete from FS_Rule where id in (" & Replace(Request("id"),"***",",") & ")")
	Response.Redirect("Rule.asp")
	Response.End
end if

if Request.Form("Result") = "add" then
	if Request.Form("SiteId")="" then 
		Response.Write("<script>alert('请选择规则应用站点');history.back();</script>")
		Response.End
	end if
	if Request.Form("RuleName")="" then 
		Response.Write("<script>alert('请填写规则名称');history.back();</script>")
		Response.End
	end if
	if Not JudgePopedomTF(Session("Name"),"P080201") then Call ReturnError1()
    Dim Sql,RsEditObj
	Set RsEditObj = Server.CreateObject ("ADODB.RecordSet")
	Sql = "Select * from FS_Rule"
	RsEditObj.Open Sql,CollectConn,1,3
	RsEditObj.AddNew
	RsEditObj("RuleName") = NoCSSHackAdmin(Request.Form("RuleName"),"规则名称")
	RsEditObj("SiteId") = Request.Form("SiteId")
	Dim KeywordSetting
	If InStr(Request.Form("KeywordSetting"),"[过滤字符串]")<>0 then
		KeywordSetting = Split(Request.Form("KeywordSetting"),"[过滤字符串]",-1,1)
		RsEditObj("HeadSeting") = KeywordSetting(0)
		RsEditObj("FootSeting") = KeywordSetting(1)
	Else
		RsEditObj("HeadSeting") = ""
		RsEditObj("FootSeting") = ""
	End If
	RsEditObj("ReContent") = Request.Form("ReContent")
	RsEditObj("AddDate") = Now()
	RsEditObj.update
	RsEditObj.close
	Set RsEditObj = Nothing
	Response.Redirect("Rule.asp")
	Response.End
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>自动替换规则设置</title>
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script language="JavaScript" src="../../SysJS/PublicJS.js"></script>
<% if Request("Action") <> "AddRule" then %>
<script language="JavaScript" src="../../SysJS/ContentMenu.js"></script>
<% end if %>
<body<% if Request("Action") <> "AddRule" then %> onClick="SelectRule();"<% end if %> leftmargin="2" topmargin="2" onselectstart="return false;">
<%
if Request("action") = "AddRule" then
	Call Add()
else
	Call Main()
end if
Sub Main()
	if Not JudgePopedomTF(Session("Name"),"P080200") then Call ReturnError1()
%>
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="35" align="center" alt="添加关键字" onClick="AddRule();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">新建</td>
			<td width=2 class="Gray">|</td>
          <td width="35" align="center" alt="修改关键字" onClick="EditRule();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">修改</td>
			<td width=2 class="Gray">|</td>
          <td width="35" align="center" alt="删除关键字" onClick="DelRule();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">删除</td>
			<td width=2 class="Gray">|</td>
		  <td width="35" align="center" alt="后退" onClick="history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
          <td>&nbsp;</td>
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
    <td width="19%" height="26" class="ButtonListLeft"> 
      <div align="center">规则名称</div></td>
    <td width="20%" height="20" class="ButtonList"> <div align="center">应用站点</div></td>
    <td width="20%" height="20" class="ButtonList"> <div align="center">时间</div></td>
  </tr>
  <%
Dim RsSite,Sitesql,CheckInfo,StrPage,Select_Count,Select_PageCount,i,ApplyStation,RsTempObj
Set RsSite = Server.CreateObject ("ADODB.RecordSet")
SiteSql = "select * from FS_Rule order by id desc"
RsSite.Open SiteSql,CollectConn,1,1
if Not RsSite.Eof then
	StrPage = Request.QueryString("Page")
	if StrPage <= 1 or StrPage = "" then 
		StrPage = 1
	else 
		StrPage = CInt(StrPage)
	end if
	RsSite.PageSize = 12
	Select_Count = RsSite.RecordCount
	Select_PageCount = RsSite.PageCount
	if StrPage > Select_PageCount then StrPage = Select_PageCount
	RsSite.AbsolutePage = CInt(StrPage)
	for i=1 to RsSite.PageSize
		if RsSite.Eof then Exit For
		if Not ISNull(RsSite("Siteid")) then
			Sql = "Select ID,SiteName from FS_Site where ID=" & RsSite("Siteid")
			Set RsTempObj = CollectConn.Execute(Sql)
			if Not RsTempObj.Eof then
				ApplyStation = RsTempObj("SiteName")
			else
				ApplyStation = "站点不存在"
			end if
			Set RsTempObj = Nothing
		else
			ApplyStation = "站点不存在"
		end if
%>
  <tr> 
    <td><table border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="../../Images/Common.gif" width="24" height="22"></td>
          <td><span class="TempletItem" RuleID="<% = RsSite("ID") %>"> 
            <% = RsSite("RuleName") %>
            </span></td>
        </tr>
      </table></td>
    <td> <div align="center">
        <% = ApplyStation %>
      </div></td>
    <td> <div align="center"> 
        <% = RsSite("AddDate") %>
      </div></td>
  </tr>
  <%
		RsSite.MoveNext
	next
  %>
  <tr> 
    <td colspan="3"> <table  width="100%" border="0" cellpadding="5" cellspacing="0">
        <tr> 
          <td height="30"> <div align="right"> 
              <%
				Response.Write"&nbsp;共<b>" & Select_PageCount & "</b>页<b>" & Select_Count & "</b>条记录，每页<b>" & RsSite.pagesize & "</b>条，本页是第<b>" & StrPage &"</b>页"
				if Int(StrPage)>1 then
					Response.Write "&nbsp;<a href=?Page=1>第一页</a>&nbsp;"
					Response.Write "&nbsp;<a href=?Page=" & CStr(CInt(StrPage) - 1) & ">上一页</a>&nbsp;"
				end if
				if Int(StrPage) < Select_PageCount then
					Response.Write "&nbsp;<a href=?Page=" & CStr(CInt(StrPage) + 1 ) & ">下一页</a>"
					Response.Write "&nbsp;<a href=?Page="& Select_PageCount &">最后一页</a>&nbsp;"
				end if
				Response.Write"<br>"
				RsSite.close
				Set RsSite = Nothing
				%>
            </div></td>
        </tr>
      </table></td>
  </tr>
<% 
end if
%>
</table>
<%End Sub%>
<%
Sub Add()
	if Not JudgePopedomTF(Session("Name"),"P080201") then Call ReturnError1()
	Dim SiteList,RsSiteObj
	Set RsSiteObj = Server.CreateObject("Adodb.RecordSet")
	RsSiteObj.Source = "Select ID,SiteName from FS_Site order by id desc"
	RsSiteObj.open RsSiteObj.Source,CollectConn,1,3
	do while Not RsSiteObj.Eof
		SiteList = SiteList & "<option value=" & RsSiteObj("ID") & "" & ">" & RsSiteObj("SiteName") & "</option><br>"
		RsSiteObj.MoveNext	
	loop
	RsSiteObj.Close
	Set RsSiteObj = Nothing
%>
<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="dddddd">
  <form name="form1" method="post" action="" id="form1">
    <tr bgcolor="#FFFFFF"> 
      <td height="25" colspan="5" valign="middle"> 
        <table width="100%" height="25" border="0" cellpadding="0" cellspacing="0" bgcolor="#EEEEEE">
          <tr> 
            <td width="35" align="center" alt="保存" onClick="document.form1.submit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">保存</td>
			<td width=2 class="Gray">|</td>
		    <td width="35" align="center" alt="后退" onClick="history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
            <td>&nbsp; <input name="Result" type="hidden" id="Result" value="add"></td>
          </tr>
        </table></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="100" height="34"> 
        <div align="center">规则名称</div></td>
      <td bgcolor="#FFFFFF"> 
        <input style="width:100%;" name="RuleName" type="text" id="RuleName"> 
        <div align="right"></div></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="34"> 
        <div align="center">应用到</div></td>
      <td> 
        <select style="width:100%;" name="SiteId" id="select">
          <% = SiteList %>
        </select></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="110"> 
        <div align="center">过滤字符串</div></td>
      <td> &nbsp;&nbsp;输入区域： <span onClick="if(document.Form1.KeywordSetting.rows>2)document.Form1.KeywordSetting.rows-=1" style='cursor:hand'><b>缩小</b></span> 
        <span onClick="document.Form1.KeywordSetting.rows+=1" style='cursor:hand'><b>扩大</b></span> 
        &nbsp;&nbsp;可用标签:<font onClick="addTag('[过滤字符串]')" style="CURSOR: hand"><b>[过滤字符串]</b></font> 
        &nbsp;&nbsp;&nbsp;<font onClick="addTag('[变量]')" style="CURSOR: hand"><b>[变量]</b></font><br>
        <br>
	  <textarea name="KeywordSetting" onfocus="getActiveText(this)" onclick="getActiveText(this)"  onchange="getActiveText(this)" rows="5" id="textarea2" style="width:100%;"></textarea> 
        <div align="right"></div></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="center">替换为</div></td>
      <td> 
        <textarea name="ReContent" rows="5" style="width:100%;"></textarea></td>
    </tr>
  </form>
</table>
<%End Sub%>
</body>
</html>
<%
Set CollectConn = Nothing
Set Conn = Nothing
%>
<script language="JavaScript">
var DocumentReadyTF=false;
var ListObjArray = new Array();
var ContentMenuArray=new Array();
function document.onreadystatechange()
{
	if (DocumentReadyTF) return;
	<% if Request("Action") <> "AddRule" then %>
	IntialListObjArray();
	InitialContentListContentMenu();
	<% end if %>
	DocumentReadyTF=true;
}
function InitialContentListContentMenu()
{
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.EditRule();",'修改','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.DelRule();",'删除','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('seperator','','');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('location.reload();','刷新','');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('prompt(\'本页面路径属性\',\'<%=Request.ServerVariables("SCRIPT_NAME")%>\');','路径属性','');
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
	var EventObjInArray=false,SelectContent='',DisabledContentMenuStr='';
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
			if (SelectContent=='') SelectContent=ListObjArray[i].Obj.RuleID;
			else SelectContent=SelectContent+'***'+ListObjArray[i].Obj.RuleID;
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
					if (SelectContent=='') SelectContent=ListObjArray[i].Obj.RuleID;
					else SelectContent=SelectContent+'***'+ListObjArray[i].Obj.RuleID;
				}
			}
		}
	}
	if (SelectContent=='') DisabledContentMenuStr=',修改,删除,';
	else
	{
		if (SelectContent.indexOf('***')==-1) DisabledContentMenuStr='';
		else DisabledContentMenuStr=',修改,'
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
		if (CurrObj.RuleID!=null)
		{
			ListObjArray[ListObjArray.length]=new FolderFileObj(CurrObj,j,false);
			j++;
		}
	}
}
function SelectRule()
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
function AddRule()
{
	location='?Action=AddRule';
}
function EditRule()
{
	var SelectedRule='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.RuleID!=null)
			{
				if (SelectedRule=='') SelectedRule=ListObjArray[i].Obj.RuleID;
				else  SelectedRule=SelectedRule+'***'+ListObjArray[i].Obj.RuleID;
			}
		}
	}
	if (SelectedRule!='')
	{
		if (SelectedRule.indexOf('***')==-1) window.location='Rulemodify.asp?RuleId='+SelectedRule;
		else alert('请选择一个关键字');
	}
	else alert('请选择关键字');
}
function DelRule()
{
	var SelectedRule='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.RuleID!=null)
			{
				if (SelectedRule=='') SelectedRule=ListObjArray[i].Obj.RuleID;
				else  SelectedRule=SelectedRule+'***'+ListObjArray[i].Obj.RuleID;
			}
		}
	}
	if (SelectedRule!='')
	{
		if (confirm('确定要删除吗？')==true) window.location='?action=Del&Id='+SelectedRule;
	}
	else alert('请选择关键字');
}


currObj = "uuuu";
function getActiveText(obj)
{
	currObj = obj;
}

function addTag(code)
{
	addText(code);
}

function addText(ibTag)
{
	var isClose = false;
	var obj_ta = currObj;
//alert("ok");
	if (obj_ta.isTextEdit)
	{
	//alert("nooooo");
		obj_ta.focus();
		var sel = document.selection;
		var rng = sel.createRange();
		rng.colapse;

		if((sel.type == "Text" || sel.type == "None") && rng != null)
		{
			rng.text = ibTag;
		}

		obj_ta.focus();

		return isClose;
	}
	else
		return false;
}	
-->
</script>
