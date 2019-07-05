<% Option Explicit %>
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
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
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P040607") then Call ReturnError1()
Dim RsLogObj,SunNumAsp
Dim LogID,Sql,LogIDArray,i,Action,DelFlag,DatAllowDate
DatAllowDate=dateadd("d",-G_FS_HoldLogDay,date())
Action = Request("Action")
if Action = "Table" then
	if Not JudgePopedomTF(Session("Name"),"P040606") then Call ReturnError1()
		If IsSqlDataBase=1 then
			Sql = "Delete from FS_Log where logintime<='" & DatAllowDate & " 23:59:59'"
		Else
			Sql = "Delete from FS_Log where logintime<=#" & DatAllowDate & " 23:59:59#"
		End If
	Conn.Execute(Sql)
elseif Action = "Del" then
	if Not JudgePopedomTF(Session("Name"),"P040606") then Call ReturnError1()
	LogID = Request("LogID")
	LogIDArray = Split(LogID,"***")
	for i = LBound(LogIDArray) to UBound(LogIDArray)
		if LogIDArray(i) <> "" then
			If IsSqlDataBase=1 then
				Sql = "Delete from FS_Log Where ID="+LogIDArray(i) & " and logintime<='" & DatAllowDate & " 23:59:59'"
			Else
				Sql = "Delete from FS_Log Where ID="+LogIDArray(i) & " and logintime<=#" & DatAllowDate & " 23:59:59#"
			End If
			Conn.Execute(Sql)
		end if
	next
end if
%>
<html>
<head>
<title>数据统计</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<body topmargin="2" leftmargin="2" onclick="SelectLog();" onselectstart="return false;" oncontextmenu="return false;">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=36 align="center" alt="删除" onClick="DelSelected();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">删除</td>
		  <td width=6 class="Gray">|</td>
		  <td width=69  align="center" alt="删除全部" onClick="DelAll();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">删除全部</td>
		  <td width=9 class="Gray">|</td>
		  <td width=39 align="center" alt="后退" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
		  <td width="826">&nbsp; </td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
<table width="100%" height="" border="0" cellpadding="2" cellspacing="0">
  <tr>
  	<td colspan="5" height="2"></td>
  </tr>
  <tr> 
    <td height="26" class="ButtonListLeft"> <div align="center">用户名</div></td>
    <td class="ButtonList"> <div align="center">密码</div></td>
    <td class="ButtonList"> <div align="center">登录IP</div></td>
    <td class="ButtonList"> <div align="center">登录日期</div></td>
    <td class="ButtonList"> <div align="center">操作系统</div></td>
  </tr>
  <%
Sql = "Select * from FS_Log Order By LoginTime Desc"
Set RsLogObj = Server.CreateObject(G_FS_RS)
RsLogObj.Open Sql,Conn,1,1
SunNumAsp = RsLogObj.RecordCount
if not  RsLogObj.Bof And not RsLogObj.Eof  then 
	Dim page_size,page_no,page_total,record_all,PageNums
	page_size=Request.QueryString("page_size")
	if page_size<=0 or page_size="" then page_size=20
	If isnumeric(Request.Form("PageNums")) then
		if Request.Form("PageNums")<>0 then
			page_size = Cint(Request.Form("PageNums"))
		end if
	End if
	page_no=request.querystring("page_no")
    if page_no<=1 or page_no="" then page_no=1
	If Request.QueryString("page_no")="" then
		page_no=1
	end if
	RsLogObj.PageSize=page_size
	page_total=RsLogObj.PageCount
	if (cint(page_no)>page_total) then page_no=page_total
	RsLogObj.AbsolutePage=page_no
	record_all=RsLogObj.RecordCount
  	for i=1 to RsLogObj.PageSize
    	if RsLogObj.eof then exit for
		if RsLogObj("Result")=0  then
%>
  <tr class="TempletItem"> 
    <td height="18"><table border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><img src="../../Images/Folder/Log.gif" width="18" height="18"></td>
          <td><span class="TempletItem" style="color:red;" LogID="<% =RsLogObj("ID") %>"><%=RsLogObj("LogUser")%></span></td>
        </tr>
      </table></td>
    <td><div align="center"><font color="red"><%=RsLogObj("ErrorPas")%></font></div></td>
    <td><div align="center"><font color="red"><%=RsLogObj("LogIP")%></font></div></td>
    <td><div align="center"><font color="red"><%=RsLogObj("LoginTime")%></font></div></td>
    <td><div align="center"><font color="red"><%=RsLogObj("OS")%></font></div></td>
  </tr>
  <%
		else
%>
  <tr class="TempletItem"> 
    <td height="18"><table border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><img src="../../Images/Folder/Log.gif" width="18" height="18"></td>
          <td><span class="TempletItem" LogID="<% =RsLogObj("ID") %>"><%=RsLogObj("LogUser")%></span></td>
        </tr>
      </table></td>
    <td><div align="center"></div></td>
    <td><div align="center"><%=RsLogObj("LogIP")%></div></td>
    <td><div align="center"><%=RsLogObj("LoginTime")%></div></td>
    <td><div align="center"><%=RsLogObj("OS")%></div></td>
  </tr>
  <%
		end if
		RsLogObj.MoveNext
	Next
end if
if page_total > 1 then
%>
  <tr>
    <td colspan="5"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td colspan="5" height="25" valign="middle"> <div align="right"> 
              <% =  "NO.<b>"& page_no &"</b>,&nbsp;&nbsp;" %>
              <% = "Totel:<b>"& page_total &"</b>,&nbsp;RecordCounts:<b>" & record_all &"</b>&nbsp;&nbsp;&nbsp;"%>
              <%
			if Page_Total=1 then
					response.Write "&nbsp;<img src=../../images/FirstPage.gif border=0 alt=首页></img>&nbsp;"
					response.Write "&nbsp;<img src=../../images/prePage.gif border=0 alt=上一页></img>&nbsp;"
					response.Write "&nbsp;<img src=../../images/nextPage.gif border=0 alt=下一页></img>&nbsp;"
					response.Write "&nbsp;<img src=../../images/endPage.gif border=0 alt=尾页></img>&nbsp;"
			else
				if cint(Page_No)<>1 and cint(Page_No)<>Page_Total then
					response.Write "&nbsp;<a href=?page_no=1&page_size="& page_size &"&Keywords="&Request("Keywords")&"><img src=../../images/FirstPage.gif border=0 alt=首页></img></a>&nbsp;"
					response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)-1)&"&page_size="&page_size&"&Keywords="&Request("Keywords")&"><img src=../../images/prePage.gif border=0 alt=上一页></img></a>&nbsp;"
					response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)+1)&"&page_size="&page_size&"&Keywords="&Request("Keywords")&"><img src=../../images/nextPage.gif border=0 alt=下一页></img></a>&nbsp;"
					response.Write "&nbsp;<a href=?page_no="& Page_Total &"&page_size="& page_size &"&Keywords="&Request("Keywords")&"><img src=../../images/endPage.gif border=0 alt=尾页></img></a>&nbsp;"
				elseif cint(Page_No)=1 then
					response.Write "&nbsp;<img src=../../images/FirstPage.gif border=0 alt=首页></img></a>&nbsp;"
					response.Write "&nbsp;<img src=../../images/prePage.gif border=0 alt=上一页></img>&nbsp;"
					response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)+1)&"&page_size="& page_size &"&Keywords="&Request("Keywords")&"><img src=../../images/nextPage.gif border=0 alt=下一页></img></a>&nbsp;"
					response.Write "&nbsp;<a href=?page_no="& Page_Total &"&page_size="& page_size &"&Keywords="&Request("Keywords")&"><img src=../../images/endpage.gif border=0 alt=尾页></img></a>&nbsp;"
				else
					response.Write "&nbsp;<a href=?page_no=1&page_size="& page_size &"&Keywords="&Request("Keywords")&"><img src=../../images/FirstPage.gif border=0 alt=首页></img>&nbsp;"
					response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)-1)&"&page_size="& page_size &"&Keywords="&Request("Keywords")&"><img src=../../images/prePage.gif border=0 alt=上一页></img></a>&nbsp;"
					response.Write "&nbsp;<img src=../../images/nextPage.gif border=0 alt=下一页></img></a>&nbsp;"
					response.Write "&nbsp;<img src=../../images/endpage.gif border=0 alt=尾页></img>&nbsp;"
				end if
			end if
			%>
            </div></td>
          <td width="50" valign="middle"> <select onChange="ChangePage(this.value);" style="width:100%;" name="select">
              <% for i=1 to Page_Total %>
              <option <% if cint(Page_No) = i then Response.Write("selected")%> value="<% = i %>"> 
              <% = i %>
              </option>
              <% next %>
            </select></td>
        </tr>
      </table></td>
  </tr>
  <%end if %>
</table>
</body>
</html>
<script language="JavaScript">
var DocumentReadyTF=false;
var ListObjArray = new Array();
function document.onreadystatechange()
{
	if (DocumentReadyTF) return;
	IntialListObjArray();
	DocumentReadyTF=true;
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
		if (CurrObj.LogID!=null)
		{
			ListObjArray[ListObjArray.length]=new FolderFileObj(CurrObj,j,false);
			j++;
		}
	}
}
function SelectLog()
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
function DelSelected()
{
	var SelectedLog='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.LogID!=null)
			{
				if (SelectedLog=='') SelectedLog=ListObjArray[i].Obj.LogID;
				else  SelectedLog=SelectedLog+'***'+ListObjArray[i].Obj.LogID;
			}
		}
	}
	if (SelectedLog!='')
	{
		if (confirm('确定要删除吗？<%=G_FS_HoldLogDay%>天之内的日志将不会被删除')) location='?Action=Del&LogID='+SelectedLog;
	}
	else alert('请选择日志');
}
function DelAll()
{
	if (confirm('确定要删除吗？<%=G_FS_HoldLogDay%>天之内的日志将不会被删除')) location='?Action=Table';
}
function ChangePage(PageNum)
{
	var page_size;
	page_size = <% =page_size %>
	window.location.href='?page_no='+PageNum+'&page_size='+page_size;
}
function PriPage()
{
	var PageNum='<% = cint(page_no) - 1 %>';
	ChangePage(PageNum);
}
function NextPage()
{
	var PageNum='<% = cint(page_no) + 1 %>';
	ChangePage(PageNum);
}
</script>
<%
Set Conn = Nothing
RsLogObj.close
Set RsLogObj = Nothing
%>