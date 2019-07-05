<% Option Explicit %>
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<%
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
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

%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P070500") then Call ReturnError1()
 dim FriendLinkObj,FriendLinkContent,FLType,FLAdressList,TempStr,TempNum,i
     Set  FriendLinkObj = Server.CreateObject(G_FS_RS)
	 FriendLinkObj.open "select * from FS_FriendLink order by Type desc,ID asc",conn,1,1
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet" type="text/css">
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<script language="JavaScript" src="../../SysJS/ContentMenu.js"></script>
<body topmargin="2" onClick="SelectFriendLink();" leftmargin="2" ondragstart="return false;" onselectstart="return false;">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
		  <td width=35  align="center" alt="新建" onClick="AddFriendLink();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">新建</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="修改" onClick="EditFriendLink();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">修改</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="删除" onClick="DelFriendLink();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">删除</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="后退" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
		  <td>&nbsp; </td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
<table width="100%" height="" border="0" cellpadding="0" cellspacing="0">
  <tr>
  	<td colspan="5" height="2"></td>
  </tr>
  <tr>
        <tr> 
          
    <td width="21%" height="26" class="ButtonListLeft"> 
      <div align="center">名称</div></td>
          
    <td width="29%" height="26" class="ButtonList"> 
      <div align="center">显示内容</div></td>
          
    <td width="11%" height="26" class="ButtonList"> 
      <div align="center">类型</div></td>
          
    <td width="22%" height="26" class="ButtonList"> 
      <div align="center">显示页面</div></td>
          
    <td width="17%" height="26" class="ButtonList"> 
      <div align="center">添加时间</div></td>
  </tr>
  <%
if Not FriendLinkObj.eof then 
    Dim page_size,page_no,page_total,record_all
	page_size = 20
	page_no=request.querystring("page_no")
	if page_no<=1 or page_no="" then page_no=1
	If Request.QueryString("page_no")="" then
		page_no=1
	end if
	FriendLinkObj.PageSize=page_size
	page_total=FriendLinkObj.PageCount
	if (cint(page_no)>page_total) then page_no=page_total
	FriendLinkObj.AbsolutePage=page_no
	record_all=FriendLinkObj.RecordCount
	for i=1 to FriendLinkObj.PageSize
	if FriendLinkObj.eof then exit for
	 if FriendLinkObj("Type") = "0" then
	 	FriendLinkContent = "<a href="&FriendLinkObj("Url")&">"&left(FriendLinkObj("Content"),10)&"</a>"
		FLType = "文字"
	 else
	 	FriendLinkContent = "<a href="&FriendLinkObj("Url")&"><img src="&confimsn("Domain")& FriendLinkObj("Content") &" width=30 height=20 border=0></a>"
		FLType = "图片"
	 end if
     if Instr(1,FriendLinkObj("Address"),"1",1)<>0 then
		FLAdressList = "首页 "
	 else
		FLAdressList = ""
	 end if
	 if Instr(1,FriendLinkObj("Address"),"2",1)<>0 then
		FLAdressList = FLAdressList & "栏目 "
	  else
		FLAdressList = FLAdressList & ""
	  end if
     if Instr(1,FriendLinkObj("Address"),"3",1)<>0 then
		FLAdressList = FLAdressList & "新闻 "
	 else
		FLAdressList = FLAdressList & ""
	 end if
     if Instr(1,FriendLinkObj("Address"),"4",1)<>0 then
		FLAdressList = FLAdressList & "专题 "
	 else
		FLAdressList = FLAdressList & ""
	 end if
  %>
  <tr> 
    <td><table border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><img src="../../Images/FriendLink.gif" width="24" height="22"></td>
          <td><span class="TempletItem" FriendLinkID="<%=FriendLinkObj("ID")%>"><%=FriendLinkObj("Name")%></span></td>
        </tr>
      </table></td>
    <td height="25"><div align="center"><%=FriendLinkContent%></div></td>
    <td><div align="center"><%=FLType%></div></td>
    <td><div align="center"><%=FLAdressList%></div></td>
    <td><div align="center"><%=FriendLinkObj("AddTime")%></div></td>
  </tr>
  <tr height="5"><td colspan="5"></td></tr>
  <%FriendLinkObj.MoveNext
   Next
  %>
</table></td></tr>
<%
end if
if page_total>1 then
%>
<tr>
<td height="18">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr height="1">
	  <td width="42%" height="18"><table width="99%" border="0" cellspacing="0" cellpadding="0">
		  <tr>
		  </tr>
	   </table> </td>
	
	<td width="51%" height="25" valign="middle"> <div align="right">
		<% =  "NO.<b>"& page_no &"</b>,&nbsp;&nbsp;" %>
		<% = "Totel:<b>"& page_total &"</b>,&nbsp;RecordCounts:<b>" & record_all &"</b>&nbsp;&nbsp;&nbsp;"%>
		<%
			if Page_Total=1 then
					response.Write "&nbsp;<img src=../images/First1.gif border=0 alt=首页></img>&nbsp;"
					response.Write "&nbsp;<img src=../images/prePage.gif border=0 alt=上一页></img>&nbsp;"
					response.Write "&nbsp;<img src=../images/nextPage.gif border=0 alt=下一页></img>&nbsp;"
					response.Write "&nbsp;<img src=../images/endPage.gif border=0 alt=尾页></img>&nbsp;"
			else
				if cint(Page_No)<>1 and cint(Page_No)<>Page_Total then
					response.Write "&nbsp;<a href=?page_no=1" &"&Keywords="&Request("Keywords")&"><img src=../images/First1.gif border=0 alt=首页></img></a>&nbsp;"
					response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)-1)& "&Keywords="&Request("Keywords")&"><img src=../images/prePage.gif border=0 alt=上一页></img></a>&nbsp;"
					response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)+1)& "&Keywords="&Request("Keywords")&"><img src=../images/nextPage.gif border=0 alt=下一页></img></a>&nbsp;"
					response.Write "&nbsp;<a href=?page_no="& Page_Total &"&Keywords="&Request("Keywords")&"><img src=../images/endPage.gif border=0 alt=尾页></img></a>&nbsp;"
				elseif cint(Page_No)=1 then
					response.Write "&nbsp;<img src=../images/First1.gif border=0 alt=首页></img></a>&nbsp;"
					response.Write "&nbsp;<img src=../images/prePage.gif border=0 alt=上一页></img>&nbsp;"
					response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)+1)& "&Keywords="&Request("Keywords")&"><img src=../images/nextPage.gif border=0 alt=下一页></img></a>&nbsp;"
					response.Write "&nbsp;<a href=?page_no="& Page_Total &"&Keywords="&Request("Keywords")&"><img src=../images/endpage.gif border=0 alt=尾页></img></a>&nbsp;"
				else
					response.Write "&nbsp;<a href=?page_no=1" &"&Keywords="&Request("Keywords")&"><img src=../images/First1.gif border=0 alt=首页></img>&nbsp;"
					response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)-1) &"&Keywords="&Request("Keywords")&"><img src=../images/prePage.gif border=0 alt=上一页></img></a>&nbsp;"
					response.Write "&nbsp;<img src=../images/nextPage.gif border=0 alt=下一页></img></a>&nbsp;"
					response.Write "&nbsp;<img src=../images/endpage.gif border=0 alt=尾页></img>&nbsp;"
				end if
			end if
			%>
	</div></td>
	<td width="7%" valign="middle"><select onChange="ChangePage(this.value);" style="width:50;" name="select">
	  <% for i=1 to Page_Total %>
	  <option <% if cint(Page_No) = i then Response.Write("selected")%> value="<% = i %>">
	  <% = i %>
	  </option>
	  <% next %>
	</select></td>
  </tr>
</table>
</td>
</tr>
<%
end if
FriendLinkObj.close
Set FriendLinkObj = Nothing
Set Conn = Nothing
%>
</table>
</body>
</html>
<script>
var DocumentReadyTF=false;
var ListObjArray = new Array();
var ContentMenuArray=new Array();
function document.onreadystatechange()
{
	if (DocumentReadyTF) return;
	IntialListObjArray();
	InitialContentListContentMenu();
	DocumentReadyTF=true;
}
function InitialContentListContentMenu()
{
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.EditFriendLink();",'修改','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.DelFriendLink();",'删除','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('seperator','','');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('top.GetEkMainObject().location.reload();','刷新','');
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
			if (SelectContent=='') SelectContent=ListObjArray[i].Obj.FriendLinkID;
			else SelectContent=SelectContent+'***'+ListObjArray[i].Obj.FriendLinkID;
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
					if (SelectContent=='') SelectContent=ListObjArray[i].Obj.FriendLinkID;
					else SelectContent=SelectContent+'***'+ListObjArray[i].Obj.FriendLinkID;
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
		if (CurrObj.FriendLinkID!=null)
		{
			ListObjArray[ListObjArray.length]=new FolderFileObj(CurrObj,j,false);
			j++;
		}
	}
}
function SelectFriendLink()
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
function AddFriendLink()
{
	location='OrdinaryFriendLinkAdd.asp';
}
function EditFriendLink()
{
	var SelectedFriendLink='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.FriendLinkID!=null)
			{
				if (SelectedFriendLink=='') SelectedFriendLink=ListObjArray[i].Obj.FriendLinkID;
				else  SelectedFriendLink=SelectedFriendLink+'***'+ListObjArray[i].Obj.FriendLinkID;
			}
		}
	}
	if (SelectedFriendLink!='')
	{
		if (SelectedFriendLink.indexOf('***')==-1) location='OrdinaryFriendLinkModify.asp?FLID='+SelectedFriendLink;
		else alert('请选择一个操作对象');
	}
	else alert('请选择操作对象');
}
function DelFriendLink()
{
	var SelectedFriendLink='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.FriendLinkID!=null)
			{
				if (SelectedFriendLink=='') SelectedFriendLink=ListObjArray[i].Obj.FriendLinkID;
				else  SelectedFriendLink=SelectedFriendLink+'***'+ListObjArray[i].Obj.FriendLinkID;
			}
		}
	}
	if (SelectedFriendLink!='')
		OpenWindow('Frame.asp?FileName=OrdinaryFriendLinkDel.asp&FriendLinkID='+SelectedFriendLink,220,95,window);
	else alert('请选择操作对象');
	
}
function ChangePage(PageNum)
{
	window.location.href='?page_no='+PageNum;
}
</script>