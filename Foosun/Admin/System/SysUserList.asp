<% Option Explicit %>
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Const.asp"-->
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
if Not JudgePopedomTF(Session("Name"),"P040400") then Call ReturnError1()
Dim AdminSql,RsUGObj
AdminSql = "Select * from FS_Members order by RegTime desc"
Set RsUGObj = Server.CreateObject(G_FS_RS)
RsUGObj.open AdminSql,conn,1,1
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>会员列表</title>
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<script language="JavaScript" src="../../SysJS/ContentMenu.js"></script>
<body topmargin="2" leftmargin="2" onClick="SelectAdmin();" onselectstart="return false;">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=35 align="center" alt="添加会员" onClick="AddUser();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">新建</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35  align="center" alt="修改会员" onClick="EditUser();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">修改</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="删除会员" onClick="DelUser();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">删除</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="锁定会员" onClick="LockUser();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">锁定</td>
		  <td width=2 class="Gray">|</td>
          <td width=35 align="center" alt="解锁会员" onClick="UNLockUser();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">解锁</td>
		  <td width=2 class="Gray">|</td>
		  <td width=45 align="center" alt="设置会员组" onClick="SetGroupOfUser();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">设置组</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="后退" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
		  <td>&nbsp; </td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
  	<td colspan="5" height="2"></td>
  </tr>
  <tr>
  <td valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td height="26" class="ButtonListLeft"> <div align="center">会员名称</div></td>
          <td height="20" class="ButtonList"> <div align="center">组名</div></td>
          <td height="20" class="ButtonList"> <div align="center">性别</div></td>
          <td height="20" class="ButtonList"> <div align="center">真实姓名</div></td>
          <td height="20" class="ButtonList"> <div align="center">锁定</div></td>
          <td height="20" class="ButtonList"> <div align="center">注册时间</div></td>
        </tr>
        <% 
if Not RsUGObj.eof then 
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
	RsUGObj.PageSize=page_size
	page_total=RsUGObj.PageCount
	if (cint(page_no)>page_total) then page_no=page_total
	RsUGObj.AbsolutePage=page_no
	record_all=RsUGObj.RecordCount
	Dim i
	for i=1 to RsUGObj.PageSize
	if RsUGObj.eof then exit for
%>
        <tr> 
          <td height="26"><table border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><img src="../../Images/Folder/Userface.gif" width="18" height="18"></td>
                <td><span class="TempletItem" Lock="<% = RsUGObj("Lock") %>" UserID="<% = RsUGObj("ID") %>"> 
                  <% = RsUGObj("MemName") %>
            </span></td>
              </tr>
            </table></td>
          <td><div align="center"> 
              <%
Dim RsAdminGroupObj
Set RsAdminGroupObj = Conn.Execute("Select Name from FS_MemGroup where GroupID='" & RsUGObj("GroupID") & "'")
if Not RsAdminGroupObj.Eof then
	Response.Write(RsAdminGroupObj("Name"))
else
	Response.Write("未分配组")
end if
%>
            </div></td>
          <td><div align="center"> 
              <%
if RsUGObj("Sex") = 0 then
	Response.Write("男")
else
	Response.Write("女")
end if
%>
            </div></td>
          <td><div align="center"> 
              <% = RsUGObj("Name") %>
            </div></td>
          <td><div align="center"> 
              <%
	if RsUGObj("Lock") = 1 then
		Response.Write("是")
	else
		Response.Write("否")
	end if
	%>
            </div></td>
          <td> <div align="center"> 
              <% = RsUGObj("RegTime") %>
            </div></td>
        </tr>
        <%
		RsUGObj.MoveNext
	Next
end if	
%>
      </table>
  </td>
</tr>
<%if page_total>1 then%>
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
					response.Write "&nbsp;<a href=?page_no=1&page_size="& page_size &"&Keywords="&Request("Keywords")&"><img src=../images/First1.gif border=0 alt=首页></img></a>&nbsp;"
					response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)-1)&"&page_size="&page_size&"&Keywords="&Request("Keywords")&"><img src=../images/prePage.gif border=0 alt=上一页></img></a>&nbsp;"
					response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)+1)&"&page_size="&page_size&"&Keywords="&Request("Keywords")&"><img src=../images/nextPage.gif border=0 alt=下一页></img></a>&nbsp;"
					response.Write "&nbsp;<a href=?page_no="& Page_Total &"&page_size="& page_size &"&Keywords="&Request("Keywords")&"><img src=../images/endPage.gif border=0 alt=尾页></img></a>&nbsp;"
				elseif cint(Page_No)=1 then
					response.Write "&nbsp;<img src=../images/First1.gif border=0 alt=首页></img></a>&nbsp;"
					response.Write "&nbsp;<img src=../images/prePage.gif border=0 alt=上一页></img>&nbsp;"
					response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)+1)&"&page_size="& page_size &"&Keywords="&Request("Keywords")&"><img src=../images/nextPage.gif border=0 alt=下一页></img></a>&nbsp;"
					response.Write "&nbsp;<a href=?page_no="& Page_Total &"&page_size="& page_size &"&Keywords="&Request("Keywords")&"><img src=../images/endpage.gif border=0 alt=尾页></img></a>&nbsp;"
				else
					response.Write "&nbsp;<a href=?page_no=1&page_size="& page_size &"&Keywords="&Request("Keywords")&"><img src=../images/First1.gif border=0 alt=首页></img>&nbsp;"
					response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)-1)&"&page_size="& page_size &"&Keywords="&Request("Keywords")&"><img src=../images/prePage.gif border=0 alt=上一页></img></a>&nbsp;"
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
<% end if %>
</table>
</body>
</html>
<%
Set Conn = Nothing
%>
<script>
var DocumentReadyTF=false;
var ListObjArray = new Array();
var ContentMenuArray=new Array();
function document.onreadystatechange()
{
	if (DocumentReadyTF) return;
	IntialListObjArray();
	InitialClassListContentMenu();
	DocumentReadyTF=true;
}
function InitialClassListContentMenu()
{
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.AddUser();",'新建','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.EditUser();",'修改','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.DelUser();",'删除','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('seperator','','');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('top.GetEkMainObject().location.reload();','刷新','');
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
	var EventObjInArray=false,SelectUser='',DisabledContentMenuStr='';
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
			if (SelectUser=='') SelectUser=ListObjArray[i].Obj.UserID;
			else SelectUser=SelectUser+'***'+ListObjArray[i].Obj.UserID;
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
					if (SelectUser=='') SelectUser=ListObjArray[i].Obj.UserID;
					else SelectUser=SelectUser+'***'+ListObjArray[i].Obj.UserID;
				}
			}
		}
	}
	if (SelectUser=='') DisabledContentMenuStr=',修改,删除,锁定,解锁,设置组,';
	else
	{
		if (SelectUser.indexOf('***')==-1) DisabledContentMenuStr='';
		else DisabledContentMenuStr=',修改,设置组,'
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
		if (CurrObj.UserID!=null)
		{
			ListObjArray[ListObjArray.length]=new FolderFileObj(CurrObj,j,false);
			j++;
		}
	}
}
function SelectAdmin()
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
			{
				if (MaxIndex>0)
					for (i=MaxIndex-1;i<EndIndex;i++)
					{
						ListObjArray[i].Obj.className='TempletSelectItem';
						ListObjArray[i].Selected=true;
					}
				else
				{
					ListObjArray[ElIndex].Obj.className='TempletSelectItem';
					ListObjArray[ElIndex].Selected=true;
				}
			}
			else
			{
				if (ObjInArray)
				{
					for (i=EndIndex;i<MaxIndex-1;i++)
					{
						ListObjArray[i].Obj.className='TempletSelectItem';
						ListObjArray[i].Selected=true;
					}
					if (ElIndex>=0)
					{
						ListObjArray[ElIndex].Obj.className='TempletSelectItem';
						ListObjArray[ElIndex].Selected=true;
					}
				}
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
function ChangePage(PageNum)
{
	var page_size=<% = page_size %>
	window.location.href='?page_no='+PageNum+'&page_size='+page_size;
}
function AddUser()
{
	location='UserAdd.asp';
}
function EditUser()
{
	var SelectedUser='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.UserID!=null)
			{
				if (SelectedUser=='') SelectedUser=ListObjArray[i].Obj.UserID;
				else  SelectedUser=SelectedUser+'***'+ListObjArray[i].Obj.UserID;
			}
		}
	}
	if (SelectedUser!='')
	{
		if (SelectedUser.indexOf('***')==-1) location='UserModify.asp?ID='+SelectedUser;
		else alert('一次只能够修改一个会员');
	}
	else alert('请选择会员');
}
function DelUser()
{
	var SelectedUser='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.UserID!=null)
			{
				if (SelectedUser=='') SelectedUser=ListObjArray[i].Obj.UserID;
				else  SelectedUser=SelectedUser+'***'+ListObjArray[i].Obj.UserID;
			}
		}
	}
	if (SelectedUser!='')
		 OpenWindow('Frame.asp?FileName=UserDell.asp&PageTitle=删除会员&OperateType=Dell&ID='+SelectedUser,220,105,window);
	else alert('请选择会员');
}
function LockUser()
{
	var SelectedUser='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.UserID!=null)
			{
				if (SelectedUser=='') SelectedUser=ListObjArray[i].Obj.UserID;
				else  SelectedUser=SelectedUser+'***'+ListObjArray[i].Obj.UserID;
			}
		}
	}
	if (SelectedUser!='')
		 OpenWindow('Frame.asp?FileName=UserDell.asp&PageTitle=锁定会员&OperateType=Lock&ID='+SelectedUser,220,105,window);
	else alert('请选择会员');
}
function UNLockUser()
{
	var SelectedUser='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.UserID!=null)
			{
				if (SelectedUser=='') SelectedUser=ListObjArray[i].Obj.UserID;
				else  SelectedUser=SelectedUser+'***'+ListObjArray[i].Obj.UserID;
			}
		}
	}
	if (SelectedUser!='')
		 OpenWindow('Frame.asp?FileName=UserDell.asp&PageTitle=解锁会员&OperateType=UnLock&ID='+SelectedUser,220,105,window);
	else alert('请选择会员');
}
function SetGroupOfUser()
{
	var SelectedUser='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.UserID!=null)
			{
				if (SelectedUser=='') SelectedUser=ListObjArray[i].Obj.UserID;
				else  SelectedUser=SelectedUser+'***'+ListObjArray[i].Obj.UserID;
			}
		}
	}
	if (SelectedUser!='')
	{
		if (SelectedUser.indexOf('***')==-1) OpenWindow('Frame.asp?FileName=UserGroupSet.asp&PageTitle=设置会员组&ID='+SelectedUser,280,105,window);
		else alert('一次只能够设置一个会员');
	}
	else alert('请选择会员');
}
</script>