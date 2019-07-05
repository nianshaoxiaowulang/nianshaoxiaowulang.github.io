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
if Not JudgePopedomTF(Session("Name"),"P080510") then Call ReturnError1()
Dim Action,Sql,RsVisitObj,ID,IDArray,i
Action = Request("Action")
if Action = "Del" then
	ID = Request("ID")
	IDArray = Split(ID,"***")
	for i = LBound(IDArray) to UBound(IDArray)
		if IDArray(i) <> "" then
			Conn.Execute("Delete from FS_FlowStatistic Where ID="+IDArray(i))
		end if
	next
elseif Action = "DelTable" then
	Conn.Execute("Delete from FS_FlowStatistic")
end if
Sql = "Select * from FS_FlowStatistic Order By VisitTime Desc"
Set RsVisitObj = Server.CreateObject(G_FS_RS)
RsVisitObj.Open Sql,Conn,1,1
%>
<html>
<head>
<title>来访者信息列表</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<body topmargin="2" leftmargin="2" onClick="SelectVisit();" onselectstart="return false;">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=35 align="center" alt="删除" onClick="DelSelected();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">删除</td>
		  <td width=2 class="Gray">|</td>
		  <td width=55 align="center" alt="删除全部" onClick="DelAll();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">删除全部</td>
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
  <td valign="top">
  <table width="100%" border="0" cellpadding="2" cellspacing="0">
        <td width="15%" height="26" class="ButtonListLeft"> <div align="center">操作系统</div></td>
        <td width="17%" height="26" class="ButtonList"> <div align="center">浏览器</div></td>
        <td width="16%" height="26" class="ButtonList"> <div align="center">IP地址</div></td>
        <td width="28%" height="26" class="ButtonList"> <div align="center">地区</div></td>
        <td width="18%" height="26" class="ButtonList"> <div align="center">访问时间</div></td>
        </tr>
        <%if  RsVisitObj.Bof And RsVisitObj.Eof then%>
        <tr> 
          <td colspan="5" align="center"></td>
        </tr>
        <% else %>
        <%
if not  RsVisitObj.Bof And not RsVisitObj.Eof  then 
	Dim page_no,page_total,record_all
	page_no=request.querystring("page_no")
	if page_no<=1 or page_no="" then page_no=1	
	If Request.QueryString("page_no")="" then
		page_no=1
	end if
	RsVisitObj.PageSize=20
	page_total=RsVisitObj.PageCount
	if (cint(page_no)>page_total) then page_no=page_total
	RsVisitObj.AbsolutePage=page_no
	record_all=RsVisitObj.RecordCount
  	for i=1 to RsVisitObj.PageSize
    	if RsVisitObj.eof then exit for
%>
        <tr class="TempletItem"> 
          <td><table border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><img src="../../Images/Folder/folderclosed.gif"></td>
                <td><span class="TempletItem" VisitID="<% = RsVisitObj("ID") %>"><%=RsVisitObj("OSType")%></span></td>
              </tr>
            </table></td>
          <td><div align="center"><%=RsVisitObj("ExploreType")%></div></td>
          <td><div align="center"><%=RsVisitObj("IP")%></div></td>
          <td><div align="center"><%=RsVisitObj("Area")%></div></td>
          <td><div align="center"><%=RsVisitObj("VisitTime")%></div></td>
        </tr>
        <%
		RsVisitObj.MoveNext
	Next
end if
%>
      </table>
	</td>
	</tr>
	<tr>
<td height="18">
<table width="100%" border="0" cellspacing="0" cellpadding="0" class="ButtonListLeft">
          <td height="25" valign="middle"> 
            <div align="right"> 
              <% =  "NO.<b>"& page_no &"</b>,&nbsp;&nbsp;" %>
		<% = "Totel:<b>"& page_total &"</b>,&nbsp;RecordCounts:<b>" & record_all &"</b>&nbsp;&nbsp;&nbsp;"%>
		<%
			if Page_Total=1 then
					response.Write "&nbsp;<img src=../../Images/FirstPage.gif border=0 alt=首页></img>&nbsp;"
					response.Write "&nbsp;<img src=../../Images/prePage.gif border=0 alt=上一页></img>&nbsp;"
					response.Write "&nbsp;<img src=../../Images/nextpage.gif border=0 alt=下一页></img>&nbsp;"
					response.Write "&nbsp;<img src=../../Images/endPage.gif border=0 alt=尾页></img>&nbsp;"
			else
				if cint(Page_No)<>1 and cint(Page_No)<>Page_Total then
					response.Write "&nbsp;<a href=?page_no=1" &"&Keywords="&Request("Keywords")&"><img src=../../Images/FirstPage.gif border=0 alt=首页></img></a>&nbsp;"
					response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)-1) &"&Keywords="&Request("Keywords")&"><img src=../../Images/prePage.gif border=0 alt=上一页></img></a>&nbsp;"
					response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)+1) &"&Keywords="&Request("Keywords")&"><img src=../../Images/nextpage.gif border=0 alt=下一页></img></a>&nbsp;"
					response.Write "&nbsp;<a href=?page_no="& Page_Total &"&Keywords="&Request("Keywords")&"><img src=../../Images/endPage.gif border=0 alt=尾页></img></a>&nbsp;"
				elseif cint(Page_No)=1 then
					response.Write "&nbsp;<img src=../../Images/FirstPage.gif border=0 alt=首页></img></a>&nbsp;"
					response.Write "&nbsp;<img src=../../Images/prePage.gif border=0 alt=上一页></img>&nbsp;"
					response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)+1) &"&Keywords="&Request("Keywords")&"><img src=../../Images/nextpage.gif border=0 alt=下一页></img></a>&nbsp;"
					response.Write "&nbsp;<a href=?page_no="& Page_Total &"&Keywords="&Request("Keywords")&"><img src=../../Images/endPage.gif border=0 alt=尾页></img></a>&nbsp;"
				else
					response.Write "&nbsp;<a href=?page_no=1" &"&Keywords="&Request("Keywords")&"><img src=../../Images/FirstPage.gif border=0 alt=首页></img>&nbsp;"
					response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)-1)&"&Keywords="&Request("Keywords")&"><img src=../../Images/prePage.gif border=0 alt=上一页></img></a>&nbsp;"
					response.Write "&nbsp;<img src=../../Images/nextpage.gif border=0 alt=下一页></img></a>&nbsp;"
					response.Write "&nbsp;<img src=../../Images/endPage.gif border=0 alt=尾页></img>&nbsp;"
				end if
			end if
			%>
	</div></td>
	      <td width="100" valign="middle">
<select onChange="ChangePage(this.value);" style="width:50;" name="select">
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
  </table>
  <%end if%>
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
		if (CurrObj.VisitID!=null)
		{
			ListObjArray[ListObjArray.length]=new FolderFileObj(CurrObj,j,false);
			j++;
		}
	}
}
function SelectVisit()
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
function DelSelected()
{
	var SelectedVisit='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.VisitID!=null)
			{
				if (SelectedVisit=='') SelectedVisit=ListObjArray[i].Obj.VisitID;
				else  SelectedVisit=SelectedVisit+'***'+ListObjArray[i].Obj.VisitID;
			}
		}
	}
	if (SelectedVisit!='')
	{
		if (confirm('确定要删除吗？'))location='?Action=Del&ID='+SelectedVisit;
	}
	else alert('请选择要删除的记录');
}

function DelAll()
{
	if (confirm('确定要删除吗？'))location='?Action=DelTable';
}
function ChangePage(PageNum)
{
	window.location.href='?page_no='+PageNum;
}
</script>
<%
Set Conn = Nothing
RsVisitObj.close
Set RsVisitObj = Nothing
%>