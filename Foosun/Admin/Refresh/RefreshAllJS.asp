<% Option Explicit %>
<!--#include file="../../../Inc/Function.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<%
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P031100") then Call ReturnError1()
Dim Sql,RsLogObj,RsClassJSObj
if Replace(Request.Form("KeyWords"),"'","")<>"" then
	Sql = "Select * from FS_SysJs where FileName like '%"&Replace(Request.Form("KeyWords"),"'","")&"%' order by FileType asc,AddTime desc"
else
	Sql = "Select * from FS_SysJs order by FileType asc,AddTime desc"
end if
Set RsClassJSObj = Server.CreateObject(G_FS_RS)
RsClassJSObj.Open Sql,Conn,1,1
%>
<html>
<head>
<title>生成全部JS</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<script language="JavaScript" src="../../SysJS/ContentMenu.js"></script>
<body topmargin="2" leftmargin="2" onClick="SelectJS();" onselectstart="return false;">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=35 align="center" alt="生成" onClick="Refresh();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">生成</td>
		  <td width=2 class="Gray">|</td>
		  <td width=55 align="center" alt="生成全部" onClick="RefreshAll();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">生成全部</td>
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
    <td width="19%" height="26" class="ButtonListLeft">
<div align="center">中文名称</div></td>
    <td width="18%" height="25" class="ButtonList"><div align="center">英文名称</div></td>
    <td width="14%" height="25" class="ButtonList"><div align="center">文件类型</div></td>
    <td width="12%" height="25" class="ButtonList"><div align="center">新闻类型</div></td>
    <td width="14%" height="25" class="ButtonList"><div align="center">所属栏目</div></td>
    <td width="17%" height="25" class="ButtonList"><div align="center">更新时间</div></td>
  </tr>
  <%
if not  RsClassJSObj.Bof And not RsClassJSObj.Eof  then 
	Dim page_no,page_total,record_all,TempTypeStr,TempNewsType,TempRsObj,TempClassName,PageNums
	page_no=request.querystring("page_no")
    if page_no<=1 or page_no="" then page_no=1
	If Request.QueryString("page_no")="" then
		page_no=1
	end if
	RsClassJSObj.PageSize=20
	page_total=RsClassJSObj.PageCount
	if (cint(page_no)>page_total) then page_no=page_total
	RsClassJSObj.AbsolutePage=page_no
	record_all=RsClassJSObj.RecordCount
	Dim i
  	for i=1 to RsClassJSObj.PageSize
    	if RsClassJSObj.eof then exit for
			If RsClassJSObj("FileType")=1 then
				TempTypeStr = "栏目自定义"
			ElseIf RsClassJSObj("FileType")=2 then
				TempTypeStr = "系统自定义"
			Else
				TempTypeStr = "系统自带"
			End If
			Select Case RsClassJSObj("NewsType")
				Case "RecNews" TempNewsType = "推荐新闻"
				Case "NewNews" TempNewsType = "最新新闻"
				Case "MarqueeNews" TempNewsType = "滚动新闻"
				Case "SBSNews" TempNewsType = "并排新闻"
				Case "PicNews" TempNewsType = "图片新闻"
				Case "HotNews" TempNewsType = "热点新闻"
				Case "WordNews" TempNewsType = "文字新闻"
				Case "TitleNews" TempNewsType = "标题新闻"
				Case "ProclaimNews" TempNewsType = "公告新闻"
				Case "FilterNews" TempNewsType = "幻灯片新闻"
			End Select
			Set TempRsObj = Conn.Execute("Select ClassCName from FS_NewsClass where ClassID='"&RsClassJSObj("ClassID")&"'")
				If Not TempRsObj.eof then
					TempClassName = TempRsObj("ClassCName")
				Else
					TempClassName = "--"
				End If
%>
  <tr class="TempletItem"> 
    <td height="22"><table border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><img src="../../Images/Folder/folderclosed.gif"></td>
          <td><span JSID="<% = RsClassJSObj("ID") %>"><%=RsClassJSObj("FileCName")%></span></td>
        </tr>
      </table></td>
    <td height="22"><div align="center"><%=RsClassJSObj("FileName")%></div></td>
    <td height="22"><div align="center"><%=TempTypeStr%></div></td>
    <td height="22"><div align="center"><%=TempNewsType%></div></td>
    <td height="22"><div align="center"><%=TempClassName%></div></td>
    <td height="22"><div align="center"><%=RsClassJSObj("AddTime")%></div></td>
  </tr>
  <%
		RsClassJSObj.MoveNext
	Next
if page_total>1 then
%>
  <tr class="TempletItem"> 
    <td colspan="6">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="6" valign="bottom"> <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td valign="middle"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="62%" height="25"><table width="99%" border="0" cellspacing="0" cellpadding="0">
                    <tr> </tr>
                  </table></td>
                <td height="25" valign="middle"> <div align="right"> 
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
		response.Write "&nbsp;<a href=?page_no=1" & "&Keywords="&Request("Keywords")&"><img src=../../Images/FirstPage.gif border=0 alt=首页></img></a>&nbsp;"
		response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)-1) & "&Keywords="&Request("Keywords")&"><img src=../../Images/prePage.gif border=0 alt=上一页></img></a>&nbsp;"
		response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)+1) & "&Keywords="&Request("Keywords")&"><img src=../../Images/nextpage.gif border=0 alt=下一页></img></a>&nbsp;"
		response.Write "&nbsp;<a href=?page_no="& Page_Total & "&Keywords="&Request("Keywords")&"><img src=../../Images/endPage.gif border=0 alt=尾页></img></a>&nbsp;"
	elseif cint(Page_No)=1 then
		response.Write "&nbsp;<img src=../../Images/FirstPage.gif border=0 alt=首页></img></a>&nbsp;"
		response.Write "&nbsp;<img src=../../Images/prePage.gif border=0 alt=上一页></img>&nbsp;"
		response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)+1) & "&Keywords="&Request("Keywords")&"><img src=../../Images/nextpage.gif border=0 alt=下一页></img></a>&nbsp;"
		response.Write "&nbsp;<a href=?page_no="& Page_Total & "&Keywords="&Request("Keywords")&"><img src=../../Images/endPage.gif border=0 alt=尾页></img></a>&nbsp;"
	else
		response.Write "&nbsp;<a href=?page_no=1" & "&Keywords="&Request("Keywords")&"><img src=../../Images/FirstPage.gif border=0 alt=首页></img>&nbsp;"
		response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)-1) &"&Keywords="&Request("Keywords")&"><img src=../../Images/prePage.gif border=0 alt=上一页></img></a>&nbsp;"
		response.Write "&nbsp;<img src=../../Images/nextpage.gif border=0 alt=下一页></img></a>&nbsp;"
		response.Write "&nbsp;<img src=../../Images/endPage.gif border=0 alt=尾页></img>&nbsp;"
	end if
end if
%>
                    <select onChange="ChangePage(this.value);" style="width:50;" name="select">
                      <% for i=1 to Page_Total %>
                      <option <% if cint(Page_No) = i then Response.Write("selected")%> value="<% = i %>"> 
                      <% = i %>
                      </option>
                      <% next %>
                    </select>
                  </div></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
  <%end if%>
</table>
<%end if%>
</body>
</html>
<script language="JavaScript">
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
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.Refresh();','生成','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.RefreshAll();','生成全部','disabled');
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
	var EventObjInArray=false,SelectAds='',DisabledContentMenuStr='';
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
			if (SelectAds=='') SelectAds=ListObjArray[i].Obj.JSID;
			else SelectAds=SelectAds+'***'+ListObjArray[i].Obj.JSID;
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
					if (SelectAds=='') SelectAds=ListObjArray[i].Obj.JSID;
					else SelectAds=SelectAds+'***'+ListObjArray[i].Obj.JSID;
				}
			}
		}
	}
	if (SelectAds=='') DisabledContentMenuStr=',生成,';
	else DisabledContentMenuStr='';
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
		if (CurrObj.JSID!=null)
		{
			ListObjArray[ListObjArray.length]=new FolderFileObj(CurrObj,j,false);
			j++;
		}
	}
}
function SelectJS()
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
function Refresh()
{
	var SelectedJS='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.JSID!=null)
			{
				if (SelectedJS=='') SelectedJS=ListObjArray[i].Obj.JSID;
				else  SelectedJS=SelectedJS+'***'+ListObjArray[i].Obj.JSID;
			}
		}
	}
	if (SelectedJS!='') OpenWindow('Frame.asp?FileName=SaveJsFile.asp&PageTitle=生成JS&FileID='+SelectedJS,300,110,window);
	else alert('请选择要生成的JS');
}
function RefreshAll()
{
	OpenWindow('Frame.asp?FileName=SaveJsFile.asp&PageTitle=生成JS',300,110,window);
}
function ChangePage(PageNum)
{
	window.location.href='?page_no='+PageNum;
}
</script>
<%
Set Conn = Nothing
RsClassJSObj.Close
Set RsClassJSObj = Nothing
%>