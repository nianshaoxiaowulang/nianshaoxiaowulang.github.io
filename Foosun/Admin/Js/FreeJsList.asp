<% Option Explicit %>
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<%
'==============================================================================
'软件名称：风讯网站信息管理系统
'当前版本：Foosun Content Manager System v3.1 
'最新更新：2004.12
'==============================================================================
'商业注册联系：028-85098980-601,602 技术支持：028-85098980-606、607,客户支持：608
'产品咨询QQ：159410,655071,66252421
'技术支持:所有程序使用问题，请提问到bbs.foosun.net我们将及时回答您
'程序开发：风讯开发组 & 风讯插件开发组
'论坛支持：风讯在线论坛(http://bbs.foosun.net)
'官方网站：www.Foosun.net  演示站点：test.cooin.com    
'网站建设专区：www.cooin.com
'==============================================================================
'免费版本请在新闻首页保留版权信息，并做上本站LOGO友情连接
'==============================================================================
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P060703") then Call ReturnError1()
Dim TempType,RsJSObj,JsSql,JSType,FileNum,TempNumStr,TempObj,MannerStr,JsEName,JSFlag
 JsSql = "select * from FS_FreeJS order by Type asc,ID asc"
 JSFlag = "自由JS列表"
Set RsJSObj = Server.CreateObject(G_FS_RS)
RsJSObj.Open JsSql,conn,1,1
%>
<html>
<head>
<link href="../../../CSS/FS_css.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>自由JS列表</title>
</head>
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<script language="JavaScript" src="../../SysJS/ContentMenu.js"></script>
<body topmargin="2" leftmargin="2" onclick="SelectJS();"  ondragstart="return false;" onselectstart="return false;">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
		  <td width=35  align="center" alt="新建" onClick="AddFreeJSStyle();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">新建</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="修改" onClick="EditFreeJSStyle();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">修改</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="删除" onClick="DelFreeJSStyle();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">删除</td>
		  <td width=2 class="Gray">|</td>
		  <td width=55 align="center" alt="查看此JS中的新闻" onClick="ShowJSNews();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">查看新闻</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="后退" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
		  <td>&nbsp; </td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
<table width="100%" height="94%" border="0" cellpadding="0" cellspacing="0">
  <tr>
  	<td colspan="5" height="2"></td>
  </tr>
	  <tr>
	  <td valign="top">
	  <table width="100%" border="0" cellpadding="2" cellspacing="0">
        <tr>
			
          <td width="21%" height="26" class="ButtonListLeft">
<div align="center">名称</div></td>
			
          <td width="18%" height="26" class="ButtonList">
<div align="center">类型</div></td>
			
          <td width="20%" height="26" class="ButtonList">
<div align="center">样式</div></td>
			
          <td width="21%" height="26" class="ButtonList">
<div align="center">新闻条数</div></td>
			
          <td width="20%" height="26" class="ButtonList">
<div align="center">添加时间</div></td> 
		  </tr>
  <%
if Not RsJSObj.Eof then
  Dim page_no,page_total,record_all,PageNums,i
	page_no=request.querystring("page_no")
	if page_no<=1 or page_no="" then page_no=1
	If Request.QueryString("page_no")="" then
		page_no=1
	end if
	RsJSObj.PageSize=20
	page_total=RsJSObj.PageCount
	if (cint(page_no)>page_total) then page_no=page_total
	RsJSObj.AbsolutePage=page_no
	record_all=RsJSObj.RecordCount
	for i=1 to RsJSObj.PageSize
		if RsJSObj.eof then exit for
		select  case RsJSObj("Type")
			case "0"  JSType = "文字"
			case "1"  JSType = "图片"
		 end select
		Set TempObj = Conn.Execute("select count(ID) from FS_FreeJsFile where JSName='"&RsJSObj("EName")&"'")
		if TempObj.eof then
			FileNum = "0"
		else
			FileNum = TempObj(0)
		end if
		TempNumStr = FileNum&"/"&RsJSObj("NewsNum")
		Select case RsJSObj("Manner")
		   case "1" MannerStr = "样式A"
		   case "2" MannerStr = "样式B"
		   case "3" MannerStr = "样式C"
		   case "4" MannerStr = "样式D"
		   case "5" MannerStr = "样式E"
		   case "6" MannerStr = "样式A"
		   case "7" MannerStr = "样式B"
		   case "8" MannerStr = "样式C"
		   case "9" MannerStr = "样式D"
		   case "10" MannerStr = "样式E"
		   case "11" MannerStr = "样式F"
		   case "12" MannerStr = "样式G"
		   case "13" MannerStr = "样式H"
		   case "14" MannerStr = "样式I"
		   case "15" MannerStr = "样式J"
		   case "16" MannerStr = "样式K"
		   case "17" MannerStr = "样式L"
		End Select
			%>
			  <tr> 
				
          <td> <table border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><img src="../../Images/Folder/folderclosed.gif"></td>
                <td><span JsID="<%=RsJSObj("ID")%>" class="TempletItem" onDblClick="ShowJSNews();"><%=RsJSObj("CName")%></span></td>
              </tr>
            </table>
           </td>
          <td> 
            <div align="center"><%=JSType%></div></td>
				
          <td> 
            <div align="center"><%=MannerStr%></div></td>
				
          <td> 
            <div align="center"><%=TempNumStr%></div></td>
				
          <td> 
            <div align="center"><%=RsJSObj("AddTime")%></div></td>
			  </tr>
			  <%
  
		RsJSObj.MoveNext
	next
end if
%>
</table>
</td>
<%if page_total>1 then%>
</tr>
 <tr> 
<td valign="middle"  height="10">
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
						response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)-1) &"&Keywords="&Request("Keywords")&"><img src=../images/prePage.gif border=0 alt=上一页></img></a>&nbsp;"
						response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)+1) &"&Keywords="&Request("Keywords")&"><img src=../images/nextPage.gif border=0 alt=下一页></img></a>&nbsp;"
						response.Write "&nbsp;<a href=?page_no="& Page_Total &"&Keywords="&Request("Keywords")&"><img src=../images/endPage.gif border=0 alt=尾页></img></a>&nbsp;"
					elseif cint(Page_No)=1 then
						response.Write "&nbsp;<img src=../images/First1.gif border=0 alt=首页></img></a>&nbsp;"
						response.Write "&nbsp;<img src=../images/prePage.gif border=0 alt=上一页></img>&nbsp;"
						response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)+1) &"&Keywords="&Request("Keywords")&"><img src=../images/nextPage.gif border=0 alt=下一页></img></a>&nbsp;"
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
	</table></td>
	</tr>
	<%end if%>
	</table>
</body>
</html>
<%
RsJSObj.close
set RsJSObj=nothing
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
	InitialContentListContentMenu();
	DocumentReadyTF=true;
}
function InitialContentListContentMenu()
{
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.AddFreeJSStyle();",'新建','');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.EditFreeJSStyle();",'修改','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.DelFreeJSStyle();",'删除','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.ShowJSNews();",'查看新闻','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.RefreshList();",'刷新','');
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
			if (SelectAds=='') SelectAds=ListObjArray[i].Obj.JsID;
			else SelectAds=SelectAds+'***'+ListObjArray[i].Obj.JsID;
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
					if (SelectAds=='') SelectAds=ListObjArray[i].Obj.JsID;
					else SelectAds=SelectAds+'***'+ListObjArray[i].Obj.JsID;
				}
			}
		}
	}
	if (SelectAds=='') DisabledContentMenuStr=',修改,删除,查看新闻,';
	else
	{
		if (SelectAds.indexOf('***')==-1) DisabledContentMenuStr='';
		else DisabledContentMenuStr=',修改,删除,查看新闻,'
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
		if (CurrObj.JsID!=null)
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
	window.location.href='?page_no='+PageNum;
}
function AddFreeJSStyle()
{
	top.GetEkMainObject().location.href='JsAdd.asp';
}
function EditFreeJSStyle()
{
	var SelectedJS='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.JsID!=null)
			{
				if (SelectedJS=='') SelectedJS=ListObjArray[i].Obj.JsID;
				else  SelectedJS=SelectedJS+'***'+ListObjArray[i].Obj.JsID;
			}
		}
	}
	if (SelectedJS!='')
	{
		if (SelectedJS.indexOf('***')==-1) top.GetEkMainObject().location.href='JsModify.asp?JsID='+SelectedJS
		else alert('请选择一个JS');
	}
	else alert('请选择JS');
}
function DelFreeJSStyle()
{
	var SelectedJS='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.JsID!=null)
			{
				if (SelectedJS=='') SelectedJS=ListObjArray[i].Obj.JsID;
				else  SelectedJS=SelectedJS+'***'+ListObjArray[i].Obj.JsID;
			}
		}
	}
	if (SelectedJS!='')
	{
		OpenWindow('Frame.asp?FileName=JsDell.asp&PageTitle=删除自由JS&JsID='+SelectedJS,220,95,window);
	}
}
function ShowJSNews()
{
	var SelectedJS='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.JsID!=null)
			{
				if (SelectedJS=='') SelectedJS=ListObjArray[i].Obj.JsID;
				else  SelectedJS=SelectedJS+'***'+ListObjArray[i].Obj.JsID;
			}
		}
	}
	if (SelectedJS!='')
	{
		if (SelectedJS.indexOf('***')==-1) location.href='FreeJsFileList.asp?JsID='+SelectedJS
		else alert('请选择一个JS');
	}
	else alert('请选择JS');
}
function RefreshList()
{
	location.href=location.href;
}
</script>
