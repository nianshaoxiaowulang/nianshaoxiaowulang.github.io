<% Option Explicit %>
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
if Not ((JudgePopedomTF(Session("Name"),"P060500")) OR (JudgePopedomTF(Session("Name"),"P060600"))) then Call ReturnError()
Dim Types,RsFileObj,RsFileSql
Types = Request("Types")
if Replace(Request.Form("KeyWords"),"'","")<>"" then
	If Types = "Class" then
		if Not JudgePopedomTF(Session("Name"),"P060500") then Call ReturnError1()
		RsFileSql = "Select * from FS_SysJs where FileCName like '%"&Replace(Request.Form("KeyWords"),"'","")&"%' and FileType=1 order by AddTime desc"
	Elseif Types = "System" then
		if Not JudgePopedomTF(Session("Name"),"P060600") then Call ReturnError1()
		RsFileSql = "Select * from FS_SysJs where FileCName like '%"&Replace(Request.Form("KeyWords"),"'","")&"%' and FileType<>1 order by FileType asc,AddTime desc"
	End IF
else
	If Types = "Class" then
		if Not JudgePopedomTF(Session("Name"),"P060500") then Call ReturnError1()
		RsFileSql = "Select * from FS_SysJs where FileType=1 order by AddTime desc"
	Elseif Types = "System" then
		if Not JudgePopedomTF(Session("Name"),"P060600") then Call ReturnError1()
		RsFileSql = "Select * from FS_SysJs where FileType<>1 order by FileType asc,AddTime desc"
	End IF
end if
Set RsFileObj = Server.CreateObject(G_FS_RS)
RsFileObj.Open RsFileSql,Conn,1,1
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>模板列表</title> 
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script language="JavaScript" src="../../SysJS/PublicJS.js"></script>
<script language="JavaScript" src="../../SysJS/ContentMenu.js"></script>
<body topmargin="2" leftmargin="2" onClick="SelectJS();" onselectstart="return false;">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
		  <td width=35  align="center" alt="新建" onClick="AddJS();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">新建</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="修改" onClick="EditJS();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">修改</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="删除" onClick="DelJS();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">删除</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="后退" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
		  <td>&nbsp; </td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
<table width="100%" height="100%" border="0" cellpadding="3" cellspacing="0">
  <tr>
  	<td colspan="5" height="2"></td>
  </tr>
  <tr> 
    <td width="23%" height="26" class="ButtonListLeft">
<div align="center">中文名称</div></td>
    <td width="19%" height="26" class="ButtonList">
<div align="center">英文名称</div></td>
    <td width="18%" height="26" class="ButtonList">
<div align="center">文件类型</div></td>
    <td width="20%" height="26" class="ButtonList">
<div align="center">所属栏目</div></td>
    <td width="20%" height="26" class="ButtonList">
<div align="center">更新时间</div></td>
  </tr>
<%
if not  RsFileObj.Bof And not RsFileObj.Eof  then 
	Dim page_no,page_total,record_all,TempTypeStr,TempNewsType,TempRsObj,TempClassName,PageNums,i
	page_no=request.querystring("page_no")
    if page_no<=1 or page_no="" then page_no=1
	If Request.QueryString("page_no")="" then
		page_no=1
	end if
	RsFileObj.PageSize=20
	page_total=RsFileObj.PageCount
	if (cint(page_no)>page_total) then page_no=page_total
	RsFileObj.AbsolutePage=page_no
	record_all=RsFileObj.RecordCount
  	for i=1 to RsFileObj.PageSize
    	if RsFileObj.eof then exit for
		Dim FileTypeStr,ClassName,ClassNameObj
		If RsFileObj("FileType") = 1 then
			FileTypeStr = "栏目JS"
		elseif RsFileObj("FileType") = 2 then
			FileTypeStr = "系统JS"
		else
			FileTypeStr = "系统自带"
		end if
		Set ClassNameObj = Conn.Execute("Select ClassCName from FS_NewsClass where ClassID='"&RsFileObj("ClassID")&"'")
		If Not ClassNameObj.eof then
			ClassName = ClassNameObj("ClassCName")
		Else
			ClassName = "--"
		End If
%>
  <tr> 
    <td height="23"> <table border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><img src="../../Images/Folder/folderclosed.gif"></td>
          <td><span class="TempletItem" FileID="<%=RsFileObj("ID")%>" FileType="<%=RsFileObj("FileType")%>">
<% = RsFileObj("FileCName") %></span></td>
        </tr>
      </table>
      </td>
    <td height="23"> 
      <div align="center"><% = RsFileObj("FileName") %></div></td>
    <td height="23"> 
      <div align="center"><% = FileTypeStr %>
      </div></td>
    <td height="23"> 
      <div align="center"><% = ClassName %></div></td>
    <td height="23"> 
      <div align="center"><% = RsFileObj("AddTime") %></div></td>
  </tr>
<%
		RsFileObj.MoveNext
	Next
end if
%>
  <tr class="TempletItem"> 
    <td colspan="5">&nbsp;</td>
  </tr>
  <%if page_total>1 then%>
  <tr><td colspan="5" valign="bottom">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td valign="middle"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                
                  <td width="62%" height="25"><table width="99%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                      </tr>
                   </table> </td>
                <td height="25" valign="middle"> <div align="right">
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
		response.Write "&nbsp;<a href=?Types="&Types&"&page_no=1" &"&Keywords="&Request("Keywords")&"><img src=../images/First1.gif border=0 alt=首页></img></a>&nbsp;"
		response.Write "&nbsp;<a href=?Types="&Types&"&page_no="&cstr(cint(Page_No)-1) &"&Keywords="&Request("Keywords")&"><img src=../images/prePage.gif border=0 alt=上一页></img></a>&nbsp;"
		response.Write "&nbsp;<a href=?Types="&Types&"&page_no="&cstr(cint(Page_No)+1) &"&Keywords="&Request("Keywords")&"><img src=../images/nextPage.gif border=0 alt=下一页></img></a>&nbsp;"
		response.Write "&nbsp;<a href=?Types="&Types&"&page_no="& Page_Total &"&Keywords="&Request("Keywords")&"><img src=../images/endPage.gif border=0 alt=尾页></img></a>&nbsp;"
	elseif cint(Page_No)=1 then
		response.Write "&nbsp;<img src=../images/First1.gif border=0 alt=首页></img></a>&nbsp;"
		response.Write "&nbsp;<img src=../images/prePage.gif border=0 alt=上一页></img>&nbsp;"
		response.Write "&nbsp;<a href=?Types="&Types&"&page_no="&cstr(cint(Page_No)+1) &"&Keywords="&Request("Keywords")&"><img src=../images/nextPage.gif border=0 alt=下一页></img></a>&nbsp;"
		response.Write "&nbsp;<a href=?Types="&Types&"&page_no="& Page_Total &"&Keywords="&Request("Keywords")&"><img src=../images/endpage.gif border=0 alt=尾页></img></a>&nbsp;"
	else
		response.Write "&nbsp;<a href=?Types="&Types&"&page_no=1" &"&Keywords="&Request("Keywords")&"><img src=../images/First1.gif border=0 alt=首页></img>&nbsp;"
		response.Write "&nbsp;<a href=?Types="&Types&"&page_no="&cstr(cint(Page_No)-1) &"&Keywords="&Request("Keywords")&"><img src=../images/prePage.gif border=0 alt=上一页></img></a>&nbsp;"
		response.Write "&nbsp;<img src=../images/nextPage.gif border=0 alt=下一页></img></a>&nbsp;"
		response.Write "&nbsp;<img src=../images/endpage.gif border=0 alt=尾页></img>&nbsp;"
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
</table>
  </td></tr>
  <% end if %>
</table>
</body>
</html>
<%
RsFileObj.Close
Set RsFileObj = Nothing
Set Conn = Nothing
%>
<script>
var Type = '<% = Types %>';
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
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.EditJS();",'修改','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.DelJS();",'删除','disabled');
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
			if (SelectAds=='') SelectAds=ListObjArray[i].Obj.FileID;
			else SelectAds=SelectAds+'***'+ListObjArray[i].Obj.FileID;
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
					if (SelectAds=='') SelectAds=ListObjArray[i].Obj.FileID;
					else SelectAds=SelectAds+'***'+ListObjArray[i].Obj.FileID;
				}
			}
		}
	}
	if (SelectAds=='') DisabledContentMenuStr=',修改,删除,';
	else
	{
		if (SelectAds.indexOf('***')==-1) DisabledContentMenuStr='';
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
		if (CurrObj.FileID!=null)
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
function AddJS()
{
	location='SysJsAdd.asp?Types='+Type;
}
function EditJS()
{
	var SelectedFile='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.FileID!=null)
			{
				if (SelectedFile=='') SelectedFile=ListObjArray[i].Obj.FileID;
				else  SelectedFile=SelectedFile+'***'+ListObjArray[i].Obj.FileID;
			}
		}
	}
	if (SelectedFile!='')
	{
		if (SelectedFile.indexOf('***')==-1) location='SysJsModify.asp?Types='+Type+'&FileID='+SelectedFile;
		else alert('一次只能够修改一个JS');
	}
	else alert('请选择要修改的JS');
}
function DelJS()
{
	var SelectedFile='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.FileID!=null)
			{
				if (SelectedFile=='') SelectedFile=ListObjArray[i].Obj.FileID;
				else  SelectedFile=SelectedFile+'***'+ListObjArray[i].Obj.FileID;
			}
		}
	}
	if (SelectedFile!='')
	{
		OpenWindow('Frame.asp?FileName=SysJsDell.asp&PageTitle=删除JS&FileID='+SelectedFile,220,105,window);
	}
	else alert('请选择要删除的JS');
}
function ChangePage(PageNum)
{
	window.location.href='?Types='+Type+'&page_no='+PageNum;
}
</script>