<% Option Explicit %>
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<%
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
'�������2�ο��������뾭����Ѷ��˾������������׷����������
'==============================================================================

Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P020400") then Call ReturnError1()
Dim SpecialID,TempSpecialObj,FileListObj,FileClassObj
if Request("SpecialID")<>"" then
	SpecialID = Cstr(Request("SpecialID"))
else
	Response.Write("<script>alert(""�������ݴ���!!"");</script>")
	response.end
end if
Set TempSpecialObj = Conn.Execute("Select CName,SpecialID from FS_Special where SpecialID = '"&SpecialID&"'")
Set FileListObj = Server.CreateObject(G_FS_RS)
FileListObj.open "Select * from FS_News where SpecialID like '%"&SpecialID&"%' order by ID desc",Conn,1,1
If TempSpecialObj.eof then
	Response.Write("<script>alert(""�������ݴ���"");</script>")
	response.end
End if
%>
<html>
<head>
<link href="../../../CSS/FS_css.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Ƶ��/ר�������б�</title>
</head>
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<script language="JavaScript" src="../../SysJS/ContentMenu.js"></script>
<body topmargin="2" leftmargin="2" onclick="SelectSpecialFile();"  ondragstart="return false;" onselectstart="return false;">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=35 align="center" alt="ɾ��ר������" onClick="DelSpecialNews();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">ɾ��</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35  align="center" alt="����" onClick="CutOperation();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="����" onClick="CopyOperation();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="ճ��" onClick="PasteOperation();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">ճ��</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="����" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
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
    <td valign="top"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="49%" height="26" class="ButtonListLeft"> 
            <div align="center">��������</div></td>
          <td width="28%" height="26" class="ButtonList"> 
            <div align="center">������Ŀ</div></td>
          <td width="23%" height="26" class="ButtonList"> 
            <div align="center">����ʱ��</div></td>
        </tr>
<%if Not FileListObj.eof then 
		page_size = 20
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
			FileListObj.PageSize=page_size
			page_total=FileListObj.PageCount
			if (cint(page_no)>page_total) then page_no=page_total
			FileListObj.AbsolutePage=page_no
			record_all=FileListObj.RecordCount
			Dim i
			for i=1 to FileListObj.PageSize
			if FileListObj.eof then exit for
			Dim TempFlagStr
			If FileListObj("HeadNewsTF")="1" then
				TempFlagStr = "<img src=""../../Images/Info/TitleNews.gif"" border=""0"">"
			Elseif FileListObj("PicNewsTF")="0" and FileListObj("HeadNewsTF")="0" then
				TempFlagStr = "<img src=""../../Images/Info/WordNews.gif"" border=""0"">"
			else
				TempFlagStr = "<img src=""../../Images/Info/PicNews.gif"" border=""0"">"
			End if
			Set FileClassObj = Conn.Execute("Select ClassCName from FS_NewsClass where ClassID='"&FileListObj("ClassID")&"'")
		%>
        <tr> 
		<td><table border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td><% = TempFlagStr %></td>
                <td><span SpecialID="<%=TempSpecialObj("SpecialID")%>" class="TempletItem" NewsID="<%=FileListObj("NewsID")%>" align="center"><% = GotTopic(FileListObj("Title"),40) %></span></td>
              </tr>
            </table></td> 
          <td><div align="center" class="TempletItem"><%=FileClassObj("ClassCName")%></div></td>
          <td><div align="center" class="TempletItem"><%=FileListObj("AddDate")%></div></td>
        </tr>
        <%
		FileClassObj.Close
		FileListObj.MoveNext
	next
	FileListObj.close
	set FileListObj=nothing
	TempSpecialObj.close
	Set TempSpecialObj = Nothing
end if
		%>
      </table></td>
  </tr>
	  <%if page_total>1 then%>
  <tr> 
    	<td valign="middle" height="10">
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr height="1">
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
								response.Write "&nbsp;<img src=""../../Images/FirstPage.gif"" border=0 alt=��ҳ></img>&nbsp;"
								response.Write "&nbsp;<img src=""../../Images/prePage.gif"" border=0 alt=��һҳ></img>&nbsp;"
								response.Write "&nbsp;<img src=""../../Images/nextpage.gif"" border=0 alt=��һҳ></img>&nbsp;"
								response.Write "&nbsp;<img src=""../../Images/endPage.gif"" border=0 alt=βҳ></img>&nbsp;"
						else
							if cint(Page_No)<>1 and cint(Page_No)<>Page_Total then
								response.Write "&nbsp;<a href=?page_no=1&page_size="& page_size &"&SpecialID="&SpecialID&"&Keywords="&Request("Keywords")&"><img src=""../../Images/FirstPage.gif"" border=0 alt=��ҳ></img></a>&nbsp;"
								response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)-1)&"&SpecialID="&SpecialID&"&page_size="&page_size&"&Keywords="&Request("Keywords")&"><img src=""../../Images/prePage.gif"" border=0 alt=��һҳ></img></a>&nbsp;"
								response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)+1)&"&SpecialID="&SpecialID&"&page_size="&page_size&"&Keywords="&Request("Keywords")&"><img src=""../../Images/nextpage.gif"" border=0 alt=��һҳ></img></a>&nbsp;"
								response.Write "&nbsp;<a href=?page_no="& Page_Total &"&page_size="& page_size &"&SpecialID="&SpecialID&"&Keywords="&Request("Keywords")&"><img src=""../../Images/endPage.gif"" border=0 alt=βҳ></img></a>&nbsp;"
							elseif cint(Page_No)=1 then
								response.Write "&nbsp;<img src=""../../Images/FirstPage.gif"" border=0 alt=��ҳ></img></a>&nbsp;"
								response.Write "&nbsp;<img src=""../../Images/prePage.gif"" border=0 alt=��һҳ></img>&nbsp;"
								response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)+1)&"&page_size="& page_size &"&SpecialID="&SpecialID&"&Keywords="&Request("Keywords")&"><img src=""../../Images/nextpage.gif"" border=0 alt=��һҳ></img></a>&nbsp;"
								response.Write "&nbsp;<a href=?page_no="& Page_Total &"&page_size="& page_size &"&SpecialID="&SpecialID&"&Keywords="&Request("Keywords")&"><img src=""../../Images/endPage.gif"" border=0 alt=βҳ></img></a>&nbsp;"
							else
								response.Write "&nbsp;<a href=?page_no=1&page_size="& page_size &"&SpecialID="&SpecialID&"&Keywords="&Request("Keywords")&"><img src=""../../Images/FirstPage.gif"" border=0 alt=��ҳ></img>&nbsp;"
								response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)-1)&"&page_size="& page_size &"&SpecialID="&SpecialID&"&Keywords="&Request("Keywords")&"><img src=""../../Images/prePage.gif"" border=0 alt=��һҳ></img></a>&nbsp;"
								response.Write "&nbsp;<img src=""../../Images/nextpage.gif"" border=0 alt=��һҳ></img></a>&nbsp;"
								response.Write "&nbsp;<img src=""../../Images/endPage.gif"" border=0 alt=βҳ></img>&nbsp;"
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
<% end if %>
</table>
</body>
</html>
<%
Set Conn = Nothing
%>
<script>
var SpecialID='<% = SpecialID %>';
var DocumentReadyTF=false;
var ListObjArray=new Array();
var ContentMenuArray=new Array();
function document.onreadystatechange()
{
	if (DocumentReadyTF) return;
	InitialContentListContentMenu();
	DocumentReadyTF=true;
}
function ContentMenuFunction(ExeFunction,Description,EnabledStr)
{
	this.ExeFunction=ExeFunction;
	this.Description=Description;
	this.EnabledStr=EnabledStr;
}
function InitialContentListContentMenu()
{
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.DelSpecialNews();",'ɾ��','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('seperator','','');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.CutOperation();','����','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.CopyOperation();','����','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.PasteOperation();','ճ��','disabled');
	//ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('seperator','','');
	//ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('top.GetEkMainObject().location.reload();','ˢ��','');
	//ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('prompt(\'��ҳ��·������\',\'<%=Request.ServerVariables("SCRIPT_NAME")%>\');','·������','');
	IntialListObjArray();
}
function ContentMenuShowEvent()
{
	ChangeSpecialMenuStatus();
}
function ChangeSpecialMenuStatus()
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
			if (SelectContent=='') SelectContent=ListObjArray[i].Obj.NewsID;
			else SelectContent=SelectContent+'***'+ListObjArray[i].Obj.NewsID;
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
					if (SelectContent=='') SelectContent=ListObjArray[i].Obj.NewsID;
					else SelectContent=SelectContent+'***'+ListObjArray[i].Obj.NewsID;
				}
			}
		}
	}
	if (SelectContent=='') DisabledContentMenuStr=',ɾ��,����,����,';
	if ((top.MainInfo.SourceNews=='')&&(top.MainInfo.SourceDownLoad=='')) DisabledContentMenuStr=DisabledContentMenuStr+',ճ��,';
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
		if (CurrObj.NewsID!=null)
		{
			ListObjArray[ListObjArray.length]=new FolderFileObj(CurrObj,j,false);
			j++;
		}
	}
}
function SelectSpecialFile()
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
function DelSpecialNews()
{
	var SelectedSpecialNews='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.NewsID!=null)
			{
				if (SelectedSpecialNews=='') SelectedSpecialNews=ListObjArray[i].Obj.NewsID;
				else  SelectedSpecialNews=SelectedSpecialNews+'***'+ListObjArray[i].Obj.NewsID;
			}
		}
	}
	if (SelectedSpecialNews!='')
		OpenWindow('Frame.asp?FileName=SpecialDell.asp&Types=DellFile&SpecialID='+SpecialID+'&PageTitle=ɾ��ר������&NewsID='+SelectedSpecialNews,220,95,window);
	else alert('��ѡ��ר������');
}
function CutOperation()
{
	top.MainInfo.MoveTF=true;
	var SelectedSpecialNews='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.NewsID!=null)
			{
				if (SelectedSpecialNews=='') SelectedSpecialNews=ListObjArray[i].Obj.NewsID;
				else  SelectedSpecialNews=SelectedSpecialNews+'***'+ListObjArray[i].Obj.NewsID;
			}
		}
	}
	if (SelectedSpecialNews!='')
	{
		top.MainInfo.SourceClass=SpecialID;
		top.MainInfo.SourceNews=SelectedSpecialNews;
	}
}

function CopyOperation()
{
	top.MainInfo.MoveTF=false;
	var SelectedSpecialNews='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.NewsID!=null)
			{
				if (SelectedSpecialNews=='') SelectedSpecialNews=ListObjArray[i].Obj.NewsID;
				else  SelectedSpecialNews=SelectedSpecialNews+'***'+ListObjArray[i].Obj.NewsID;
			}
		}
	}
	if (SelectedSpecialNews!='')
	{
		top.MainInfo.SourceClass=SpecialID;
		top.MainInfo.SourceNews=SelectedSpecialNews;
	}
}

function PasteOperation()
{
	top.MainInfo.ObjectClass=SpecialID;
	var MoveOrCopyClassPara='MoveTF:'+top.MainInfo.MoveTF+',SourceClass:'+top.MainInfo.SourceClass+',SourceNews:'+top.MainInfo.SourceNews+',ObjectClass:'+SpecialID+',';
	OpenWindow('Frame.asp?FileName=SpecialNewsMoveOrCopy.asp&PageTitle=�ƶ�����ר������&MoveOrCopyClassPara='+MoveOrCopyClassPara,260,110,window);
}
function ChangePage(PageNum)
{
	window.location.href='?page_no='+PageNum+'&SpecialID='+SpecialID;
}
</script>
