<% Option Explicit %>
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="inc/Config.asp" -->
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
Dim SiteID
SiteID = Request("SiteID")
Dim DBC,Conn,CollectConn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
DBC.ConnStr = CollectDBConnectionStr
Set CollectConn = DBC.OpenConnection()
Set DBC = Nothing
'�ж�Ȩ��
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P080400") then Call ReturnError1()
'�ж�Ȩ�޽���
Dim Action
Action = Request("Action")
if Action = "DelAll" then
	if Not JudgePopedomTF(Session("Name"),"P080400") then Call ReturnError1()
	CollectConn.Execute("Delete from FS_News where History=1")
end if
Dim NewsSql,RsNewsObj,CurrPage,AllPageNum,RecordNum,i,SysClassCName,SiteName,RsTempObj
CurrPage = Request("CurrPage")
NewsSql = "Select * from FS_News where History=1 Order by ID Desc"
Set RsNewsObj = Server.CreateObject("ADODB.RecordSet")
RsNewsObj.Open NewsSql,CollectConn,1,1
%>
<HTML>
<HEAD>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>���Ųɼ�</TITLE>
</HEAD>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script language="JavaScript" src="../../SysJS/PublicJS.js"></script>
<script language="JavaScript" src="../../SysJS/ContentMenu.js"></script>
<BODY topmargin="2" leftmargin="2" onClick="SelectNews();" onselectstart="return false;">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="35" align="center" alt="ɾ��" onClick="DelNews();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">ɾ��</td>
			<td width=2 class="Gray">|</td>
          <td width="70" align="center" alt="ɾ��ȫ��" onClick="DelAll();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">ɾ��ȫ��</td>
			<td width=2 class="Gray">|</td>
		  <td width="35" align="center" alt="���" onClick="MoveNewsToSystem();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">���</td>
			<td width=2 class="Gray">|</td>
		  <td width="35" align="center" alt="����" onClick="history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
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
    <td height="26" nowrap class="ButtonListLeft"> <div align="center">����</div></td>
    <td width="15%" height="24" nowrap class="ButtonList"> 
      <div align="center">���ų���</div></td>
    <td width="15%" height="24" nowrap class="ButtonList"> 
      <div align="center">Ŀ����Ŀ</div></td>
    <td width="15%" height="24" nowrap class="ButtonList"> 
      <div align="center">�ɼ�վ��</div></td>
    <td width="15%" height="24" nowrap class="ButtonList"> 
      <div align="center">�������</div></td>
  </tr>
  <%
if Not RsNewsObj.Eof then
	if CurrPage = "" then
		CurrPage = 1
	else
		CurrPage = CInt(CurrPage)
	end if
	RsNewsObj.PageSize = 18
	RecordNum = RsNewsObj.RecordCount
	AllPageNum = RsNewsObj.PageCount
	if CurrPage > AllPageNum then CurrPage = AllPageNum
	RsNewsObj.AbsolutePage = Cint(CurrPage)
	for i = 1 to RsNewsObj.PageSize
		if RsNewsObj.Eof then Exit For
		Set RsTempObj = Conn.Execute("Select ClassCName from FS_NewsClass where ClassID='" & RsNewsObj("ClassID") & "'")
		if Not RsTempObj.Eof then
			SysClassCName = RsTempObj("ClassCName")
		else
			SysClassCName = "��Ŀ������"
		end if
		RsTempObj.Close
		Set RsTempObj = Nothing
		Set RsTempObj = CollectConn.Execute("Select SiteName from FS_Site where ID=" & RsNewsObj("SiteID"))
		if Not RsTempObj.Eof then
			SiteName = RsTempObj("SiteName")
		else
			SiteName = "δ֪"
		end if
		RsTempObj.Close
		Set RsTempObj = Nothing
%>
  <tr> 
    <td height="26" nowrap>
<table border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><img src="../../Images/Info/WordNews.gif" width="24" height="22"></td>
          <td><span class="TempletItem" NewsID=<% = RsNewsObj("ID") %>>
<% = Left(RsNewsObj("Title"),20) %></span></td>
        </tr>
      </table></td>
    <td nowrap><div align="center"> 
        <% = Len(RsNewsObj("Content")) %>
        �ַ�</div></td>
    <td nowrap><div align="center"> 
        <% = SysClassCName %>
      </div></td>
    <td nowrap><div align="center"> 
        <% = SiteName %>
      </div></td>
    <td nowrap><div align="center"> 
        <% = RsNewsObj("AddDate") %>
      </div></td>
  </tr>
  <%
		RsNewsObj.MoveNext
	next
%>
  <tr> 
    <td height="30" colspan="5" nowrap><table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td> <div align="right"> 
              <%
			Response.Write"&nbsp;��<b>"& AllPageNum & "</b>ҳ<b>" & RecordNum & "</b>����¼��ÿҳ<b>" & RsNewsObj.pagesize & "</b>������ҳ�ǵ�<b>"& CurrPage &"</b>ҳ"
			if Int(CurrPage) > 1 then
				Response.Write"&nbsp;<a href=?CurrPage=1>��ҳ</a>&nbsp;"
				Response.Write"&nbsp;<a href=?CurrPage=" & Cstr(CInt(CurrPage)-1) & ">��ҳ</a>&nbsp;"
			end if
			if Int(CurrPage) < AllPageNum then
				Response.Write"&nbsp;<a href=?CurrPage=" & Cstr(Cint(CurrPage)+1) & ">��ҳ</a>"
				Response.Write"&nbsp;<a href=?CurrPage=" & AllPageNum & ">ĩҳ</a>&nbsp;"
			end if
			Response.Write"<br>"
		%>
            </div></td>
        </tr>
      </table></td>
  </tr>
  <%
end if
%>
</table>
</BODY>
</HTML>
<%
Set CollectConn = Nothing
Set Conn = Nothing
Set RsNewsObj = Nothing
%>
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
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.DelNews();",'ɾ��','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.MoveNewsToSystem();','���','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('seperator','','');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('location.reload();','ˢ��','');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('prompt(\'��ҳ��·������\',\'<%=Request.ServerVariables("SCRIPT_NAME")%>\');','·������','');
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
	if (SelectContent=='') DisabledContentMenuStr=',ɾ��,���,';
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
		if (CurrObj.NewsID!=null)
		{
			ListObjArray[ListObjArray.length]=new FolderFileObj(CurrObj,j,false);
			j++;
		}
	}
}
function SelectNews()
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
function EditNews()
{
	var SelectedNews='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.NewsID!=null)
			{
				if (SelectedNews=='') SelectedNews=ListObjArray[i].Obj.NewsID;
				else  SelectedNews=SelectedNews+'***'+ListObjArray[i].Obj.NewsID;
			}
		}
	}
	if (SelectedNews!='')
	{
		if (SelectedNews.indexOf('***')==-1) window.location='EditNews.asp?NewsIDStr='+SelectedNews;
		else alert('��ѡ��һ������');
	}
	else alert('��ѡ������');
}
function DelNews()
{
	var SelectedNews='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.NewsID!=null)
			{
				if (SelectedNews=='') SelectedNews=ListObjArray[i].Obj.NewsID;
				else  SelectedNews=SelectedNews+'***'+ListObjArray[i].Obj.NewsID;
			}
		}
	}
	if (SelectedNews!='')
		OpenWindow('Frame.asp?FileName=DelNews.asp&PageTitle=ɾ������&NewsIDStr='+SelectedNews,200,120,window);
	else alert('��ѡ������');
}
function DelAll()
{
	if (confirm('ȷ��Ҫɾ����')) location='?Action=DelAll'
}
function MoveNewsToSystem()
{
	var SelectedNews='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.NewsID!=null)
			{
				if (SelectedNews=='') SelectedNews=ListObjArray[i].Obj.NewsID;
				else  SelectedNews=SelectedNews+'***'+ListObjArray[i].Obj.NewsID;
			}
		}
	}
	if (SelectedNews!='')
		OpenWindow('Frame.asp?FileName=MoveNewsToSystem.asp&PageTitle=�������&DelNews=true&NewsIDStr='+SelectedNews,200,120,window);
	else alert('��ѡ������');
}
</script>