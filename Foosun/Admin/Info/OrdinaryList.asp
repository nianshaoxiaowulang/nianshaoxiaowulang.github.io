<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<%
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
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

%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P070400") then Call ReturnError1()
Dim OperateType,Sql,RsRoutineObj
OperateType = Request("Type")
Sql = "Select * from FS_Routine where Type=" & OperateType
Set RsRoutineObj = Server.CreateObject(G_FS_RS)
RsRoutineObj.open Sql,conn,1,1
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�����б�</title>
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<script language="JavaScript" src="../../SysJS/ContentMenu.js"></script>
<body topmargin="2" leftmargin="2" onClick="SelectOrdinary();" ondragstart="return false;" onselectstart="return false;">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=35 align="center" alt="�½�" onClick="AddContent();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">�½�</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35  align="center" alt="�޸�" onClick="EditContent();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">�޸�</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="ɾ��" onClick="DelContent();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">ɾ��</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="����" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
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
  <td valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
	      <td width="50%" height="26" class="ButtonListLeft"> 
            <div align="center">����</div></td>
          <td width="50%" height="26" class="ButtonList"> 
            <div align="center">���ӵ�ַ</div></td>
  </tr>
  <%
if Not RsRoutineObj.eof then
	Dim page_size,page_no,page_total,record_all,i
	page_size=20
	page_no=request.querystring("page_no")
	if page_no <= 1 or page_no = "" then page_no=1
	If Request.QueryString("page_no")="" then
		page_no=1
	end if
	RsRoutineObj.PageSize=page_size
	page_total=RsRoutineObj.PageCount
	if (cint(page_no) > page_total) then page_no=page_total
	RsRoutineObj.AbsolutePage=page_no
	record_all=RsRoutineObj.RecordCount
	for i=1 to RsRoutineObj.PageSize
	if RsRoutineObj.eof then exit for
%>
  <tr> 
          <td><table border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><img src="../../Images/Common.gif" width="24" height="22"></td>
                <td><span class="TempletItem" OrdinaryID="<% = RsRoutineObj("ID") %>"><% = RsRoutineObj("Name") %></span></td>
              </tr>
            </table></td>
    <td><div align="center"><% = RsRoutineObj("Url") %></div></td>
  </tr>
  <%
	RsRoutineObj.MoveNext
Next
end if
%>
</table>
</td>
</tr>
<%

if page_total>1 then%>
<tr>
<td height="18">
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
 <tr> 
<td valign="middle" height="10">
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
						response.Write "&nbsp;<img src=../images/First1.gif border=0 alt=��ҳ></img>&nbsp;"
						response.Write "&nbsp;<img src=../images/prePage.gif border=0 alt=��һҳ></img>&nbsp;"
						response.Write "&nbsp;<img src=../images/nextPage.gif border=0 alt=��һҳ></img>&nbsp;"
						response.Write "&nbsp;<img src=../images/endPage.gif border=0 alt=βҳ></img>&nbsp;"
				else
					if cint(Page_No)<>1 and cint(Page_No)<>Page_Total then
						response.Write "&nbsp;<a href=?page_no=1&Type="& OperateType &"&Keywords="&Request("Keywords")&"><img src=../images/First1.gif border=0 alt=��ҳ></img></a>&nbsp;"
						response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)-1)&"&Type="& OperateType &"&Keywords="&Request("Keywords")&"><img src=../images/prePage.gif border=0 alt=��һҳ></img></a>&nbsp;"
						response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)+1)&"&Type="& OperateType &"&Keywords="&Request("Keywords")&"><img src=../images/nextPage.gif border=0 alt=��һҳ></img></a>&nbsp;"
						response.Write "&nbsp;<a href=?page_no="& Page_Total &"&Type="& OperateType &"&Keywords="&Request("Keywords")&"><img src=../images/endPage.gif border=0 alt=βҳ></img></a>&nbsp;"
					elseif cint(Page_No)=1 then
						response.Write "&nbsp;<img src=../images/First1.gif border=0 alt=��ҳ></img></a>&nbsp;"
						response.Write "&nbsp;<img src=../images/prePage.gif border=0 alt=��һҳ></img>&nbsp;"
						response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)+1)&"&Type="& OperateType &"&Keywords="&Request("Keywords")&"><img src=../images/nextPage.gif border=0 alt=��һҳ></img></a>&nbsp;"
						response.Write "&nbsp;<a href=?page_no="& Page_Total &"&Type="& OperateType &"&Keywords="&Request("Keywords")&"><img src=../images/endpage.gif border=0 alt=βҳ></img></a>&nbsp;"
					else
						response.Write "&nbsp;<a href=?page_no=1&Type="& OperateType &"&Keywords="&Request("Keywords")&"><img src=../images/First1.gif border=0 alt=��ҳ></img>&nbsp;"
						response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)-1)&"&Type="& OperateType & "&Keywords="&Request("Keywords")&"><img src=../images/prePage.gif border=0 alt=��һҳ></img></a>&nbsp;"
						response.Write "&nbsp;<img src=../images/nextPage.gif border=0 alt=��һҳ></img></a>&nbsp;"
						response.Write "&nbsp;<img src=../images/endpage.gif border=0 alt=βҳ></img>&nbsp;"
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
</table>
</body>
</html>
<%
Set Conn = Nothing
Set RsRoutineObj = Nothing
%>
<script language="JavaScript">
var OperateType='<% = OperateType %>';
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
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.EditContent();",'�޸�','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.DelContent();",'ɾ��','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('seperator','','');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('top.GetEkMainObject().location.reload();','ˢ��','');
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
			if (SelectContent=='') SelectContent=ListObjArray[i].Obj.OrdinaryID;
			else SelectContent=SelectContent+'***'+ListObjArray[i].Obj.OrdinaryID;
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
					if (SelectContent=='') SelectContent=ListObjArray[i].Obj.OrdinaryID;
					else SelectContent=SelectContent+'***'+ListObjArray[i].Obj.OrdinaryID;
				}
			}
		}
	}
	if (SelectContent=='') DisabledContentMenuStr=',�޸�,ɾ��,';
	else
	{
		if (SelectContent.indexOf('***')==-1) DisabledContentMenuStr='';
		else DisabledContentMenuStr=',�޸�,'
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
		if (CurrObj.OrdinaryID!=null)
		{
			ListObjArray[ListObjArray.length]=new FolderFileObj(CurrObj,j,false);
			j++;
		}
	}
}
function SelectOrdinary()
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
function AddContent()
{
	location='OrdinaryEdit.asp?OperateType='+OperateType;
}
function EditContent()
{
	var SelectedOrdinary='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.OrdinaryID!=null)
			{
				if (SelectedOrdinary=='') SelectedOrdinary=ListObjArray[i].Obj.OrdinaryID;
				else  SelectedOrdinary=SelectedOrdinary+'***'+ListObjArray[i].Obj.OrdinaryID;
			}
		}
	}
	if (SelectedOrdinary!='')
	{
		if (SelectedOrdinary.indexOf('***')==-1) location='OrdinaryEdit.asp?OperateType='+OperateType+'&OrdinaryID='+SelectedOrdinary;
		else alert('��ѡ��һ����������');
	}
	else alert('��ѡ���������');
}
function DelContent()
{
	var SelectedOrdinary='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.OrdinaryID!=null)
			{
				if (SelectedOrdinary=='') SelectedOrdinary=ListObjArray[i].Obj.OrdinaryID;
				else  SelectedOrdinary=SelectedOrdinary+'***'+ListObjArray[i].Obj.OrdinaryID;
			}
		}
	}
	if (SelectedOrdinary!='')
		OpenWindow('Frame.asp?FileName=OrdinaryDelete.asp&PageTitle=�������&OperateType='+OperateType+'&OrdinaryID='+SelectedOrdinary,220,95,window);
	else alert('��ѡ���������');
}
function ChangePage(PageNum)
{
	window.location.href='?page_no='+PageNum+'&Type='+ OperateType;
}
</script>