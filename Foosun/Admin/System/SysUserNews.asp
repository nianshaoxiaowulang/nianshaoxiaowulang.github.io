<% Option Explicit %>
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Const.asp"-->
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
'if Not JudgePopedomTF(Session("Name"),"P040400") then Call ReturnError1()
Dim NewsAdminSql,RsUGObj
NewsAdminSql = "Select * from FS_MemberNews order by ID desc"
Set RsUGObj = Server.CreateObject(G_FS_RS)
RsUGObj.open NewsAdminSql,conn,1,1
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��Ա�б�</title>
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
          <td width=35 align="center" alt="��ӻ�Ա����" onClick="AddUserNews();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">�½�</td>
          <td width=2 class="Gray">|</td>
          <td width=35  align="center" alt="�޸Ļ�Ա����" onClick="EditUserNews();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">�޸�</td>
          <td width=2 class="Gray">|</td>
          <td width=35 align="center" alt="ɾ����Ա����" onClick="DelUserNews();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">ɾ��</td>
          <td width=2 class="Gray">|</td>
          <td width=35 align="center" alt="������Ա����" onClick="LockUserNews();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
          <td width=2 class="Gray">|</td>
          <td width=35 align="center" alt="������Ա����" onClick="UNLockUserNews();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
          <td width=2 class="Gray">|</td>
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
          <td width="33%" height="26" class="ButtonListLeft"> <div align="center">��������</div></td>
          <td width="13%" height="26" class="ButtonList"> <div align="center">���Ȩ��</div></td>
          <td width="24%" height="26" class="ButtonList"> <div align="center">����ʱ��</div></td>
          <td width="18%" height="26" class="ButtonList"> <div align="center">������</div></td>
          <td width="12%" height="26" class="ButtonList"> <div align="center">�Ƿ�����</div></td>
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
          <td height="26"><table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="16"><img  border="0" src="../Images/cmsv31_show.png" width="12" height="12"></td>
                <td width="306"><span class="TempletItem" Lock="<% = RsUGObj("isLock") %>" UserID="<% = RsUGObj("ID") %>"> 
                  <% = RsUGObj("Title") %>
                  </span></td>
              </tr>
            </table></td>
          <td><div align="center"> 
			<%
				If RsUGObj("PoPid")=0 Then
					Response.Write("������")
				Elseif RsUGObj("PoPid")=1 Then
					Response.Write("һ���Ա")
				Elseif RsUGObj("PoPid")=2 Then
					Response.Write("�м���Ա")
				Elseif RsUGObj("PoPid") = 3 Then
					Response.Write("�߼���Ա")
				Elseif RsUGObj("PoPid") = 4 Then
					Response.Write("VIP��Ա")
				Else	
					Response.Write("�������")
				End if
			%>
            </div></td>
          <td><div align="center"> 
              <% = RsUGObj("AddTime") %>
            </div></td>
          <td><div align="center"> 
			<% = RsUGObj("Author") %>
			</div></td>
          <td> <div align="center"> 
              <%
			  If RsUGObj("isLock")=0 then
					Response.Write("δ����")
			  Else
					Response.Write("<font color=red>������</font>")
			  End if
			  %>
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
					response.Write "&nbsp;<img src=../images/First1.gif border=0 alt=��ҳ></img>&nbsp;"
					response.Write "&nbsp;<img src=../images/prePage.gif border=0 alt=��һҳ></img>&nbsp;"
					response.Write "&nbsp;<img src=../images/nextPage.gif border=0 alt=��һҳ></img>&nbsp;"
					response.Write "&nbsp;<img src=../images/endPage.gif border=0 alt=βҳ></img>&nbsp;"
			else
				if cint(Page_No)<>1 and cint(Page_No)<>Page_Total then
					response.Write "&nbsp;<a href=?page_no=1&page_size="& page_size &"&Keywords="&Request("Keywords")&"><img src=../images/First1.gif border=0 alt=��ҳ></img></a>&nbsp;"
					response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)-1)&"&page_size="&page_size&"&Keywords="&Request("Keywords")&"><img src=../images/prePage.gif border=0 alt=��һҳ></img></a>&nbsp;"
					response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)+1)&"&page_size="&page_size&"&Keywords="&Request("Keywords")&"><img src=../images/nextPage.gif border=0 alt=��һҳ></img></a>&nbsp;"
					response.Write "&nbsp;<a href=?page_no="& Page_Total &"&page_size="& page_size &"&Keywords="&Request("Keywords")&"><img src=../images/endPage.gif border=0 alt=βҳ></img></a>&nbsp;"
				elseif cint(Page_No)=1 then
					response.Write "&nbsp;<img src=../images/First1.gif border=0 alt=��ҳ></img></a>&nbsp;"
					response.Write "&nbsp;<img src=../images/prePage.gif border=0 alt=��һҳ></img>&nbsp;"
					response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)+1)&"&page_size="& page_size &"&Keywords="&Request("Keywords")&"><img src=../images/nextPage.gif border=0 alt=��һҳ></img></a>&nbsp;"
					response.Write "&nbsp;<a href=?page_no="& Page_Total &"&page_size="& page_size &"&Keywords="&Request("Keywords")&"><img src=../images/endpage.gif border=0 alt=βҳ></img></a>&nbsp;"
				else
					response.Write "&nbsp;<a href=?page_no=1&page_size="& page_size &"&Keywords="&Request("Keywords")&"><img src=../images/First1.gif border=0 alt=��ҳ></img>&nbsp;"
					response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)-1)&"&page_size="& page_size &"&Keywords="&Request("Keywords")&"><img src=../images/prePage.gif border=0 alt=��һҳ></img></a>&nbsp;"
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
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.AddUserNews();",'�½�','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.EditUserNews();",'�޸�','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.DelUserNews();",'ɾ��','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('seperator','','');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('top.GetEkMainObject().location.reload();','ˢ��','');
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
	if (SelectUser=='') DisabledContentMenuStr=',�޸�,ɾ��,����,����,';
	else
	{
		if (SelectUser.indexOf('***')==-1) DisabledContentMenuStr='';
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
function AddUserNews()
{
	location='UserNewsAdd.asp';
}
function EditUserNews()
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
		if (SelectedUser.indexOf('***')==-1) location='UserNewsModify.asp?ID='+SelectedUser;
		else alert('һ��ֻ�ܹ��޸�һ������');
	}
	else alert('��ѡ�񹫸�');
}
function DelUserNews()
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
		 OpenWindow('Frame.asp?FileName=UserNewsDell.asp&PageTitle=ɾ������&OperateType=Dell&ID='+SelectedUser,220,105,window);
	else alert('��ѡ�񹫸�');
}
function LockUserNews()
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
		 OpenWindow('Frame.asp?FileName=UserNewsDell.asp&PageTitle=����&OperateType=isLock&ID='+SelectedUser,220,105,window);
	else alert('��ѡ�񹫸�');
}
function UNLockUserNews()
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
		 OpenWindow('Frame.asp?FileName=UserNewsDell.asp&PageTitle=����&OperateType=UnLock&ID='+SelectedUser,220,105,window);
	else alert('��ѡ�񹫸�');
}
</script>