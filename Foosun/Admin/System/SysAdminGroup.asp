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
if Not JudgePopedomTF(Session("Name"),"P040100") then Call ReturnError1()
Dim AdminGroupSql,RsAdminGroupObj,CheckInfo,GroupMen,RsTempObj
CheckInfo = ""
AdminGroupSql = "Select * from FS_AdminGroup"
Set RsAdminGroupObj = Server.CreateObject(G_FS_RS)
RsAdminGroupObj.open AdminGroupSql,Conn,1,1
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>ϵͳ����Ա���б�</title>
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<script language="JavaScript" src="../../SysJS/ContentMenu.js"></script>
<body topmargin="2" leftmargin="2" onClick="SelectAdminGroup();" onselectstart="return false;">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle" bgcolor="#EEEEEE"> 
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=34 align="center" alt="�½�����Ա��" onClick="AddAdminGroup();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">�½�</td>
		  <td width=6 class="Gray">|</td>
		  <td width=34  align="center" alt="�޸Ĺ���Ա��" onClick="EditAdminGroup();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">�޸�</td>
		  <td width=6 class="Gray">|</td>
		  <td width=34 align="center" alt="ɾ������Ա��" onClick="DelAdminGroup();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">ɾ��</td>
		  <td width=6 class="Gray">|</td>
		  <td width=62 align="center" alt="����Ȩ��" onClick="SetAdminGroupPop();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����Ȩ��</td>
		  <td width=10 class="Gray">|</td>
		  <td width=37 align="center" alt="����" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  <td width="760">&nbsp; </td>
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
          <td width="20%" height="26" class="ButtonListLeft"> 
            <div align="center">����Ա��</div></td>
          <td width="40%" height="26" class="ButtonList"> 
            <div align="center">���Ա</div></td>
          <td height="26" class="ButtonList"> 
            <div align="center">���</div></td>
        </tr>
        <% 
if Not RsAdminGroupObj.eof then
	Dim Temp_id
	Temp_id = 1
	 Dim page_size,page_no,page_total,record_all,PageNums
	page_size=Request.QueryString("page_size")
	if page_size<=0 or page_size="" then page_size=18
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
	RsAdminGroupObj.PageSize=page_size
	page_total=RsAdminGroupObj.PageCount
	if (cint(page_no)>page_total) then page_no=page_total
	RsAdminGroupObj.AbsolutePage=page_no
	record_all=RsAdminGroupObj.RecordCount
	Dim i
	for i=1 to RsAdminGroupObj.PageSize
	if RsAdminGroupObj.eof then exit for
		GroupMen = ""
		Set RsTempobj = Conn.Execute("Select Name from FS_Admin where GroupID=" & RsAdminGroupObj("ID"))
		if Not RsTempObj.Eof then
			do while Not RsTempObj.Eof
				if GroupMen = "" then
					GroupMen = RsTempobj("Name")
				else
					GroupMen = GroupMen & "|" & RsTempobj("Name")
				end if
				RsTempObj.MoveNext
			loop
		else
			GroupMen = "���޳�Ա"
		end if
		RsTempObj.Close
		Set RsTempObj = Nothing
%>
        <tr> 
          <td height="20"><table border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><img src="../../Images/Folder/UserfaceGroup.gif" width="18" height="18"></td>
                <td><span class="TempletItem" AdminGroupID="<% = RsAdminGroupObj("ID") %>"> 
            <% = RsAdminGroupObj("GroupName") %>
            </span></td>
              </tr>
            </table></td>
          <td height="20" class="TempletItem"> 
            <div align="center"> 
              <% = GroupMen %>
            </div></td>
          <td height="20" class="TempletItem"  title="<%=RsAdminGroupObj("Comment")%>"> 
            <div align="center">
              <% = Left(RsAdminGroupObj("Comment"),15) %>
             </div></td>
        </tr>
        <%
  		Temp_id = Temp_id + 1
		RsAdminGroupObj.MoveNext
	Next
end if
%>
      </table>
</td>
</tr>
<%if page_total>1 then%>
<tr height="18">
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
					response.Write "&nbsp;<a href=?page_no=1&page_size="& page_size &"&Keywords="&Request("Keywords")&"><img src=../../images/FirstPage.gif border=0 alt=��ҳ></img></a>&nbsp;"
					response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)-1)&"&page_size="&page_size&"&Keywords="&Request("Keywords")&"><img src=../../images/prePage.gif border=0 alt=��һҳ></img></a>&nbsp;"
					response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)+1)&"&page_size="&page_size&"&Keywords="&Request("Keywords")&"><img src=../../images/nextPage.gif border=0 alt=��һҳ></img></a>&nbsp;"
					response.Write "&nbsp;<a href=?page_no="& Page_Total &"&page_size="& page_size &"&Keywords="&Request("Keywords")&"><img src=../../images/endPage.gif border=0 alt=βҳ></img></a>&nbsp;"
				elseif cint(Page_No)=1 then
					response.Write "&nbsp;<img src=../../images/FirstPage.gif border=0 alt=��ҳ></img></a>&nbsp;"
					response.Write "&nbsp;<img src=../../images/prePage.gif border=0 alt=��һҳ></img>&nbsp;"
					response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)+1)&"&page_size="& page_size &"&Keywords="&Request("Keywords")&"><img src=../../images/nextPage.gif border=0 alt=��һҳ></img></a>&nbsp;"
					response.Write "&nbsp;<a href=?page_no="& Page_Total &"&page_size="& page_size &"&Keywords="&Request("Keywords")&"><img src=../../images/endpage.gif border=0 alt=βҳ></img></a>&nbsp;"
				else
					response.Write "&nbsp;<a href=?page_no=1&page_size="& page_size &"&Keywords="&Request("Keywords")&"><img src=../../images/FirstPage.gif border=0 alt=��ҳ></img>&nbsp;"
					response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)-1)&"&page_size="& page_size &"&Keywords="&Request("Keywords")&"><img src=../../images/prePage.gif border=0 alt=��һҳ></img></a>&nbsp;"
					response.Write "&nbsp;<img src=../../images/nextPage.gif border=0 alt=��һҳ></img></a>&nbsp;"
					response.Write "&nbsp;<img src=../../images/endpage.gif border=0 alt=βҳ></img>&nbsp;"
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
<script language="JavaScript">
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
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.AddAdminGroup();",'�½�','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.EditAdminGroup();",'�޸�','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.DelAdminGroup();",'ɾ��','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.SetAdminGroupPop();','Ȩ��','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('seperator','','');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('top.GetEkMainObject().location.reload();','ˢ��','');
	//ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('prompt(\'��ҳ��·������\',\'<%=Request.ServerVariables("SCRIPT_NAME")%>\');','·������','');
}
function ContentMenuFunction(ExeFunction,Description,EnabledStr)
{
	this.ExeFunction=ExeFunction;
	this.Description=Description;
	this.EnabledStr=EnabledStr;
}
function ContentMenuShowEvent()
{
	ChangeAdminMenuStatus();
}
function ChangeAdminMenuStatus()
{
	var EventObjInArray=false,SelectAdmin='',DisabledContentMenuStr='';
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
			if (SelectAdmin=='') SelectAdmin=ListObjArray[i].Obj.AdminGroupID;
			else SelectAdmin=SelectAdmin+'***'+ListObjArray[i].Obj.AdminGroupID;
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
					if (SelectAdmin=='') SelectAdmin=ListObjArray[i].Obj.AdminGroupID;
					else SelectAdmin=SelectAdmin+'***'+ListObjArray[i].Obj.AdminGroupID;
				}
			}
		}
	}
	if (SelectAdmin=='') DisabledContentMenuStr=',�޸�,ɾ��,Ȩ��,';
	else
	{
		if (SelectAdmin.indexOf('***')==-1) DisabledContentMenuStr='';
		else DisabledContentMenuStr=',�޸�,Ȩ��,'
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
		if (CurrObj.AdminGroupID!=null)
		{
			ListObjArray[ListObjArray.length]=new FolderFileObj(CurrObj,j,false);
			j++;
		}
	}
}
function SelectAdminGroup()
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
function AddAdminGroup()
{
	location='AdminGroupEdit.asp';
}
function EditAdminGroup()
{
	var SelectedAdminGroup='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.AdminGroupID!=null)
			{
				if (SelectedAdminGroup=='') SelectedAdminGroup=ListObjArray[i].Obj.AdminGroupID;
				else  SelectedAdminGroup=SelectedAdminGroup+'***'+ListObjArray[i].Obj.AdminGroupID;
			}
		}
	}
	if (SelectedAdminGroup!='')
	{
		if (SelectedAdminGroup.indexOf('***')==-1) location='AdminGroupEdit.asp?ID='+SelectedAdminGroup;
		else alert('һ��ֻ�ܹ��޸�һ������Ա��');
	}
	else alert('��ѡ��Ҫ�޸ĵĹ���Ա��');
}
function DelAdminGroup()
{
	var SelectedAdminGroup='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.AdminGroupID!=null)
			{
				if (SelectedAdminGroup=='') SelectedAdminGroup=ListObjArray[i].Obj.AdminGroupID;
				else  SelectedAdminGroup=SelectedAdminGroup+'***'+ListObjArray[i].Obj.AdminGroupID;
			}
		}
	}
	if (SelectedAdminGroup!='')
		OpenWindow('Frame.asp?FileName=SysAdminGroup_Del.asp&PageTitle=ɾ������Ա��&ID='+SelectedAdminGroup,230,110,window);
	else alert('��ѡ��Ҫ�޸ĵĹ���Ա��');
}
function SetAdminGroupPop()
{
	var SelectedAdminGroup='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.AdminGroupID!=null)
			{
				if (SelectedAdminGroup=='') SelectedAdminGroup=ListObjArray[i].Obj.AdminGroupID;
				else  SelectedAdminGroup=SelectedAdminGroup+'***'+ListObjArray[i].Obj.AdminGroupID;
			}
		}
	}
	if (SelectedAdminGroup!='')
	{
		if (SelectedAdminGroup.indexOf('***')==-1) location='SetAdminGroupPop.asp?ID='+SelectedAdminGroup;
		else alert('һ��ֻ�ܹ�����һ������Ա���Ȩ��');
	}
	else alert('��ѡ�����Ա��');
}
function ChangePage(PageNum)
{
	var page_size=<% = page_size %>
	window.location.href='?page_no='+PageNum+'&page_size='+page_size;
}
</script>