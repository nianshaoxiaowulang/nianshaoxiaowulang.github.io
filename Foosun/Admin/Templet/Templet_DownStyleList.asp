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
if Not JudgePopedomTF(Session("Name"),"P031000") then Call ReturnError1()
Dim StyleSql,RsStyleObj
StyleSql = "Select * from FS_DownListStyle Order By Id desc"
Set RsStyleObj = Conn.Execute(StyleSql)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�����б���ʽ�б�</title>
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<script language="JavaScript" src="../../SysJS/ContentMenu.js"></script>
<body topmargin="2" leftmargin="2" onClick="ClickStyle();" ondragstart="return false;" onselectstart="return false;">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=35 align="center" alt="�½�" onClick="AddStyle();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">�½�</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35  align="center" alt="�޸�" onClick="EditStyle();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">�޸�</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="ɾ��" onClick="DelStyle();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">ɾ��</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="�鿴" onClick="BrowStyle();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">�鿴</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="����" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  <td>&nbsp; </td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
<table width="100%" border="0" cellpadding="1" cellspacing="0">
  <tr>
  	<td colspan="5" height="2"></td>
  </tr>
  <tr> 
    <td height="26" class="ButtonListLeft"> 
      <div align="center">����</div></td>
    <td height="26" class="ButtonList"> 
      <div align="center">���</div></td>
  </tr>
  <%
do while Not RsStyleObj.Eof
%>
  <tr> 
    <td> 
      <table border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><img src="../../Images/Folder/Log.gif" width="18" height="18"></td>
          <td><span class="TempletItem" StyleID="<% = RsStyleObj("ID") %>"><% = RsStyleObj("Name") %></span></td>
        </tr>
      </table>
    </td>
    <td> 
      <div align="center"> 
        <% = RsStyleObj("ID") %>
      </div></td>
  </tr>
  <%
	RsStyleObj.MoveNext
loop
%>
</table>
</body>
</html>
<%
Set RsStyleObj = Nothing
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
	InitialContentListContentMenu();
	DocumentReadyTF=true;
}
function InitialContentListContentMenu()
{
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.EditStyle();",'�޸�','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.DelStyle();",'ɾ��','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.BrowStyle();','�鿴','disabled');
	//ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('seperator','','');
	//ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('top.GetEkMainObject().location.reload();','ˢ��','');
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
			if (SelectContent=='') SelectContent=ListObjArray[i].Obj.StyleID;
			else SelectContent=SelectContent+'***'+ListObjArray[i].Obj.StyleID;
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
					if (SelectContent=='') SelectContent=ListObjArray[i].Obj.StyleID;
					else SelectContent=SelectContent+'***'+ListObjArray[i].Obj.StyleID;
				}
			}
		}
	}
	if (SelectContent=='') DisabledContentMenuStr=',�޸�,ɾ��,�鿴,';
	else
	{
		if (SelectContent.indexOf('***')==-1) DisabledContentMenuStr='';
		else DisabledContentMenuStr=',�޸�,�鿴,'
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
		if (CurrObj.StyleID!=null)
		{
			ListObjArray[ListObjArray.length]=new FolderFileObj(CurrObj,j,false);
			j++;
		}
	}
}
function ClickStyle()
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
function AddStyle()
{
	location='Templet_DownStyleAdd.asp';
}
function EditStyle()
{
	var SelectedStyle='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.StyleID!=null)
			{
				if (SelectedStyle=='') SelectedStyle=ListObjArray[i].Obj.StyleID;
				else  SelectedStyle=SelectedStyle+'***'+ListObjArray[i].Obj.StyleID;
			}
		}
	}
	if (SelectedStyle!='')
	{
		if (SelectedStyle.indexOf('***')==-1)
			location='Templet_DownStyleAdd.asp?ID='+SelectedStyle;
		else alert('һ��ֻ�ܹ��޸�һ����ʽ');
	}
	else alert('��ѡ��Ҫ�޸���ʽ');
}
function DelStyle()
{
	var SelectedStyle='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.StyleID!=null)
			{
				if (SelectedStyle=='') SelectedStyle=ListObjArray[i].Obj.StyleID;
				else  SelectedStyle=SelectedStyle+'***'+ListObjArray[i].Obj.StyleID;
			}
		}
	}
	if (SelectedStyle!='')
		OpenWindow('Frame.asp?FileName=Templet_DownStyleDel.asp&PageTitle=ɾ�������б���ʽ&ID='+SelectedStyle,200,120,window);
	else alert('��ѡ��Ҫɾ������ʽ');
}
function BrowStyle()
{
	var SelectedStyle='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.StyleID!=null)
			{
				if (SelectedStyle=='') SelectedStyle=ListObjArray[i].Obj.StyleID;
				else  SelectedStyle=SelectedStyle+'***'+ListObjArray[i].Obj.StyleID;
			}
		}
	}
	if (SelectedStyle!='')
	{
		if (SelectedStyle.indexOf('***')==-1)
			OpenWindow('Frame.asp?FileName=Templet_DownStyleBrow.asp&PageTitle=�鿴��ʽ&ID='+SelectedStyle,360,190,window);
		else alert('һ��ֻ�ܹ��鿴һ����ʽ');
	}
	else alert('��ѡ��Ҫ�鿴��ʽ');
}
</script>