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
if Not JudgePopedomTF(Session("Name"),"P080600") then Call ReturnError1()
Dim RsPlusObj,SpecialPicStr,TempObj,FileNum,TempShowNavi
Set RsPlusObj = Conn.Execute("Select * from FS_Plus order by ID desc")
%>
<html>
<head>
<link href="../../../CSS/FS_css.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>����������б�</title>
</head>
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<script language="JavaScript" src="../../SysJS/ContentMenu.js"></script>
<body topmargin="2" onclick="SelectPlus();" leftmargin="2" ondragstart="return false;" onselectstart="return false;">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=35 align="center" alt="�½�" onClick="AddPlus();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">�½�</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35  align="center" alt="�޸�" onClick="EditPlus();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">�޸�</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="ɾ��" onClick="DelPlus();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">ɾ��</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="��ʾ" onClick="OpenPlus();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">��ʾ</td>
		  <td width=2 class="Gray">|</td>
          <td width=35 align="center" alt="����" onClick="ClosePlus();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
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
    <td valign="top"> <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="28%" height="26" class="ButtonListLeft"> 
            <div align="center">�������</div></td>
          <td width="34%" height="20" class="ButtonList">
<div align="center">���ӵ�ַ</div></td>
          <td width="9%" height="20" class="ButtonList">
<div align="center">�򿪷�ʽ</div></td>
          <td width="10%" height="20" class="ButtonList">
<div align="center">��ʾ״̬</div></td>
          <td width="19%" height="20" class="ButtonList">
<div align="center">���ʱ��</div></td>
        </tr>
<%
	do while Not RsPlusObj.Eof 
	Dim  OpenTypes,StateTypes 
	If RsPlusObj("OpenType")="1" then
		OpenTypes = "�´���"
	Else
		OpenTypes = "ԭ����"
	End If
	If RsPlusObj("ShowTF")="1" then
	   StateTypes = "��ʾ"
	Else
	   StateTypes = "����"
	End If
%>
        <tr> 
          <td><table border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><img src="../../Images/Folder/folderclosed.gif"></td>
                <td><span class="TempletItem" PlusID="<%=RsPlusObj("ID")%>" States="<%=RsPlusObj("ShowTF")%>"><%=RsPlusObj("Name")%></span></td>
              </tr>
            </table></td>
          <td height="20"><div align="center"><a href="PlusRedirect.asp?id=<%=RsPlusObj("id")%>" target="<%if RsPlusObj("OpenType")="1" then Response.Write("_New") else Response.Write("_self") end if%>">ת���ַ</a></div></td>
          <td height="20"><div align="center"><%=OpenTypes%></div></td>
          <td height="20"><div align="center"><%=StateTypes%></div></td>
          <td height="20"><div align="center"><%=RsPlusObj("AddTime")%></div></td>
        </tr>
        <%
		RsPlusObj.MoveNext
	loop
	RsPlusObj.close
	set RsPlusObj=nothing
%>
      </table></td>
  </tr>
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
	InitialContentListContentMenu();
	DocumentReadyTF=true;
}
function InitialContentListContentMenu()
{
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.EditPlus();",'�޸�','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.DelPlus();",'ɾ��','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('seperator','','');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.OpenPlus();','��ʾ','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.ClosePlus();','����','disabled');
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
	var EventObjInArray=false,SelectPlus='',DisabledContentMenuStr='';
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
			if (SelectPlus=='') SelectPlus=ListObjArray[i].Obj.PlusID;
			else SelectPlus=SelectPlus+'***'+ListObjArray[i].Obj.PlusID;
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
					if (SelectPlus=='') SelectPlus=ListObjArray[i].Obj.PlusID;
					else SelectPlus=SelectPlus+'***'+ListObjArray[i].Obj.PlusID;
				}
			}
		}
	}
	if (SelectPlus=='') DisabledContentMenuStr=',�޸�,ɾ��,��ʾ,����,';
	else
	{
		if (SelectPlus.indexOf('***')==-1) DisabledContentMenuStr='';
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
		if (CurrObj.PlusID!=null)
		{
			ListObjArray[ListObjArray.length]=new FolderFileObj(CurrObj,j,false);
			j++;
		}
	}
}
function SelectPlus()
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
function AddPlus()
{
	location='PlusAdd.asp';
}
function EditPlus()
{
	var SelectedPlus='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.PlusID!=null)
			{
				if (SelectedPlus=='') SelectedPlus=ListObjArray[i].Obj.PlusID;
				else  SelectedPlus=SelectedPlus+'***'+ListObjArray[i].Obj.PlusID;
			}
		}
	}
	if (SelectedPlus!='')
	{
		if (SelectedPlus.indexOf('***')==-1) location='PlusModify.asp?PlusID='+SelectedPlus;
		else alert('��ѡ��һ�����');
	}
	else alert('��ѡ����');
}
function DelPlus()
{
	var SelectedPlus='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.PlusID!=null)
			{
				if (SelectedPlus=='') SelectedPlus=ListObjArray[i].Obj.PlusID;
				else  SelectedPlus=SelectedPlus+'***'+ListObjArray[i].Obj.PlusID;
			}
		}
	}
	if (SelectedPlus!='')
		OpenWindow('Frame.asp?FileName=PlusDell.asp&PageTitle=ɾ�����&PlusID='+SelectedPlus,220,105,window);
	else alert('��ѡ����');
}
function OpenPlus()
{
	var SelectedPlus='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.PlusID!=null)
			{
				if (SelectedPlus=='') SelectedPlus=ListObjArray[i].Obj.PlusID;
				else  SelectedPlus=SelectedPlus+'***'+ListObjArray[i].Obj.PlusID;
			}
		}
	}
	if (SelectedPlus!='')
		OpenWindow('Frame.asp?FileName=PlusDell.asp&PageTitle=��ʾ���&Types=Shows&PlusID='+SelectedPlus,220,105,window);
	else alert('��ѡ����');
}
function ClosePlus()
{
	var SelectedPlus='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.PlusID!=null)
			{
				if (SelectedPlus=='') SelectedPlus=ListObjArray[i].Obj.PlusID;
				else  SelectedPlus=SelectedPlus+'***'+ListObjArray[i].Obj.PlusID;
			}
		}
	}
	if (SelectedPlus!='')
		OpenWindow('Frame.asp?FileName=PlusDell.asp&PageTitle=���ز��&Types=Hide&PlusID='+SelectedPlus,220,105,window);
	else alert('��ѡ����');
}
</script>
