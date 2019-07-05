<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<%
	Dim DBC,Conn
	On Error Resume Next
	Set DBC = New DataBaseClass
	Set Conn = DBC.OpenConnection()
	Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P030800") then Call ReturnError1()
Dim LableSql,RsLableObj,RsLblTypeObj,TypeID,LableID
TypeID=request("TypeID")
if TypeID<0 or TypeID="" then TypeID=0
LableSql = "Select * from FS_Lable where Type="&TypeID&" Order By Id Desc"
Set RsLableObj = Server.CreateObject(G_FS_RS)
RSLableObj.open LableSql,conn,1,1

LableSql="Select * from FS_LableType where ParentID="&TypeID
Set RsLblTypeObj = Server.CreateObject(G_FS_RS)
RsLblTypeObj.open LableSql,conn,1,1
dim UPTypeObj,SQLStr,TempID
set UPTypeObj=Server.CreateObject(G_FS_RS)
SQLStr="select * from FS_LableType where ID ="&TypeID
UPTypeObj.open SQLStr,conn,1,1
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��ǩ�б�</title>
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script language="JavaScript" src="../../SysJS/PublicJS.js"></script>
<script language="JavaScript" src="../../SysJS/ContentMenu.js"></script>
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
function FolderFileObj(Obj,Index,Selected)
{
	this.Obj=Obj;
	this.Index=Index;
	this.Selected=Selected;
}
function ContentMenuFunction(ExeFunction,Description,EnabledStr)
{
	this.ExeFunction=ExeFunction;
	this.Description=Description;
	this.EnabledStr=EnabledStr;
}
function InitialClassListContentMenu()
{
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.AddLable();",'�½���ǩ','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.EditFolderOrLable();",'�޸�','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.DelFolderAndLable();",'ɾ��','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.BackUpLable();','����','disabled');
}
function ContentMenuShowEvent()
{
	ChangeLableMenuStatus();
}
function RefreshList()
{
	location.href=location.href;
}
function ChangeLableMenuStatus()
{
	var EventObjInArray=false,SelectFolder='',SelectFile='',DisabledContentMenuStr='';
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
			if (ListObjArray[i].Obj.TypeID!=null)
			{
				if (SelectFolder=='') SelectFolder=ListObjArray[i].Obj.TypeID;
				else SelectFolder=SelectFolder+'***'+ListObjArray[i].Obj.TypeID
			}
			if (ListObjArray[i].Obj.LableID!=null)
			{
				if (SelectFile=='') SelectFile=ListObjArray[i].Obj.LableID;
				else SelectFile=SelectFile+'***'+ListObjArray[i].Obj.LableID
			}
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
					if (ListObjArray[i].Obj.TypeID!=null)
					{
						if (SelectFolder=='') SelectFolder=ListObjArray[i].Obj.TypeID;
						else SelectFolder=SelectFolder+'***'+ListObjArray[i].Obj.TypeID
					}
					if (ListObjArray[i].Obj.LableID!=null)
					{
						if (SelectFile=='') SelectFile=ListObjArray[i].Obj.LableID;
						else SelectFile=SelectFile+'***'+ListObjArray[i].Obj.LableID
					}
				}
			}
		}
	}
	if ((SelectFolder=='')&&(SelectFile=='')) DisabledContentMenuStr=',�޸�,ɾ��,����,';
	else
	{
		if ((SelectFile!='')&&(SelectFolder==''))
		{
			if (SelectFile.indexOf('***')!=-1) DisabledContentMenuStr=',�޸�,';
			else DisabledContentMenuStr='';
		}
		if ((SelectFolder!='')&&(SelectFile!='')) DisabledContentMenuStr=DisabledContentMenuStr+',�޸�,����,';
		if ((SelectFolder!='')&&(SelectFile==''))
		{
			if (SelectFolder.indexOf('***')!=-1) DisabledContentMenuStr=',�޸�,����,';
			else DisabledContentMenuStr=',����,';
		}
	}
	for (var i=0;i<ContentMenuArray.length;i++)
	{
		if (DisabledContentMenuStr.indexOf(ContentMenuArray[i].Description)!=-1) ContentMenuArray[i].EnabledStr='disabled';
		else  ContentMenuArray[i].EnabledStr='';
	}
}
function IntialListObjArray()
{
	var CurrObj=null,j=1;
	for (var i=0;i<document.all.length;i++)
	{
		CurrObj=document.all(i);
		if ((CurrObj.TypeID!=null)||(CurrObj.LableID!=null))
		{
			ListObjArray[ListObjArray.length]=new FolderFileObj(CurrObj,j,false);
			j++;
		}
	}
}
function ChangeLable(Obj)
{
	location.href='Templet_LableList.asp?TypeID='+Obj.TypeID;
}
function AddLableFolder()
{
	location='LableTypeAddNew.asp?BigTypeID='+BigTypeID;
}
function AddLable()
{
	location='LableAddNew.asp?BigTypeID='+BigTypeID;
}
function EditFolderOrLable()
{
	var SelectedLable='',SelectedType='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.TypeID!=null)
			{
				if (SelectedType=='') SelectedType=ListObjArray[i].Obj.TypeID;
				else  SelectedType=SelectedType+'***'+ListObjArray[i].Obj.TypeID;
			}
			if (ListObjArray[i].Obj.LableID!=null)
			{
				if (SelectedLable=='') SelectedLable=ListObjArray[i].Obj.LableID;
				else  SelectedLable=SelectedLable+'***'+ListObjArray[i].Obj.LableID;
			}
		}
	}
	if (!((SelectedLable=='')&&(SelectedType=='')))
	{
		if (SelectedLable!='')
		{
			if (SelectedLable.indexOf('***')==-1) location='LableAddNew.asp?ID='+SelectedLable+'&BigTypeID='+BigTypeID;
			else alert('һ��ֻ�ܹ��༭һ����ǩ');
		}
		if (SelectedType!='')
		{
			if (SelectedType.indexOf('***')==-1) location='LableTypeAddNew.asp?BigTypeID='+BigTypeID+'&ID='+SelectedType;
			else alert('һ��ֻ�ܹ��༭һ����');
		}
	}
	else alert('��ѡ��Ҫ�༭�ı�ǩ');
}
function DelFolderAndLable()
{
	var SelectedLable='',SelectedType='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.TypeID!=null)
			{
				if (SelectedType=='') SelectedType=ListObjArray[i].Obj.TypeID;
				else  SelectedType=SelectedType+'***'+ListObjArray[i].Obj.TypeID;
			}
			if (ListObjArray[i].Obj.LableID!=null)
			{
				if (SelectedLable=='') SelectedLable=ListObjArray[i].Obj.LableID;
				else  SelectedLable=SelectedLable+'***'+ListObjArray[i].Obj.LableID;
			}
		}
	}
	if ((SelectedLable!='')||(SelectedType!=''))
	{
		OpenWindow('Frame.asp?PageTitle=ɾ����ǩ&DelType='+SelectedType+'&FileName=DelTypeAndLable.asp&DelLable='+SelectedLable,200,120,window);
		location.href=location.href;
	}
	else alert('û��ѡ��Ҫɾ��������߱�ǩ');
}
function BackUpLable()
{
	var SelectedLable='',SelectedType='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.TypeID!=null)
			{
				if (SelectedType=='') SelectedType=ListObjArray[i].Obj.TypeID;
				else  SelectedType=SelectedType+'***'+ListObjArray[i].Obj.TypeID;
			}
			if (ListObjArray[i].Obj.LableID!=null)
			{
				if (SelectedLable=='') SelectedLable=ListObjArray[i].Obj.LableID;
				else  SelectedLable=SelectedLable+'***'+ListObjArray[i].Obj.LableID;
			}
		}
	}
	if (SelectedLable!='') OpenWindow('Frame.asp?PageTitle=���ݱ�ǩ&BackUpLable='+SelectedLable+'&FileName=BackUpLable.asp',200,120,window);
	else alert('��ѡ��Ҫ���ݵı�ǩ')
}
</script>
<body topmargin="2" leftmargin="2" onClick="SelectLable();" onselectstart="return false;">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=55 align="center" alt="������ǩ��Ŀ" onClick="AddLableFolder();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">�½�����</td>
		  <td width=2 class="Gray">|</td>
		  <td width=55 align="center" alt="�Զ����ǩ" onClick="AddLable();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">�½���ǩ</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="�޸�" onClick="EditFolderOrLable();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">�޸�</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="ɾ��" onClick="DelFolderAndLable();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">ɾ��</td>
		  <td width=2 class="Gray">|</td>
          <td width=35 align="center" alt="����" onClick="BackUpLable();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
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
  <td height="93" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="53%" height="26" class="ButtonListLeft"> 
            <div align="left">������</div></td>
          <td width="17%" height="26" class="ButtonList"> 
            <div align="center">����</div></td>
          <td width="30%" height="26" class="ButtonList"> 
            <div align="center">����</div></td>
        </tr>
        <%
if TypeID>0 then
%>
        <tr style="background:white;cursor:default;"> 
          <td colspan="3"><div align="left" color="#FFFFFF">
              <table border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td><img src="../../Images/arrow.gif" width="18" height="18"></td>
                  <td><span UPID="<% = UPTypeObj("ParentID") %>" title="�ϼ�����" onDblClick="ChangeUp(this)">�ϼ�����</span></td>
                </tr>
              </table>
              </div></td>
        </tr>
        <%
end if
do while not RsLblTypeObj.eof
%>
        <tr style="background:white;cursor:default;"> 
          <td height="22" ><table border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td><img src="../../Images/Folder/folderclosed.gif"></td>
                <td><span TypeID="<% = RsLblTypeObj("ID")%>" title="˫������˷���" onDblClick="ChangeLable(this)">
                  <% = RsLblTypeObj("TypeName")%>
                  </span></td>
              </tr>
            </table></td>
          <td><div align="center">��ǩ����</div></td>
          <td> <div align="center">
              <% = RsLblTypeObj("Description")%>
            </div></td>
        </tr>
        <%
	RsLblTypeObj.MoveNext
Loop
do while not RsLableObj.eof
%>
        <tr style="background:white;cursor:default;"> 
          <td height="21"><table border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td><img src="../../Images/Folder/Ffolderclosed.gif" width="21" height="15"></td>
                <td><span LableID="<%= RsLableObj("ID")%>"><% = RsLableObj("LableName")%></span></td>
              </tr>
            </table></td>
          <td><div align="center">��ǩ</div></td>
          <td height="21"> <div align="center">
              <% = RsLableObj("Description") %>
            </div></td>
        </tr>
        <%
	RsLableObj.MoveNext
loop
RsLableObj.Close
%>
      </table>
</td>
</tr>
</table>
</body>
</html>
<%
Set RsLableObj = Nothing
Set Conn = Nothing
%>
<script language="JavaScript">
var BigTypeID='<% = TypeID %>';
function SelectLable()
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
function ChangeUp(Obj)
{
	location.href='Templet_LableList.asp?TypeID='+Obj.UPID;
}
</script>