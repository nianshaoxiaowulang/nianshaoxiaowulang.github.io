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
if Not JudgePopedomTF(Session("Name"),"P070300") then Call ReturnError1()
'Conn.Execute("Update Vote set State=2 where EndTime<>'0' and EndTime<='"&Now()&"'")
Dim RsVoteObj
Set RsVoteObj = Server.CreateObject(G_FS_RS)
RsVoteObj.open "Select * from FS_Vote order by AddTime desc",conn,1,1
%>
<html>
<head>
<link href="../../../CSS/FS_css.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>ͶƱ�б�</title>
</head>
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<script language="JavaScript" src="../../SysJS/ContentMenu.js"></script>
<body topmargin="2" leftmargin="2" onclick="SelectVote();" ondragstart="return false;" onselectstart="return false;">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=35 align="center" alt="�½�ͶƱ" onClick="AddVote();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">�½�</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35  align="center" alt="�޸�ͶƱ" onClick="EditVote();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">�޸�</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="ɾ��ͶƱ" onClick="DelVote();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">ɾ��</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="����ͶƱ" onClick="OpenVote();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  <td width=2 class="Gray">|</td>
          <td width=35 align="center" alt="�ر�ͶƱ" onClick="CloseVote();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">�ر�</td>
		  <td width=2 class="Gray">|</td>
		  <td width=55 align="center" alt="�鿴ͶƱ���" onClick="BrowResultOfVote();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">�鿴���</td>
		  <td width=2 class="Gray">|</td>
		  <td width=55 align="center" alt="���ô���" onClick="GetCode();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">���ô���</td>
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
	<td valign="top">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
	      <td width="41%" height="26" class="ButtonListLeft"> 
            <div align="center">��Ŀ����</div></td>
	      <td width="14%" height="26" class="ButtonList"> 
            <div align="center">����</div></td> 
	      <td width="11%" height="26" class="ButtonList"> 
            <div align="center">ѡ��</div></td> 
	      <td width="11%" height="26" class="ButtonList"> 
            <div align="center">״̬</div></td> 
	      <td width="23%" height="26" class="ButtonList"> 
            <div align="center">���ʱ��</div></td> 
  </tr>
  <%
if Not RsVoteObj.Eof then
	Dim page_size,page_no,page_total,record_all,PageNums
	page_size = 10
	page_no=request.querystring("page_no")
	if page_no<=1 or page_no="" then page_no=1
	If Request.QueryString("page_no")="" then
		page_no=1
	end if
	RsVoteObj.PageSize=page_size
	page_total=RsVoteObj.PageCount
	if (cint(page_no)>page_total) then page_no=page_total
	RsVoteObj.AbsolutePage=page_no
	record_all=RsVoteObj.RecordCount
	Dim i
	for i=1 to RsVoteObj.PageSize
		if RsVoteObj.eof then exit for
		Dim Types,States
		If RsVoteObj("Type") = "0" then
			Types = "��ѡ"
		else
			Types = "��ѡ"
		end if
		Select case RsVoteObj("State")
		  Case "0" States="�ر�"
		  Case "1" States="����"
		  Case "2" States="����"
		 End Select
%>
  <tr height="20"> 
	      <td><table border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><img src="../../Images/Vote.gif" width="24" height="22"></td>
                <td><span class="TempletItem" VoteID="<%=RsVoteObj("VoteID")%>" State="<%=RsVoteObj("State")%>"><%=RsVoteObj("Name")%></span></td>
              </tr>
            </table></td>
	<td><div align="center"><%=Types%></div></td>
	<td><div align="center"><%=RsVoteObj("OptionNum")%></div></td>
	<td><div align="center"><%=States%></div></td>
	<td><div align="center"><%=RsVoteObj("AddTime")%></div></td>
  </tr>
  <%
		RsVoteObj.MoveNext
	Next
end if
RsVoteObj.close
Set RsVoteObj = Nothing
%>
</table></td></tr>
<tr>
<td valign="middle" height="10">
<table width="100%" border="0" cellspacing="0" cellpadding="0" class="ButtonListLeft">
	  <tr height="1">
		  <td height="25" valign="middle"> 
            <div align="right"> 
              <% =  "NO.<b>"& page_no &"</b>,&nbsp;&nbsp;" %>
			<% = "Totel:<b>"& page_total &"</b>,&nbsp;RecordCounts:<b>" & record_all &"</b>&nbsp;&nbsp;&nbsp;"%>
			<%
				if Page_Total=1 then
						response.Write "&nbsp;<img src=../../images/FirstPage.gif border=0 alt=��ҳ></img>&nbsp;"
						response.Write "&nbsp;<img src=../../images/prePage.gif border=0 alt=��һҳ></img>&nbsp;"
						response.Write "&nbsp;<img src=../../images/nextPage.gif border=0 alt=��һҳ></img>&nbsp;"
						response.Write "&nbsp;<img src=../../images/endPage.gif border=0 alt=βҳ></img>&nbsp;"
				else
					if cint(Page_No)<>1 and cint(Page_No)<>Page_Total then
						response.Write "&nbsp;<a href=?page_no=1" &"&Keywords="&Request("Keywords")&"><img src=../../images/FirstPage.gif border=0 alt=��ҳ></img></a>&nbsp;"
						response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)-1) &"&Keywords="&Request("Keywords")&"><img src=../../images/prePage.gif border=0 alt=��һҳ></img></a>&nbsp;"
						response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)+1) &"&Keywords="&Request("Keywords")&"><img src=../../images/nextPage.gif border=0 alt=��һҳ></img></a>&nbsp;"
						response.Write "&nbsp;<a href=?page_no="& Page_Total &"&Keywords="&Request("Keywords")&"><img src=../../images/endPage.gif border=0 alt=βҳ></img></a>&nbsp;"
					elseif cint(Page_No)=1 then
						response.Write "&nbsp;<img src=../../images/FirstPage.gif border=0 alt=��ҳ></img></a>&nbsp;"
						response.Write "&nbsp;<img src=../../images/prePage.gif border=0 alt=��һҳ></img>&nbsp;"
						response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)+1) &"&Keywords="&Request("Keywords")&"><img src=../../images/nextPage.gif border=0 alt=��һҳ></img></a>&nbsp;"
						response.Write "&nbsp;<a href=?page_no="& Page_Total &"&Keywords="&Request("Keywords")&"><img src=../../images/endpage.gif border=0 alt=βҳ></img></a>&nbsp;"
					else
						response.Write "&nbsp;<a href=?page_no=1" &"&Keywords="&Request("Keywords")&"><img src=../../images/FirstPage.gif border=0 alt=��ҳ></img>&nbsp;"
						response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)-1)& "&Keywords="&Request("Keywords")&"><img src=../../images/prePage.gif border=0 alt=��һҳ></img></a>&nbsp;"
						response.Write "&nbsp;<img src=../../images/nextPage.gif border=0 alt=��һҳ></img></a>&nbsp;"
						response.Write "&nbsp;<img src=../../images/endpage.gif border=0 alt=βҳ></img>&nbsp;"
					end if
				end if
				%>
		</div></td>
		  <td width="50" valign="middle">
<select onChange="ChangePage(this.value);" style="width:50;" name="select">
		  <% for i=1 to Page_Total %>
		  <option <% if cint(Page_No) = i then Response.Write("selected")%> value="<% = i %>">
		  <% = i %>
		  </option>
		  <% next %>
		</select>
          </td>
	  </tr>
	</table>
	</td>
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
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.EditVote();",'�޸�','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.DelVote();",'ɾ��','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('seperator','','');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.CloseVote();','��ͣ','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.OpenVote();','����','disabled');
	//ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.BrowResultOfVote();','�鿴','disabled');
	//ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.GetCode();','���ô���','disabled');
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
			if (SelectContent=='') SelectContent=ListObjArray[i].Obj.VoteID;
			else SelectContent=SelectContent+'***'+ListObjArray[i].Obj.VoteID;
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
					if (SelectContent=='') SelectContent=ListObjArray[i].Obj.VoteID;
					else SelectContent=SelectContent+'***'+ListObjArray[i].Obj.VoteID;
				}
			}
		}
	}
	if (SelectContent=='') DisabledContentMenuStr=',�޸�,ɾ��,��ͣ,����,�鿴,���ô���,';
	else
	{
		if (SelectContent.indexOf('***')==-1) DisabledContentMenuStr='';
		else DisabledContentMenuStr=',�޸�,�鿴,���ô���,'
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
		if (CurrObj.VoteID!=null)
		{
			ListObjArray[ListObjArray.length]=new FolderFileObj(CurrObj,j,false);
			j++;
		}
	}
}
function SelectVote()
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
function AddVote()
{
	location='VoteAdd.asp';
}
function EditVote()
{
	var SelectedVote='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.VoteID!=null)
			{
				if (SelectedVote=='') SelectedVote=ListObjArray[i].Obj.VoteID;
				else  SelectedVote=SelectedVote+'***'+ListObjArray[i].Obj.VoteID;
			}
		}
	}
	if (SelectedVote!='')
	{
		if (SelectedVote.indexOf('***')==-1) location='VoteModify.asp?VoteID='+SelectedVote;
		else alert('һ��ֻ�ܹ��޸�һ��ͶƱ');
	}
	else alert('��ѡ��Ҫ�޸ĵ�ͶƱ');
}
function DelVote()
{
	var SelectedVote='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.VoteID!=null)
			{
				if (SelectedVote=='') SelectedVote=ListObjArray[i].Obj.VoteID;
				else  SelectedVote=SelectedVote+'***'+ListObjArray[i].Obj.VoteID;
			}
		}
	}
	if (SelectedVote!='')
		OpenWindow('Frame.asp?FileName=VoteDell.asp&Types=Dell&PageTitle=ɾ��ͶƱ��Ŀ&VoteID='+SelectedVote,220,105,window);
	else alert('��ѡ��Ҫ�޸ĵ�ͶƱ');
}
function OpenVote()
{
	var SelectedVote='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.VoteID!=null)
			{
				if (SelectedVote=='') SelectedVote=ListObjArray[i].Obj.VoteID;
				else  SelectedVote=SelectedVote+'***'+ListObjArray[i].Obj.VoteID;
			}
		}
	}
	if (SelectedVote!='')
		OpenWindow('Frame.asp?FileName=VoteDell.asp&Types=Open&PageTitle=����ͶƱ&VoteID='+SelectedVote,320,110,window);
	else alert('��ѡ��ͶƱ');
}
function CloseVote()
{
	var SelectedVote='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.VoteID!=null)
			{
				if (SelectedVote=='') SelectedVote=ListObjArray[i].Obj.VoteID;
				else  SelectedVote=SelectedVote+'***'+ListObjArray[i].Obj.VoteID;
			}
		}
	}
	if (SelectedVote!='')
		OpenWindow('Frame.asp?FileName=VoteDell.asp&Types=Close&PageTitle=ͶƱ��Ŀ��������&VoteID='+SelectedVote,320,110,window);
	else alert('��ѡ��ͶƱ');
}
function BrowResultOfVote()
{
	var SelectedVote='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.VoteID!=null)
			{
				if (SelectedVote=='') SelectedVote=ListObjArray[i].Obj.VoteID;
				else  SelectedVote=SelectedVote+'***'+ListObjArray[i].Obj.VoteID;
			}
		}
	}
	if (SelectedVote!='')
	{
		if (SelectedVote.indexOf('***')==-1) OpenWindow('Frame.asp?FileName=../../../<%=PlusDir%>/Vote/VoteResult.asp&PageTitle=�鿴ͶƱ���&VoteID='+SelectedVote,420,200,window);
		else alert('��ѡ��һ��ͶƱ');
	}
	else alert('��ѡ��ͶƱ');
}
function GetCode()
{
	var SelectedVote='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.VoteID!=null)
			{
				if (SelectedVote=='') SelectedVote=ListObjArray[i].Obj.VoteID;
				else  SelectedVote=SelectedVote+'***'+ListObjArray[i].Obj.VoteID;
			}
		}
	}
	if (SelectedVote!='')
	{
		if (SelectedVote.indexOf('***')==-1) OpenWindow('Frame.asp?FileName=VoteCode.asp&Types=Code&PageTitle=ͶƱ���ô���&VoteID='+SelectedVote,500,180,window);
		else alert('��ѡ��һ��ͶƱ');
	}
	else alert('��ѡ��ͶƱ');
}
function ChangePage(PageNum)
{
	window.location.href='?page_no='+PageNum;
}
</script>
