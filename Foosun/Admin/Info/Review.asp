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
if Not JudgePopedomTF(Session("Name"),"P010509") then Call ReturnError1()
Dim OperationID
OperationID = Request("OperationID")
if Request("Action") = "Del" then
	if OperationID <> "" then
		OperationID = Replace(OperationID,"***",",")
		Conn.execute("Delete From FS_Review Where ID in (" & OperationID & ")")
	end if
elseif Request("Action") = "Audit" then
	if OperationID <> "" then
		OperationID = Replace(OperationID,"***",",")
		Conn.execute("Update FS_Review Set Audit=1 Where ID in (" & OperationID & ")")
	end if
elseif Request("Action") = "CancelAudit" then
	if OperationID <> "" then
		OperationID = Replace(OperationID,"***",",")
		Conn.execute("Update FS_Review Set Audit=0 Where ID in (" & OperationID & ")")
	end if
end if

Dim NewsID,RsReviewObj,RsNewsObj,SunNumAsp,sql,isn,ShowPagesTF,DownloadID,Sqlstr
ShowPagesTF = True
If Request("NewsID")<>"" then
	NewsID = Cstr(Request("NewsID"))
	isn="where Types = 1 and NewsID='"&NewsID&"'"
elseif Request("DownloadID")<>"" Then
	DownloadID = Cstr(Request("DownloadID"))
	isn="where Types = 2 and NewsID='"&DownloadID&"'"
end if
Sql = "Select * from FS_Review  "&isn&" order by id desc"
Set RsReviewObj = Server.CreateObject(G_FS_RS)
RsReviewObj.Open Sql,Conn,1,1
SunNumAsp = RsReviewObj.RecordCount
%>
<html>
<head>
<link href="../../../CSS/FS_css.css" rel="stylesheet" type="text/css">
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<script language="JavaScript" src="../../SysJS/ContentMenu.js"></script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>������ŵ�ר��</title>
</head>
<body topmargin="2" leftmargin="2" onclick="SelectReview();" onselectstart="return false;">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=35 align="center" alt="ɾ��" onClick="DelReview();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">ɾ��</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35  align="center" alt="�޸�" onClick="EditReview();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">�޸�</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="���" onClick="Audit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">���</td>
		  <td width=2 class="Gray">|</td>
          <td width=55 align="center" alt="ȡ�����" onClick="CancelAudit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">ȡ�����</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="����" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
          <td>&nbsp;</td>
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
    <td height="26" class="ButtonListLeft"> 
      <div align="center">���� </div></td>
    <td width="8%" height="26" class="ButtonList">
<div align="center">���</div></td>
    <td width="10%" height="26" class="ButtonList"> 
      <div align="center">������</div></td>
    <td width="15%" height="26" class="ButtonList"> 
      <div align="center">����ʱ��</div></td>
	<td width="8%" height="26" class="ButtonList"> 
      <div align="center">��������</div></td>
    <td width="25%" height="26" class="ButtonList"> 
		<div align="center">
		��������
		</div>
	</td>
  </tr>
  <%
if not  RsReviewObj.Bof And not RsReviewObj.Eof  then 
	Dim page_size,page_no,page_total,record_all
	page_size=20
	page_no=request.querystring("page_no")
    if page_no<=1 or page_no="" then page_no=1
	If Request.QueryString("page_no")="" then
		page_no=1
	end if
	RsReviewObj.PageSize=page_size
	page_total=RsReviewObj.PageCount
	if (cint(page_no)>page_total) then page_no=page_total
	RsReviewObj.AbsolutePage=cint(page_no)
	record_all=RsReviewObj.RecordCount
	Dim i
  	for i=1 to RsReviewObj.PageSize
    	if RsReviewObj.eof then exit for
%>
  <tr> 
    <td nowrap> <table border="0" cellspacing="0" cellpadding="0">
        <tr height="22"> 
          <td><img src="../../Images/Info/WordNews.gif" width="24" height="22"></td>
          <td nowrap><span class="TempletItem" NewsID="<% = RsReviewObj("NewsID") %>" ReviewID="<% = RsReviewObj("ID") %>"><%=left(RsReviewObj("Content"),20)%></span></td>
        </tr>
      </table></td>
    <td nowrap><div align="center"><% if RsReviewObj("Audit") = 1 then Response.Write("�����") else Response.Write("<font color=""red"">δ���</fon>") %></div></td>
    <td nowrap> <div align="center"><%=RsReviewObj("UserID")%></div></td>
    <td nowrap> <div align="center"><%=RsReviewObj("Addtime")%></div></td>
	<td nowrap> <div align="center"><%
	If RsReviewObj("Types")="1" then 
		response.write "����"
	ElseIf RsReviewObj("Types")="2" then 
		response.write "����"
	else
		response.write "��Ʒ"
	End If
	%></div></td>
    <td nowrap> <div align="center"> 
        <%
		If RsReviewObj("Types")="1" then
			Set RsNewsObj = Conn.Execute("Select TiTle from FS_News where newsid='" &RsReviewObj("NewsID")&"'")	
			If Not RsNewsObj.eof Then 
				response.Write(""&left(RsNewsObj("TiTle"),10)&"")
			Else
				Response.Write("���������ѱ�ɾ��")
			End If 
		elseif RsReviewObj("Types")="2" Then
			Set RsNewsObj = Conn.Execute("Select Name from FS_Download where DownLoadID='" &  RsReviewObj("NewsID")&"'")	
			If Not RsNewsObj.eof Then 
				response.Write(""&left(RsNewsObj("Name"),10)&"")
			Else
				Response.Write("���������ѱ�ɾ��")
			End If
		else
			Set RsNewsObj = Conn.Execute("Select Product_Name from FS_Shop_Products where ID=" &  RsReviewObj("NewsID"))	
			If Not RsNewsObj.eof Then 
				response.Write(""&left(RsNewsObj("Product_Name"),10)&"")
			Else
				Response.Write("���������ѱ�ɾ��")
			End If
		end if
		%>
      </div></td>
  </tr>
  <%
		  RsReviewObj.movenext
	next
	RsReviewObj.close
	set RsReviewObj=nothing
	if page_total > 1 then
		ShowPagesTF = True
	else
		ShowPagesTF = False
	end if
else
	 ShowPagesTF = False
end if
if ShowPagesTF = True then	 
%>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td height="26" colspan="5" class="ButtonListLeft"> <div align="right"><strong> 
<%
	if page_total=1 then
			response.Write "&nbsp;<img src=../../images/FirstPage.gif border=0 alt=��ҳ></img>&nbsp;"
			response.Write "&nbsp;<img src=../../images/prePage.gif border=0 alt=��һҳ></img>&nbsp;"
			response.Write "&nbsp;<img src=../../images/nextPage.gif border=0 alt=��һҳ></img>&nbsp;"
			response.Write "&nbsp;<img src=../../images/endPage.gif border=0 alt=βҳ></img>&nbsp;"
	else
		if cint(page_no)<>1 and cint(page_no)<>page_total then
			response.Write "&nbsp;<a href=?page_no=1&Newsid="&request("Newsid")&"&Downloadid="&request("Downloadid")&"><img src=../../images/FirstPage.gif border=0 alt=��ҳ></img></a>&nbsp;"
			response.Write "&nbsp;<a href=?page_no="&cstr(cint(page_no)-1)&"&Newsid="&request("Newsid")&"&Downloadid="&request("Downloadid")&"><img src=../../images/prePage.gif border=0 alt=��һҳ></img></a>&nbsp;"
			response.Write "&nbsp;<a href=?page_no="&cstr(cint(page_no)+1)&"&Newsid="&request("Newsid")&"&Downloadid="&request("Downloadid")&"><img src=../../images/nextPage.gif border=0 alt=��һҳ></img></a>&nbsp;"
			response.Write "&nbsp;<a href=?page_no="& page_total &"&Newsid="&request("Newsid")&"&Downloadid="&request("Downloadid")&"><img src=../../images/endPage.gif border=0 alt=βҳ></img></a>&nbsp;"
		elseif cint(page_no)=1 then
			response.Write "&nbsp;<img src=../../images/FirstPage.gif border=0 alt=��ҳ></img></a>&nbsp;"
			response.Write "&nbsp;<img src=../../images/prePage.gif border=0 alt=��һҳ></img>&nbsp;"
			response.Write "&nbsp;<a href=?page_no="&cstr(cint(page_no)+1)&"&Newsid="&request("Newsid")&"&Downloadid="&request("Downloadid")&"><img src=../../images/nextPage.gif border=0 alt=��һҳ></img></a>&nbsp;"
			response.Write "&nbsp;<a href=?page_no="& page_total &"&Newsid="&request("Newsid")&"&Downloadid="&request("Downloadid")&"><img src=../../images/endPage.gif border=0 alt=βҳ></img></a>&nbsp;"
		else
			response.Write "&nbsp;<a href=?page_no=1&Newsid="&request("Newsid")&"&Downloadid="&request("Downloadid")&"><img src=../../images/FirstPage.gif border=0 alt=��ҳ></img></a>&nbsp;"
			response.Write "&nbsp;<a href=?page_no="&cstr(cint(page_no)-1)&"&Newsid="&request("Newsid")&"&Downloadid="&request("Downloadid")&"><img src=../../images/prePage.gif border=0 alt=��һҳ></img></a>&nbsp;"
			response.Write "&nbsp;<img src=../../images/nextPage.gif border=0 alt=��һҳ></img></a>&nbsp;"
			response.Write "&nbsp;<img src=../../images/endPage.gif border=0 alt=βҳ></img>&nbsp;"
		end if
	end if
%>
        </strong></div></td>
  </tr>
  <%
end if
%>
</table>
</body>
</html>
<script language="JavaScript">
var NewsID='<% = NewsID %>';
var	DownloadID = '<%= DownloadID%>';
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
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.EditReview();",'�޸�','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.DelReview();",'ɾ��','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.Audit();','���','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.CancelAudit();','ȡ�����','disabled');
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
	var EventObjInArray=false,SelectedReview='',DisabledContentMenuStr='';
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
			if (SelectedReview=='') SelectedReview=ListObjArray[i].Obj.ReviewID;
			else SelectedReview=SelectedReview+'***'+ListObjArray[i].Obj.ReviewID;
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
					if (SelectedReview=='') SelectedReview=ListObjArray[i].Obj.ReviewID;
					else SelectedReview=SelectedReview+'***'+ListObjArray[i].Obj.ReviewID;
				}
			}
		}
	}
	if (SelectedReview=='') DisabledContentMenuStr=',ɾ��,���,ȡ�����,�޸�,';
	else
	{
		if (SelectedReview.indexOf('***')==-1) DisabledContentMenuStr='';
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
		if (CurrObj.ReviewID!=null)
		{
			ListObjArray[ListObjArray.length]=new FolderFileObj(CurrObj,j,false);
			j++;
		}
	}
}
function SelectReview()
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
function DelReview()
{
	var SelectedReview='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.ReviewID!=null)
			{
				if (SelectedReview=='') SelectedReview=ListObjArray[i].Obj.ReviewID;
				else  SelectedReview=SelectedReview+'***'+ListObjArray[i].Obj.ReviewID;
			}
		}
	}
	if (SelectedReview!='')
	{
		if (confirm('ȷ��Ҫɾ����?')) location='?NewsID='+NewsID+'&Action=Del&OperationID='+SelectedReview+'&DownloadID='+DownloadID;	}
	else alert('��ѡ������');
}
function Audit()
{
	var SelectedReview='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.ReviewID!=null)
			{
				if (SelectedReview=='') SelectedReview=ListObjArray[i].Obj.ReviewID;
				else  SelectedReview=SelectedReview+'***'+ListObjArray[i].Obj.ReviewID;
			}
		}
	}
	if (SelectedReview!='')
		location='?NewsID='+NewsID+'&Action=Audit&OperationID='+SelectedReview+'&DownloadID='+DownloadID;
	else alert('��ѡ������');
}
function CancelAudit()
{
	var SelectedReview='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.ReviewID!=null)
			{
				if (SelectedReview=='') SelectedReview=ListObjArray[i].Obj.ReviewID;
				else  SelectedReview=SelectedReview+'***'+ListObjArray[i].Obj.ReviewID;
			}
		}
	}
	if (SelectedReview!='')
		location='?NewsID='+NewsID+'&Action=CancelAudit&OperationID='+SelectedReview+'&DownloadID='+DownloadID;
	else alert('��ѡ������');
}
function EditReview()
{
	var SelectedReview='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.ReviewID!=null)
			{
				if (SelectedReview=='') SelectedReview=ListObjArray[i].Obj.ReviewID;
				else  SelectedReview=SelectedReview+'***'+ListObjArray[i].Obj.ReviewID;
			}
		}
	}
	if (SelectedReview!='')
	{
		if (SelectedReview.indexOf('***')==-1) location='ReviewEdit.asp?NewsID='+NewsID+'&ReviewID='+SelectedReview+"&DownloadID="+DownloadID;
		else DisabledContentMenuStr=',�޸�,'
	}
	else alert('��ѡ������');

}
</script>