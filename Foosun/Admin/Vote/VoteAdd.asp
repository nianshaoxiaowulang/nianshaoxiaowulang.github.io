<% Option Explicit %>
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<%
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P070301") then Call ReturnError1()
%>
<html>
<head>
<link href="../../../CSS/FS_css.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>���ͶƱ��Ŀ</title>
</head>
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<body topmargin="2" leftmargin="2">
<form name="VoteForm" action="" method="post">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=35 align="center" alt="����" onClick="document.VoteForm.submit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="����" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
          <td>&nbsp; <input name="action" type="hidden" id="action" value="add"></td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
  <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="dddddd">
    <tr bgcolor="#FFFFFF"> 
      <td width="15%" height="26">��Ŀ����</td>
      <td width="81%"> 
        <input name="Name" type="text" id="Name" style="width:93%" value="<%=Request("Name")%>"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26">ѡ�����</td>
      <td> 
        <input name="OptionNum" type="text" id="OptionNum" style="width:80%" value="<% if Request("OptionNum")="" then Response.Write("4") else Response.Write(Request("OptionNum")) end if%>"> 
        <input type="button" name="Submit4" value="ȷ��" onclick="ChooseOption();"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26">��ֹʱ��</td>
      <td> 
        <input name="EndTime" type="text" id="EndTime" readonly style="width:80%" value="<%=Request("EndTime")%>"> 
        <input type="button" name="Submit42" value="����" onClick="OpenWindowAndSetValue('../../FunPages/SelectDate.asp',280,110,window,document.VoteForm.EndTime);document.VoteForm.EndTime.focus();"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26">����״̬</td>
      <td> 
        <input name="Type" type="radio" value="0" <% if Request("Type")="0" or Request("Type")="" then Response.Write("checked") end if%> >
        ��ѡ 
        <input type="radio" name="Type" value="1" <% if Request("Type")="1" then Response.Write("checked") end if%>>
        ��ѡ &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input name="State" type="radio" value="1" <% if Request("State")="0" or Request("State")="" then Response.Write("checked") end if%>>
        ���� 
        <input type="radio" name="State" value="0" <% if Request("State")="0" then Response.Write("checked") end if%>>
        �ر� </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26" colspan="2" id="Options">&nbsp;</td>
    </tr>
</table>
</form>
</body>
</html>
<script>
var RequestNameArray=new Array();
var RequestColorArray=new Array();
<%
Dim ForVal
For ForVal = 1 to Request.Form("OptionNum")
%>
RequestNameArray[<% = ForVal %>]='<% = Request.Form("Options"&ForVal&"")%>';
RequestColorArray[<% = ForVal %>]='<% = Request.Form("Color"&ForVal&"") %>';
<%
Next
%>
window.setTimeout('SetOptionsValue();',100);
function SetOptionsValue()
{
	if ((RequestNameArray.length==0)||(RequestColorArray.length==0)) return;
	var OptionNum=document.VoteForm.OptionNum.value;
	for (i=1;i<=OptionNum;i++)
	{
		document.all('Options'+i).value=RequestNameArray[i];
		document.all('Color'+i).value=RequestColorArray[i];
	}
}
function ChooseOption()
 {
  var OptionNum = document.VoteForm.OptionNum.value;
  var i,Optionstr;
	  Optionstr = '<table width="100%" border="0" cellspacing="5" cellpadding="0">';
  for (i=1;i<=OptionNum;i++)
      {
	   Optionstr = Optionstr+'<tr><td>&nbsp;ѡ&nbsp;��&nbsp;'+i+'</td><td>&nbsp;<input type="text" size="20" name="Options'+i+'" value="">&nbsp;ɫ��&nbsp;<input type="text" size="11" name="Color'+i+'" value="">&nbsp;<input type="button" name="Submit'+i+'" value="ѡɫ" onClick="OpenWindowAndSetValue(\'../../Editer/SelectColor.htm\',230,190,window,document.VoteForm.Color'+i+');"></td></tr>';
	   }
	Optionstr = Optionstr+'</table>';  
    document.all.Options.innerHTML = Optionstr;
  }
ChooseOption();
</script>
<%
  If Request.Form("action")="add" then
     Dim RsVoteObj,RsVoteSql,VoteName,OptionNum,TempVoteID,RsVoteOptionObj,RsVoteOptionSql,i
	 TempVoteID = GetRandomID18
	 If NoCSSHackAdmin(Request.Form("Name"),"��Ŀ����")="" then
	    Response.Write("<script>alert(""��Ŀ���Ʋ���Ϊ��"");</script>")
		Response.End
	 end if
	 if isnumeric(Request.Form("OptionNum"))=false then
	 	Response.Write("<script>alert(""ѡ���������Ϊ��������"");</script>")
		Response.End
	 End if
	 if Request.Form("EndTime")<>"" and isdate(Request.Form("EndTime"))=false then
	 	Response.Write("<script>alert(""��ֹʱ�����ͳ���,��������ý�ֹʱ��,���ÿ�"");</script>")
		Response.End
	 end if
	 For i=1 to Request.Form("OptionNum")
	     if Request.Form("Options"&i&"")="" or isnull(Request.Form("Options"&i&"")) then
		    Response.Write("<script>alert(""ѡ��"&i&"���ݲ���Ϊ��"");</script>")
			Response.End
		 end if
	 Next
     Set RsVoteObj = Server.CreateObject(G_FS_RS)
	 	 RsVoteSql = "Select * from FS_Vote"
		 RsVoteObj.Open RsVoteSql,Conn,3,3
		 RsVoteObj.AddNew
		 RsVoteObj("VoteID") = TempVoteID
		 RsVoteObj("Name") = Replace(Replace(Request.Form("Name"),"""",""),"'","")
		 RsVoteObj("OptionNum") = Cint(Request.Form("OptionNum"))
		 If Request.Form("Type") = "1" then
			 RsVoteObj("Type") = "1"
		 Else
			 RsVoteObj("Type") = "0"
		 End if
		 If Request.Form("State") = "1" then
		    If Isdate(Request.Form("EndTime")) then
				If datediff("s",now(),formatdatetime(Request.Form("EndTime")))<0 then
				 RsVoteObj("State") = "2"
				Else
				 RsVoteObj("State") = "1"
				End If
			Else
			 RsVoteObj("State") = "1"
		    End if
		 Else
		    If Isdate(Request.Form("EndTime")) then
				If datediff("s",now(),formatdatetime(Request.Form("EndTime")))<0 then
				 RsVoteObj("State") = "2"
				Else
				 RsVoteObj("State") = "0"
				End If
			Else
			 RsVoteObj("State") = "0"
		    End if
		 End if
		 If Isdate(Request.Form("EndTime")) then
			 RsVoteObj("EndTime") = formatdatetime(Request.Form("EndTime"))
		 Else
			 RsVoteObj("EndTime") = "0"
		 End if
		 RsVoteObj("AddTime") = Now()
		 RsVoteObj.Update
		 RsVoteObj.Close
		 Set RsVoteObj = Nothing
			 For i = 1 to Request.Form("OptionNum")
				 Set RsVoteOptionObj = Server.CreateObject(G_FS_RS)
				 RsVoteOptionSql = "Select * from FS_VoteOption where 1=0"
				 RsVoteOptionObj.Open RsVoteOptionSql,Conn,3,3
				 RsVoteOptionObj.AddNew
			 	 RsVoteOptionObj("VoteID") = TempVoteID
			 	 RsVoteOptionObj("OptionName") = Replace(Replace(Request.Form("Options"&i&""),"""",""),"'","")
			 	 If Request.Form("Color"&i&"")<>"" then
					 RsVoteOptionObj("OptionColor") = Request.Form("Color"&i&"")
				 End If
			 	 RsVoteOptionObj("ObjTaxis") = i
				 RsVoteOptionObj.Update
				 RsVoteOptionObj.Close
			 Next
			 Set RsVoteOptionObj = Nothing
		 Response.Redirect("VoteList.asp")
		 Response.End
  End if
  Set Conn = Nothing
%>