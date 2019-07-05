<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
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
if Not JudgePopedomTF(Session("Name"),"P070302") then Call ReturnError1()
    Dim VoteID,RsVoteModObj,RequestNameArrays,RequestColorArrays,ROPObj,ROPSql
	VoteID = Request("VoteID")
	RequestNameArrays = ""
	RequestColorArrays = ""
	Set RsVoteModObj = Conn.Execute("Select * from FS_Vote where VoteID='" & VoteID & "'")
	If RsVoteModObj.eof then
	   Response.Write("<script>alert(""参数传递错误"");dialogArguments.location.reload();window.close();</script>")
	   Response.End
	End If
	Set ROPObj = Server.CreateObject(G_FS_RS)
	ROPSql = "Select * from FS_VoteOption where VoteID='"&VoteID&"' order by ObjTaxis asc"
	ROPObj.Open ROPSql,Conn,3,3
	for i = 0 to RsVoteModObj("OptionNum")-1
			RequestNameArrays = RequestNameArrays & "," & ROPObj("OptionName")
			RequestColorArrays = RequestColorArrays & "," & ROPObj("OptionColor")
		ROPObj.MoveNext
	next
	ROPObj.Close
	Set ROPObj = Nothing
%>
<html>
<head>
<link href="../../../CSS/FS_css.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>修改投票项目</title>
</head>
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<body topmargin="2" leftmargin="2">
<form name="VoteForm" action="" method="post">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=35 align="center" alt="保存" onClick="document.VoteForm.submit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">保存</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="后退" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
          <td>&nbsp; <input name="action" type="hidden" id="action2" value="add"></td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
  <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="dddddd">
    <tr bgcolor="#FFFFFF"> 
      <td width="15%" height="26">项目名称</td>
      <td width="81%"> 
        <input name="Name" type="text" id="Name" style="width:93%" value="<%=RsVoteModObj("Name")%>"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26">选项个数</td>
      <td> 
        <input name="OptionNum" type="text" id="OptionNum" style="width:80%" value="<%=RsVoteModObj("OptionNum")%>"> 
        <input type="button" name="Submit4" value="确定" onClick="ChooseOption();"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26">截止时间</td>
      <td> 
        <input name="EndTime" type="text" id="EndTime" readonly style="width:80%" value="<% if RsVoteModObj("EndTime")="0" then Response.Write("") else Response.Write(RsVoteModObj("EndTime")) end if%>"> 
        <input type="button" name="Submit42" value="日期" onClick="OpenWindowAndSetValue('../../FunPages/SelectDate.asp',280,110,window,document.VoteForm.EndTime);document.VoteForm.EndTime.focus();"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26">类型状态</td>
      <td> 
        <input name="Type" type="radio" value="0" <% if RsVoteModObj("Type")="0" then Response.Write("checked") end if%> >
        单选 
        <input type="radio" name="Type" value="1" <% if RsVoteModObj("Type")="1" then Response.Write("checked") end if%>>
        多选 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input name="State" type="radio" value="1" <% if RsVoteModObj("State")="1" then Response.Write("checked") end if%>>
        开启 
        <input type="radio" name="State" value="0" <% if RsVoteModObj("State")="0" then Response.Write("checked") end if%>>
        关闭 </td>
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
var TempRequestNameArray,TempRequestColorArray;
TempRequestNameArray='<% = RequestNameArrays %>';
TempRequestColorArray='<% = RequestColorArrays %>';
RequestNameArray = TempRequestNameArray.split(",");
RequestColorArray = TempRequestColorArray.split(",");
window.setTimeout('SetOptionsValue();',100);
function SetOptionsValue()
{
	if ((RequestNameArray.length==0)||(RequestColorArray.length==0)) return;
	var OptionNum=document.VoteForm.OptionNum.value;
	for (i=1;i<=OptionNum;i++)
	{	if (typeof(RequestNameArray[i])=='string')
		{
		document.all('Options'+i).value=RequestNameArray[i];
		document.all('Color'+i).value=RequestColorArray[i];
		}
	}
}
function ChooseOption()
 {
  var OptionNum = document.VoteForm.OptionNum.value;
  var i,Optionstr;
	  Optionstr = '<table width="100%" border="0" cellspacing="5" cellpadding="0">';
  for (i=1;i<=OptionNum;i++)
      {
	   Optionstr = Optionstr+'<tr><td>&nbsp;选&nbsp;项&nbsp;'+i+'</td><td>&nbsp;<input type="text" size="20" name="Options'+i+'" value="">&nbsp;色彩&nbsp;<input type="text" size="11" name="Color'+i+'" value="">&nbsp;<input type="button" name="Submit'+i+'" value="选色" onClick="OpenWindowAndSetValue(\'../../Editer/SelectColor.htm\',230,190,window,document.VoteForm.Color'+i+');"></td></tr>';
	   }
	Optionstr = Optionstr+'</table>';  
    document.all.Options.innerHTML = Optionstr;
	SetOptionsValue()
  }
ChooseOption();
</script>
<%
  If Request.Form("action")="add" then
     Dim RsVoteObj,RsVoteSql,VoteName,OptionNum,RsVoteOptionObj,RsVoteOptionSql,i
	 If NoCSSHackAdmin(Request.Form("Name"),"项目名称")="" then
	    Response.Write("<script>alert(""项目名称不能为空"");</script>")
		Response.End
	 end if
	 if isnumeric(Request.Form("OptionNum"))=false then
	 	Response.Write("<script>alert(""选项个数必须为数字类型"");</script>")
		Response.End
	 End if
	 if Request.Form("EndTime")<>"" and isdate(Request.Form("EndTime"))=false then
	 	Response.Write("<script>alert(""截止时间类型出错,如果不设置截止时间,请置空"");</script>")
		Response.End
	 end if
	 For i=1 to Request.Form("OptionNum")
	     if Request.Form("Options"&i&"")="" or isnull(Request.Form("Options"&i&"")) then
		    Response.Write("<script>alert(""选项"&i&"内容不能为空"");</script>")
			Response.End
		 end if
	 Next
     Set RsVoteObj = Server.CreateObject(G_FS_RS)
	 	 RsVoteSql = "Select * from FS_Vote where VoteID='"&VoteID&"'"
		 RsVoteObj.Open RsVoteSql,Conn,3,3
		 RsVoteObj("Name") = Replace(Replace(Request.Form("Name"),"""",""),"'","")
		 RsVoteObj("OptionNum") = Cint(Request.Form("OptionNum"))
		 If Request.Form("Type") = "1" then
			 RsVoteObj("Type") = "1"
		 ElseIf Request.Form("Type") = "0" then
			 RsVoteObj("Type") = "0"
		 End if
		 If Request.Form("State") <> "0" and Request.Form("State")<>"" then
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
		 RsVoteObj.Update
		 RsVoteObj.Close
		 Set RsVoteObj = Nothing
		 Conn.Execute("Delete from FS_VoteOption where VoteID='"&VoteID&"'")
			 For i = 1 to Request.Form("OptionNum")
				 Set RsVoteOptionObj = Server.CreateObject(G_FS_RS)
				 RsVoteOptionSql = "Select * from FS_VoteOption where VoteID='"&VoteID&"' and ObjTaxis=" & i
				 RsVoteOptionObj.Open RsVoteOptionSql,Conn,3,3
				 If RsVoteOptionObj.eof then RsVoteOptionObj.AddNew
			 	 RsVoteOptionObj("VoteID") = Cstr(VoteID)
			 	 RsVoteOptionObj("OptionName") = Replace(Replace(Request.Form("Options"&i&""),"""",""),"'","")
			 	 If Request.Form("Color"&i&"")<>"" then
					 RsVoteOptionObj("OptionColor") = Request.Form("Color"&i&"")
				 Else
					 RsVoteOptionObj("OptionColor") = "red"
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