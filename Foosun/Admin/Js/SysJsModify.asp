<% option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<!--#include file="../Refresh/Function.asp" -->
<%  dim DBC,Conn,ClassParentID
    set DBC=new databaseclass
	set Conn=DBC.openconnection()
	set DBC=nothing
%>
<!--#include file="../Inc/Cls_RefreshJs.asp" -->
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not ((JudgePopedomTF(Session("Name"),"P060602")) OR (JudgePopedomTF(Session("Name"),"P060502"))) then Call ReturnError1()
Dim FileID,RsSysModObj,FunClassID,Types,ScrLink,ScrNewsType,TeempSysRootDir
If SysRootDir = "" then
	TeempSysRootDir = ""
Else
	TeempSysRootDir = SysRootDir & "/"
End If
If Request("FileID")="" or isnull(Request("FileID")) then
	Response.Write("<script>alert(""�������ݴ���"");window.close();</script>")
	Response.End
Else
	FileID = Request("FileID")
	Types = Request("Types")
	Set RsSysModObj = Conn.Execute("Select * from FS_SysJs where ID="&FileID&"")
	If RsSysModObj.eof then
		Response.Write("<script>alert(""δ��ѯ����ؼ�¼"");window.close();</script>")
		Response.End
	End IF
	FunClassID = RsSysModObj("ClassID")
	ScrLink = RsSysModObj("MoreContent")
	ScrNewsType = RsSysModObj("NewsType")
End IF
Dim TempClassListStr
	TempClassListStr = ClassList
Function ClassList()
	Dim Rs,SelectStr
	Set Rs = Conn.Execute("select ClassID,ClassCName from FS_newsclass where ParentID = '0' and DelFlag=0 order by AddTime desc")
	do while Not Rs.Eof
		If Cstr(FunClassID) = Cstr(Rs("ClassID")) then
			SelectStr = " selected"
		Else
			SelectStr = ""
		End If
		ClassList = ClassList & "<option value="""&Rs("ClassID")&""""& SelectStr & ">" & Rs("ClassCName") & chr(10) & chr(13)
		ClassList = ClassList & ChildClassList(Rs("ClassID"),"")
		Rs.MoveNext	
	loop
	Rs.Close
	Set Rs = Nothing
End Function
Function ChildClassList(ClassID,Temp)
	Dim TempRs,TempStr,SelectStrs
	Set TempRs = Conn.Execute("Select ClassID,ClassCName,ChildNum from FS_NewsClass where ParentID = '" & ClassID & "' and DelFlag=0 order by AddTime desc ")
	TempStr = Temp & " - "
	do while Not TempRs.Eof
		If Cstr(FunClassID) = Cstr(TempRs("ClassID")) then
			SelectStrs = " selected"
		Else
			SelectStrs = ""
		End If
		if TempRs("ChildNum") = 0 then
			ChildClassList = ChildClassList & "<option value="""&TempRs("ClassID")&"""" & SelectStrs & ">" & TempStr & TempRs("ClassCName") & "</option>"& chr(10) & chr(13)
		else
			ChildClassList = ChildClassList & "<option value="""&TempRs("ClassID")&"""" & SelectStrs & ">" & TempStr & TempRs("ClassCName") & "</option>"& chr(10) & chr(13)
		end if
		ChildClassList = ChildClassList & ChildClassList(TempRs("ClassID"),TempStr)
		TempRs.MoveNext
	loop
	TempRs.Close
	Set TempRs = Nothing
End Function
%>
<html>
<head>
<link href="../../../CSS/FS_css.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��ĿJS�޸�</title>
</head>
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<body topmargin="2" leftmargin="2">
<form action="" method="post" name="ClassJSForm">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=35 align="center" alt="����" onClick="document.ClassJSForm.submit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="����" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
          <td>&nbsp; <input name="action" type="hidden" id="action" value="add"> 
        </tr>
      </table>
	  </td>
  </tr>
</table>
  <table width="100%"  border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#E1E1E1">
    <tr bgcolor="#FFFFFF"> 
      <td width="15%" height="26">&nbsp;&nbsp;&nbsp;&nbsp;��������</td>
      <td width="35%"> 
        <input name="FileCName" type="text" id="FileCName" style="width:90%" value="<%=RsSysModObj("FileCName")%>"></td>
      <td width="15%">&nbsp;&nbsp;&nbsp;&nbsp;�ļ�����</td>
      <td width="35%"> 
        <input name="FileName" type="text" id="FileName" disabled style="width:90%" value="<%=RsSysModObj("FileName")%>"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26">&nbsp;&nbsp;&nbsp;&nbsp;��Ŀ����</td>
      <td> 
        <select name="ClassID" <%If Types = "System" then Response.Write("disabled")%> style="width:90%">
          <% =TempClassListStr %>
        </select></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;��������</td>
      <td> 
        <select name="NewsType" style="width:90%" onChange="ChooseNewsType(this.options[this.selectedIndex].value);">
          <option value="RecNews" <%if RsSysModObj("NewsType") = "RecNews" then Response.Write("selected")%>>�Ƽ�����</option>
          <option value="MarqueeNews" <%if RsSysModObj("NewsType") = "MarqueeNews" then Response.Write("selected")%>>��������</option>
          <option value="SBSNews" <%if RsSysModObj("NewsType") = "SBSNews" then Response.Write("selected")%>>��������</option>
          <option value="PicNews" <%if RsSysModObj("NewsType") = "PicNews" then Response.Write("selected")%>>ͼƬ����</option>
          <option value="NewNews" <%if RsSysModObj("NewsType") = "NewNews" then Response.Write("selected")%>>��������</option>
          <option value="HotNews" <%if RsSysModObj("NewsType") = "HotNews" then Response.Write("selected")%>>�ȵ�����</option>
          <option value="WordNews" <%if RsSysModObj("NewsType") = "WordNews" then Response.Write("selected")%>>��������</option>
          <option value="TitleNews" <%if RsSysModObj("NewsType") = "TitleNews" then Response.Write("selected")%>>��������</option>
          <option value="ProclaimNews" <%if RsSysModObj("NewsType") = "ProclaimNews" then Response.Write("selected")%>>��������</option>
        </select></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26">&nbsp;&nbsp;&nbsp;��������</td>
      <td> 
        <select name="MoreContent" id="MoreContent" style="width:90% " <%If Types = "System" then Response.Write("disabled")%> onChange="ChooseLink(this.options[this.selectedIndex].value);">
          <option value="1" <%If RsSysModObj("MoreContent")=1 then Response.Write("selected")%>>��</option>
          <option value="0" <%If RsSysModObj("MoreContent")=0 then Response.Write("selected")%>>��</option>
        </select></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;��������</td>
      <td> 
        <input name="LinkWord" type="text" id="LinkWord" style="width:90%" value="<%=RsSysModObj("LinkWord")%>" <%If Types = "System" then Response.Write("disabled")%>></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26">&nbsp;&nbsp;&nbsp;&nbsp;��������</td>
      <td> 
        <input name="NewsNum" type="text" id="NewsNum" style="width:90%" value="<%=RsSysModObj("NewsNum")%>"></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;ÿ������</td>
      <td> 
        <input name="RowNum" type="text" id="RowNum" style="width:90%" value="<%=RsSysModObj("RowNum")%>"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26">&nbsp;&nbsp;&nbsp;&nbsp;������ʽ</td>
      <td> 
        <input name="LinkCSS" type="text" id="LinkCSS" style="width:90%" value="<%=RsSysModObj("LinkCSS")%>" <%If Types = "System" then Response.Write("disabled")%>></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;��������</td>
      <td> 
        <input name="TitleNum" type="text" id="TitleNum" style="width:90%" value="<%=RsSysModObj("TitleNum")%>"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26">&nbsp;&nbsp;&nbsp;&nbsp;ͼƬ���</td>
      <td> 
        <input name="PicWidth" type="text" id="PicWidth" style="width:90%" value="<%=RsSysModObj("PicWidth")%>"></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;������ʽ</td>
      <td> 
        <input name="TitleCSS" type="text" id="TitleCSS" style="width:90%" value="<%=RsSysModObj("TitleCSS")%>"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26">&nbsp;&nbsp;&nbsp;&nbsp;ͼƬ�߶�</td>
      <td> 
        <input name="PicHeight" type="text" id="PicHeight" style="width:90%" value="<%=RsSysModObj("PicHeight")%>"></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;�����о�</td>
      <td> 
        <input name="RowSpace" type="text" id="RowSpace" style="width:90%" value="<%=RsSysModObj("RowSpace")%>"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26">&nbsp;&nbsp;&nbsp;&nbsp;�����ٶ�</td>
      <td> 
        <input name="MarSpeed" type="text" id="MarSpeed" style="width:90%" value="<%=RsSysModObj("MarSpeed")%>"></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;��������</td>
      <td> 
        <select name="MarDirection" id="MarDirection" style="width:90% ">
          <option value="up" <%If RsSysModObj("MarDirection")="up" then Response.Write("selected")%>>����</option>
          <option value="down" <%If RsSysModObj("MarDirection")="down" then Response.Write("selected")%>>����</option>
          <option value="left" <%If RsSysModObj("MarDirection")="left" then Response.Write("selected")%>>����</option>
          <option value="right" <%If RsSysModObj("MarDirection")="right" then Response.Write("selected")%>>����</option>
        </select></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26">&nbsp;&nbsp;&nbsp;&nbsp;������</td>
      <td> 
        <input name="MarWidth" type="text" id="MarWidth" style="width:90%" value="<%=RsSysModObj("MarWidth")%>"></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;����߶�</td>
      <td> 
        <input name="MarHeight" type="text" id="MarHeight" style="width:90%" value="<%=RsSysModObj("MarHeight")%>"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26">&nbsp;&nbsp;&nbsp;&nbsp;��ʾ����</td>
      <td> 
        <select name="ShowTitle" id="ShowTitle" style="width:90%">
          <option value="1" <%If RsSysModObj("ShowTitle")=1 then Response.Write("selected")%>>��</option>
          <option value="0" <%If RsSysModObj("ShowTitle")=0 then Response.Write("selected")%>>��</option>
        </select></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;�¿�����</td>
      <td> 
        <select name="OpenMode" id="OpenMode" style="width:90%">
          <option value="1" <%If RsSysModObj("OpenMode")=1 then Response.Write("selected")%>>��</option>
          <option value="0" <%If RsSysModObj("OpenMode")=0 then Response.Write("selected")%>>��</option>
        </select></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26">&nbsp;&nbsp;&nbsp;&nbsp;����ͼƬ</td>
      <td> 
        <input name="NaviPic" type="text" id="NaviPic" style="width:60%" value="<%=RsSysModObj("NaviPic")%>"> 
        <input id="PicChooseButton" type="button" name="Submit34" value="ѡ��ͼƬ" onClick="OpenWindowAndSetValue('../../FunPages/SelectPic.asp?CurrPath=/<% = UpFiles %>',550,300,window,document.ClassJSForm.NaviPic);"></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;�м�ͼƬ</td>
      <td> 
        <input name="RowBetween" type="text" id="RowBetween" style="width:52%" value="<%=RsSysModObj("RowBetween")%>"> 
        <input id="PicChooseButton" type="button" name="Submit34" value="ѡ��ͼƬ" onClick="OpenWindowAndSetValue('../../FunPages/SelectPic.asp?CurrPath=/<% = UpFiles %>',550,300,window,document.ClassJSForm.RowBetween);"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26">&nbsp;&nbsp;&nbsp;&nbsp;��������</td>
      <td> 
        <select name="DateType" id="DateType" style="width:90%">
          <option value="0">������������</option>
          <option value="1" <%if RsSysModObj("DateType") = "1" then Response.Write("selected") end if%>><%=Year(Now)&"-"&Month(Now)&"-"&Day(Now)%></option>
          <option value="2" <%if RsSysModObj("DateType") = "2" then Response.Write("selected") end if%>><%=Year(Now)&"."&Month(Now)&"."&Day(Now)%></option>
          <option value="3" <%if RsSysModObj("DateType") = "3" then Response.Write("selected") end if%>><%=Year(Now)&"/"&Month(Now)&"/"&Day(Now)%></option>
          <option value="4" <%if RsSysModObj("DateType") = "4" then Response.Write("selected") end if%>><%=Month(Now)&"/"&Day(Now)&"/"&Year(Now)%></option>
          <option value="5" <%if RsSysModObj("DateType") = "5" then Response.Write("selected") end if%>><%=Day(Now)&"/"&Month(Now)&"/"&Year(Now)%></option>
          <option value="6" <%if RsSysModObj("DateType") = "6" then Response.Write("selected") end if%>><%=Month(Now)&"-"&Day(Now)&"-"&Year(Now)%></option>
          <option value="7" <%if RsSysModObj("DateType") = "7" then Response.Write("selected") end if%>><%=Month(Now)&"."&Day(Now)&"."&Year(Now)%></option>
          <option value="8" <%if RsSysModObj("DateType") = "8" then Response.Write("selected") end if%>><%=Month(Now)&"-"&Day(Now)%></option>
          <option value="9" <%if RsSysModObj("DateType") = "9" then Response.Write("selected") end if%>><%=Month(Now)&"/"&Day(Now)%></option>
          <option value="10" <%if RsSysModObj("DateType") = "10" then Response.Write("selected") end if%>><%=Month(Now)&"."&Day(Now)%></option>
          <option value="11" <%if RsSysModObj("DateType") = "11" then Response.Write("selected") end if%>><%=Month(Now)&"��"&Day(Now)&"��"%></option>
          <option value="12" <%if RsSysModObj("DateType") = "12" then Response.Write("selected") end if%>><%=day(Now)&"��"&Hour(Now)&"ʱ"%></option>
          <option value="13" <%if RsSysModObj("DateType") = "13" then Response.Write("selected") end if%>><%=day(Now)&"��"&Hour(Now)&"��"%></option>
          <option value="14" <%if RsSysModObj("DateType") = "14" then Response.Write("selected") end if%>><%=Hour(Now)&"ʱ"&Minute(Now)&"��"%></option>
          <option value="15" <%if RsSysModObj("DateType") = "15" then Response.Write("selected") end if%>><%=Hour(Now)&":"&Minute(Now)%></option>
          <option value="16" <%if RsSysModObj("DateType") = "16" then Response.Write("selected") end if%>><%=Year(Now)&"��"&Month(Now)&"��"&Day(Now)&"��"%></option>
        </select></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;����·��</td>
      <td> 
        <input name="SaveFilePath" type="text" id="SaveFilePath" style="width:52%" value="<%=RsSysModObj("FileSavePath")%>"> 
        <input type="button" name="Subsadfmit" value="ѡ��·��" onClick="OpenWindowAndSetValue('../../FunPages/SelectPathFrame.asp?CurrPath=<%="/"&TeempSysRootDir&"JS"%>',400,300,window,document.ClassJSForm.SaveFilePath);document.ClassJSForm.SaveFilePath.focus();"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26">&nbsp;&nbsp;&nbsp;&nbsp;������ʽ</td>
      <td> 
        <input name="DateCSS" type="text" id="DateCSS" style="width:90%" value="<%=RsSysModObj("DateCSS")%>"></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;��ʾ��Ŀ</td>
      <td> 
        <select name="ClassName" id="ClassName" style="width:90%">
          <option value="1" <%If RsSysModObj("ClassName")=1 then Response.Write("selected")%>>��ʾ</option>
          <option value="0" <%If RsSysModObj("ClassName")=0 then Response.Write("selected")%>>����ʾ</option>
        </select></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26">&nbsp;&nbsp;&nbsp;&nbsp;��������</td>
      <td> 
        <select name="SonClass" id="SonClass" style="width:90%" <%If Types = "System" then Response.Write("disabled")%>>
          <option value="1" <%If RsSysModObj("SonClass")=1 then Response.Write("selected")%>>��</option>
          <option value="0" <%If RsSysModObj("SonClass")=0 then Response.Write("selected")%>>��</option>
        </select></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;�����Ҷ���</td>
      <td> 
        <select name="RightDate" id="RightDate" style="width:90%">
          <option value="1" <%If RsSysModObj("RightDate")=1 then Response.Write("selected")%>>��</option>
          <option value="0" <%If RsSysModObj("RightDate")=0 then Response.Write("selected")%>>��</option>
        </select></td>
    </tr>
</table>
</form>
</body>
</html>
<%
If Request.Form("action") = "add" then
	Dim ResultStr,TempFsoObj,FileNameStr
	FileNameStr = RsSysModObj("FileName")
	If NoCSSHackAdmin(Request.Form("FileCName"),"��������")="" then
		Response.Write("<script>alert(""�ļ��������Ʋ���Ϊ��"");</script>") '�ļ����Ʋ���Ϊ�ջ����зǷ��ַ�
		Response.End
	End If
	If Request.Form("SaveFilePath")="" then
		Response.Write("<script>alert(""δָ���ļ�����·��"");</script>") '�ļ����Ʋ���Ϊ�ջ����зǷ��ַ�
		Response.End
	End If
	If isnumeric(Request.Form("NewsNum"))=false then
		Response.Write("<script>alert(""����������������Ϊ������"");</script>") '����������������Ϊ������
		Response.End
	End If
	If isnumeric(Request.Form("TitleNum"))=false then
		Response.Write("<script>alert(""������������Ϊ������"");</script>")
		Response.End
	End If
	If isnumeric(Request.Form("RowNum"))=false then
		Response.Write("<script>alert(""����ÿ��������������Ϊ������"");</script>") 
		Response.End
	End If
	If isnumeric(Request.Form("RowSpace"))=false then
		Response.Write("<script>alert(""�����о����Ϊ������"");</script>") 
		Response.End
	End IF
	If Types="Class" and Request.Form("ClassID")="" then
		Response.Write("<script>alert(""��ĿID�������ݴ���"");</script>") '��ĿID�������ݴ���
		Response.End
	End If
	If Request.Form("NewsType")="PicNews" or Request.Form("NewsType")="FilterNews" then
		If isnumeric(Request.Form("PicWidth"))=false or isnumeric(Request.Form("PicHeight"))=false then
			Response.Write("<script>alert(""ͼƬ������Ϊ������"");</script>") 
			Response.End
		End If
	End If
	If Request.Form("MoreContent")=1 then
		If Request.Form("LinkWord")="" then
			Response.Write("<script>alert(""��������������"");</script>") 
			Response.End
		End If
	End If
	If Request.Form("NewsType")="MarqueeNews" or Request.Form("NewsType")="ProclaimNews" then
		If isnumeric(Request.Form("MarSpeed"))=false then
			Response.Write("<script>alert(""���Ź����ٶȱ���Ϊ������"");</script>") 
			Response.End
		End If
	End If
	Dim ClassJsAddObj,RsClassSql
	Set ClassJsAddObj = Server.CreateObject(G_FS_RS)
	RsClassSql = "Select * from FS_SysJs where ID="&FileID&""
	ClassJsAddObj.Open RsClassSql,Conn,3,3
	ClassJsAddObj("FileCName") = Request.Form("FileCName")
	if Types = "Class" then
		ClassJsAddObj("ClassID") = Cstr(Request.Form("ClassID"))
	end if
	ClassJsAddObj("NewsType") = Request.Form("NewsType")
	ClassJsAddObj("NewsNum") = Cint(Request.Form("NewsNum"))
	ClassJsAddObj("TitleNum") = Cint(Request.Form("TitleNum"))
	ClassJsAddObj("TitleCSS") = Cstr(Request.Form("TitleCSS"))
	ClassJsAddObj("RowNum") = Cint(Request.Form("RowNum"))
	ClassJsAddObj("NaviPic") = Request.Form("NaviPic")
	ClassJsAddObj("RowBetween") = Request.Form("RowBetween")
	ClassJsAddObj("FileSavePath") = Cstr(Request.Form("SaveFilePath"))
	ClassJsAddObj("RowSpace") = Cint(Request.Form("RowSpace"))
	ClassJsAddObj("DateType") = Cint(Request.Form("DateType"))
	ClassJsAddObj("DateCSS") = Cstr(Request.Form("DateCSS"))
	If Request.Form("ClassName")<>0 then
		ClassJsAddObj("ClassName") = 1
	Else
		ClassJsAddObj("ClassName") = 0
	End If
	If Request.Form("SonClass")<>0 then
		ClassJsAddObj("SonClass") = 1
	Else
		ClassJsAddObj("SonClass") = 0
	End If
	If Request.Form("RightDate")<>0 then
		ClassJsAddObj("RightDate") = 1
	Else
		ClassJsAddObj("RightDate") = 0
	End If
	If Request.Form("MoreContent")<>"" and isnull(Request.Form("MoreContent"))=false then
		ClassJsAddObj("MoreContent") = Request.Form("MoreContent")
	End if
	If Request.Form("MoreContent")<>0 then
		ClassJsAddObj("LinkWord") = Request.Form("LinkWord")
		ClassJsAddObj("LinkCSS") = Request.Form("LinkCSS")
	End If
	If Request.Form("PicWidth")<>"" and isnull(Request.Form("PicWidth"))=false then
		ClassJsAddObj("PicWidth") = Cint(Request.Form("PicWidth"))
	End If
	If Request.Form("PicHeight")<>"" and isnull(Request.Form("PicHeight"))=false then
		ClassJsAddObj("PicHeight") = Cint(Request.Form("PicHeight"))
	End If
	If Request.Form("MarSpeed")<>"" and isnull(Request.Form("MarSpeed"))=false then
		ClassJsAddObj("MarSpeed") = Cint(Request.Form("MarSpeed"))
	End If
	If Request.Form("MarDirection")<>"" and isnull(Request.Form("MarDirection"))=false then
		ClassJsAddObj("MarDirection") = Cstr(Request.Form("MarDirection"))
	End If
	If Request.Form("ShowTitle")<>"" and isnull(Request.Form("ShowTitle"))=false then
		ClassJsAddObj("ShowTitle") = Request.Form("ShowTitle")
	End If
	If Request.Form("OpenMode")<>1 then
		ClassJsAddObj("OpenMode") = 0
	Else
		ClassJsAddObj("OpenMode") = 1
	End If
	If Request.Form("MarWidth")<>"" and isnull(Request.Form("MarWidth"))=false then
		ClassJsAddObj("MarWidth") = Request.Form("MarWidth")
	End If
	If Request.Form("MarHeight")<>"" and isnull(Request.Form("MarHeight"))=false then
		ClassJsAddObj("MarHeight") = Request.Form("MarHeight")
	End If
	ClassJsAddObj.Update
	ClassJsAddObj.Close
	Set ClassJsAddObj = Nothing
	ResultStr = CreateSysJS(FileNameStr)
	if ResultStr = true then
		Response.Redirect("ClassJsList.asp?Types=" & Types)
	else
		Response.Write("<script>alert("""&ResultStr&""");location='ClassJsList.asp?Types=" & Types & "'</script>") 
	end if
	Response.End
End If
Conn.Close
Set Conn = Nothing
%>
<script>
var TempLink='<% = ScrLink %>';
var TempNewsType='<% = ScrNewsType %>';
function ChooseLink(Link)
{
	if (Link==1)
	{
	 document.ClassJSForm.LinkWord.disabled=false;
	 document.ClassJSForm.LinkCSS.disabled=false;
	 }
	else
	{
	 document.ClassJSForm.LinkWord.disabled=true;
	 document.ClassJSForm.LinkCSS.disabled=true;
	 }
 }

function ChooseNewsType(NewsType)
{
 if ((NewsType=='MarqueeNews')||(NewsType=='ProclaimNews'))
  {
	 document.ClassJSForm.MarSpeed.disabled=false;
	 document.ClassJSForm.MarDirection.disabled=false;
	 document.ClassJSForm.MarWidth.disabled=false;
	 document.ClassJSForm.MarHeight.disabled=false;
   }
  else
  {
	 document.ClassJSForm.MarSpeed.disabled=true;
	 document.ClassJSForm.MarDirection.disabled=true;
	 document.ClassJSForm.MarWidth.disabled=true;
	 document.ClassJSForm.MarHeight.disabled=true;
   }
  if ((NewsType=='PicNews')||(NewsType=='FilterNews'))
  {
	 document.ClassJSForm.PicWidth.disabled=false;
	 document.ClassJSForm.PicHeight.disabled=false;
	 document.ClassJSForm.ShowTitle.disabled=false;
   }
  else
  {
	 document.ClassJSForm.PicWidth.disabled=true;
	 document.ClassJSForm.PicHeight.disabled=true;
	 document.ClassJSForm.ShowTitle.disabled=true;
   }
 }
 ChooseLink(TempLink);
 ChooseNewsType(TempNewsType);
</script>