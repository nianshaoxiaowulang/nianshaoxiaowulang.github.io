<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<!--#include file="../Inc/Cls_JS.asp" -->
<%
'==============================================================================
'������ƣ���Ѷ��վ��Ϣ����ϵͳ
'��ǰ�汾��Foosun Content Manager System v3.1 
'���¸��£�2004.12
'==============================================================================
'��ҵע����ϵ��028-85098980-601,602 ����֧�֣�028-85098980-606��607,�ͻ�֧�֣�608
'��Ʒ��ѯQQ��159410,655071,66252421
'����֧��:���г���ʹ�����⣬�����ʵ�bbs.foosun.net���ǽ���ʱ�ش���
'���򿪷�����Ѷ������ & ��Ѷ���������
'��̳֧�֣���Ѷ������̳(http://bbs.foosun.net)
'�ٷ���վ��www.Foosun.net  ��ʾվ�㣺test.cooin.com    
'��վ����ר����www.cooin.com
'==============================================================================
'��Ѱ汾����������ҳ������Ȩ��Ϣ�������ϱ�վLOGO��������
'==============================================================================
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->

<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P060401") then Call ReturnError()
Dim TempSysRootDir
if SysRootDir = "" then
	TempSysRootDir = ""
else
	TempSysRootDir = "/" & SysRootDir
end if

dim FileID,FileObj,Types,FreeJSObj
if Request("JSID")<>"" then
	FileID = Request("JSID")
	Types = Cstr(Request("Types"))
else
	 Response.Write("<script>alert(""�������ݴ���"");</script>")
	 response.end
end if
%>
<html>
<head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>����JS����ͼƬ�޸�</title>
</head>
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<body leftmargin="0" topmargin="0">
<% if Types = "Mod" then
	Set FileObj = Conn.Execute("Select PicPath,JSName from FS_FreeJsFile where ID="&Clng(FileID)&"")
	if FileObj.eof then
		 Response.Write("<script>alert(""�������ݴ���"");</script>")
		 response.end
	end if
	Set FreeJSObj = Conn.Execute("Select Manner from FS_FreeJS where EName='"&FileObj("JSName")&"'")
%>
<form action="" name="FMForm" method="post" >
  <table width="100%" border="0" cellspacing="5" cellpadding="0">
    <tr> 
      <td height="8"></td>
    </tr>
    <tr> 
      <td><div align="center">
          <input name="PicPath" type="text" size="30" value="<%=FileObj("PicPath")%>">
          <input type="button" name="Submit" value="ѡ��ͼƬ" onClick="OpenWindowAndSetValue('../inc/SelectPic.asp?CurrPath=/<% = UpFiles %>',550,300,window,document.FMForm.PicPath);">
      </div></td>
    </tr>
    <tr> 
      <td height="5"></td>
    </tr>
    <tr> 
      <td> <div align="center">
          <input type="submit" name="Submit2" value=" �� �� ">
          <input name="action" type="hidden" id="action2" value="mod">
          <input type="button" name="Submit3" value=" ȡ �� " onClick="window.close();">
      </div></td>
    </tr>
  </table>
</form>
<% elseif Types = "Del" then%>
<form name="JSDellForm" action="" method="post">
  <table width="100%" border="0" cellspacing="5" cellpadding="0">
    <tr> 
      <td><div align="center"><img src="../../Images/Question.gif" width="39" height="37"></div></td>
      <td width="76%">��ȷ��Ҫ��JS��ɾ����ѡ����?</td>
    </tr>
    <tr> 
      <td colspan="2"><div align="center"> 
          <input type="submit" name="Submit4" value=" ȷ �� ">
          <input name="DAction" type="hidden" id="DAction" value="trues">
          <input type="button" name="Submit5" value=" ȡ �� " onclick="window.close();">
      </div></td>
    </tr>
  </table>
</form>
  <% end if %>
</body>
</html>
<%
Dim JSClassObj,ReturnValue
Set JSClassObj = New JSClass
JSClassObj.SysRootDir = TempSysRootDir
if Request("action")="mod" then
     if Request.Form("PicPath")<>"" then
	     dim PicStr
		 PicStr = Cstr(Request.Form("PicPath"))
		 Conn.Execute("Update FS_FreeJsFile set PicPath='"&PicStr&"' where ID="&FileID&"")
	  '--------------------��������JS�ļ�---------------------------------
		  Select case FreeJSObj("Manner")
			 case "1"   ReturnValue = JSClassObj.WCssA(FileObj("JSName"),True)
			 case "2"   ReturnValue = JSClassObj.WCssB(FileObj("JSName"),True)
			 case "3"   ReturnValue = JSClassObj.WCssC(FileObj("JSName"),True)
			 case "4"   ReturnValue = JSClassObj.WCssD(FileObj("JSName"),True)
			 case "5"   ReturnValue = JSClassObj.WCssE(FileObj("JSName"),True)
			 case "6"   ReturnValue = JSClassObj.PCssA(FileObj("JSName"),True)
			 case "7"   ReturnValue = JSClassObj.PCssB(FileObj("JSName"),True)
			 case "8"   ReturnValue = JSClassObj.PCssC(FileObj("JSName"),True)
			 case "9"   ReturnValue = JSClassObj.PCssD(FileObj("JSName"),True)
			 case "10"   ReturnValue = JSClassObj.PCssE(FileObj("JSName"),True)
			 case "11"   ReturnValue = JSClassObj.PCssF(FileObj("JSName"),True)
			 case "12"  ReturnValue = JSClassObj.PCssG(FileObj("JSName"),True)
			 case "13"   ReturnValue = JSClassObj.PCssH(FileObj("JSName"),True)
			 case "14"   ReturnValue = JSClassObj.PCssI(FileObj("JSName"),True)
			 case "15"   ReturnValue = JSClassObj.PCssJ(FileObj("JSName"),True)
			 case "16"   ReturnValue = JSClassObj.PCssK(FileObj("JSName"),True)
			 case "17"   ReturnValue = JSClassObj.PCssL(FileObj("JSName"),True)
	   End Select
	  FreeJSObj.Close
	  Set FreeJSObj = Nothing 
	 end if
	If ReturnValue <> "" then
		Response.write("<script>alert('" & ReturnValue & "');dialogArguments.location.reload();window.close();</script>")
	else
		Response.write("<script>dialogArguments.location.reload();window.close();</script>")
	End If
  end if
  
  if Request.Form("DAction")="trues" then
  	Dim DFArray,DF_i,CreaFileObj,CcreFileObj
	DFArray = Array("")
	DFArray = Split(FileID,"***")
	FileID = Replace(FileID,"***",",")
	For DF_i = 0 to UBound(DFArray)
		 Conn.Execute("update FS_FreeJsFile set DelFlag=1 where ID="&DFArray(DF_i)&"")
	Next
	Set CreaFileObj = Conn.Execute("Select distinct JSName from FS_FreeJsFile where ID in ("&FileID&")")
	Do while Not CreaFileObj.eof
		Set CcreFileObj = Conn.Execute("Select Manner,EName from FS_FreeJS where EName='"&CreaFileObj("JSName")&"'")
		If Not CcreFileObj.eof then
	  '--------------------��������JS�ļ�---------------------------------
		  Select case CcreFileObj("Manner")
			 case "1"   ReturnValue = JSClassObj.WCssA(CreaFileObj("JSName"),True)
			 case "2"   ReturnValue = JSClassObj.WCssB(CreaFileObj("JSName"),True)
			 case "3"   ReturnValue = JSClassObj.WCssC(CreaFileObj("JSName"),True)
			 case "4"   ReturnValue = JSClassObj.WCssD(CreaFileObj("JSName"),True)
			 case "5"   ReturnValue = JSClassObj.WCssE(CreaFileObj("JSName"),True)
			 case "6"   ReturnValue = JSClassObj.PCssA(CreaFileObj("JSName"),True)
			 case "7"   ReturnValue = JSClassObj.PCssB(CreaFileObj("JSName"),True)
			 case "8"   ReturnValue = JSClassObj.PCssC(CreaFileObj("JSName"),True)
			 case "9"   ReturnValue = JSClassObj.PCssD(CreaFileObj("JSName"),True)
			 case "10"   ReturnValue = JSClassObj.PCssE(CreaFileObj("JSName"),True)
			 case "11"   ReturnValue = JSClassObj.PCssF(CreaFileObj("JSName"),True)
			 case "12"   ReturnValue = JSClassObj.PCssG(CreaFileObj("JSName"),True)
			 case "13"   ReturnValue = JSClassObj.PCssH(CreaFileObj("JSName"),True)
			 case "14"   ReturnValue = JSClassObj.PCssI(CreaFileObj("JSName"),True)
			 case "15"   ReturnValue = JSClassObj.PCssJ(CreaFileObj("JSName"),True)
			 case "16"   ReturnValue = JSClassObj.PCssK(CreaFileObj("JSName"),True)
			 case "17"   ReturnValue = JSClassObj.PCssL(CreaFileObj("JSName"),True)
	   End Select
	  End If
	  CcreFileObj.Close
	  Set CcreFileObj = Nothing
	  CreaFileObj.MoveNext
	Loop
	For DF_i = 0 to UBound(DFArray)
		 Conn.Execute("delete from FS_FreeJsFile where ID="&DFArray(DF_i)&"")
	Next
	If ReturnValue <> "" then
		Response.write("<script>alert('" & ReturnValue & "');dialogArguments.location.reload();window.close();</script>")
	else
		Response.write("<script>dialogArguments.location.reload();window.close();</script>")
	End If
 end if
Set JSClassObj = Nothing
%>