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
'�ж�Ȩ��
if Not JudgePopedomTF(Session("Name"),"P010500") then Call ReturnError1()
'�ж�Ȩ�޽���
Dim Action,NewsIDStr,PicNewsTF,RecTF,TodayNewsTF,MarqueeNews,SBSNews,ReviewTF,FilterNews,FocusNewsTF,ClassicalNewsTF,Sql
Dim tempsql,RstempObj,tempstr,tempstrnum1,tempstrnum2,i '�������������ʱ����
Action = Request("Action")
if Action = "Submit" then
	NewsIDStr = Request("NewsIDStr")
	if NewsIDStr <> "" then
		NewsIDStr = Replace(NewsIDStr,"***","','")
		picnewsTF = Request("picnewsTF")
		if picnewsTF = "1" then
		'''''''''''''''''''''''''''''''''''''''''''''''
            tempSql = "Select * from fs_News where newsID in ('" & NewsIDStr & "')"
            Set RstempObj = Server.CreateObject("ADODB.RecordSet")
            RstempObj.Open tempSql,Conn,3,3
			for i = 1 to RstempObj.RecordCount
            		if RstempObj.Eof then Exit For
	            	 tempstr=LCase(RstempObj("Content"))
                     tempstrnum1=instr(tempstr,"src=")
					 if tempstrnum1 >0 then 
					     RstempObj("picnewsTF") = 1
						 tempstr=mid(tempstr,tempstrnum1+5,100)
						 tempstrnum2=InStr(tempstr,"""")
					     RstempObj("PicPath") =left(tempstr,tempstrnum2-1)
				         RstempObj.update
				    end if	
		            RstempObj.MoveNext
            next
 			 RsTempObj.Close
            Set RsTempObj = Nothing	
            '����������ӽ���
		else
			picnewsTF = 0
		end if
		Conn.Execute("Update FS_News set picnewsTF =" & picnewsTF & " where newsID in ('" & NewsIDStr & "')")
		RecTF = Request("RecTF")
		if RecTF = "1" then
			RecTF = 1
		else
			RecTF = 0
		end if
		Conn.Execute("Update FS_News set RecTF=" & RecTF & " where newsID in ('" & NewsIDStr & "')")
		TodayNewsTF = Request("TodayNewsTF")
		if TodayNewsTF = "1" then
			TodayNewsTF = 1
		else
			TodayNewsTF = 0
		end if
		Conn.Execute("Update FS_News set TodayNewsTF=" & TodayNewsTF & " where newsID in ('" & NewsIDStr & "')")
		MarqueeNews = Request("MarqueeNews")
		if MarqueeNews = "1" then
			MarqueeNews = 1
		else
			MarqueeNews = 0
		end if
		Conn.Execute("Update FS_News set MarqueeNews=" & MarqueeNews & " where newsID in ('" & NewsIDStr & "')")
		SBSNews = Request("SBSNews")
		if SBSNews = "1" then
			SBSNews = 1
		else
			SBSNews = 0
		end if
		Conn.Execute("Update FS_News set SBSNews=" & SBSNews & " where newsID in ('" & NewsIDStr & "')")
		ReviewTF = Request("ReviewTF")
		if ReviewTF = "1" then
			ReviewTF = 1
		else
			ReviewTF = 0
		end if
		Conn.Execute("Update FS_News set ReviewTF = " & ReviewTF & " where newsID in ('" & NewsIDStr & "')")
		FilterNews = Request("FilterNews")
		if FilterNews = "1"  and picnewsTF="1" then
			FilterNews = 1
		else
			FilterNews = 0
		end if
		Conn.Execute("Update FS_News set FilterNews = " & FilterNews & " where newsID in ('" & NewsIDStr & "')")
		FocusNewsTF = Request("FocusNewsTF")
		if FocusNewsTF = "1" and picnewsTf="1" then
			FocusNewsTF = 1
		else
			FocusNewsTF = 0
		end if
		Conn.Execute("Update FS_News set FocusNewsTF = " & FocusNewsTF & " where newsID in ('" & NewsIDStr & "')")
		ClassicalNewsTF = Request("ClassicalNewsTF")
		if ClassicalNewsTF = "1" and picnewsTf="1" then
			ClassicalNewsTF = 1
		else
			ClassicalNewsTF = 0
		end if
		
		Conn.Execute("Update FS_News set ClassicalNewsTF =" & ClassicalNewsTF & " where newsID in ('" & NewsIDStr & "')")
	end if
	Set Conn = Nothing
	Response.Write("<script>dialogArguments.location.reload();window.close();</script>")
	Response.End
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>������������</title>
</head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet">
<body topmargin="0" leftmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
 <form name="SetForm" action="" method="post">
  <tr> 
      <td height="26"><div align="center">����ѡ����������
          <input type="hidden" name="NewsIDStr" value="<% = Request("NewsIDStr") %>">
          <input type="hidden" name="Action" value="Submit">
        </div></td>
  </tr>
  <tr> 
    <td height="36"> 
      <div align="center"> 
        <input name="PicNewstf" type="checkbox" id="PicNewsTF" value="1">
        ͼƬ���� 
        <input name="RecTF" type="checkbox" id="RecTF" value="1">
        �Ƽ����� 
        <input name="TodayNewsTF" type="checkbox" id="TodayNewsTF" value="1">
        ����ͷ��</div></td>
  </tr>
  <tr> 
    <td height="36"> 
      <div align="center"> 
        <input name="MarqueeNews" type="checkbox" id="MarqueeNews" value="1">
        �������� 
        <input name="SBSNews" type="checkbox" id="SBSNews" value="1">
        �������� 
        <input name="ReviewTF" type="checkbox" id="ReviewTF" value="1">
        ��������</div></td>
  </tr>
    <tr> 
    <td height="36"> 
      <div align="center"> 
        <input name="FilterNews" type="checkbox" id="FilterNews" value="1">
        �õ����� 
        <input name="FocusNewsTF" type="checkbox" id="FocusNewsTF" value="1">
        �������� 
        <input name="ClassicalNewsTF" type="checkbox" id="ClassicalNewsTF" value="1">
        ��������</div></td>
  </tr>
      <tr> 
    <td height="36" align="center"> 
      <div align="center"><font color="#ff0000"><br>&nbsp;&nbsp;&nbsp;ע�������ÿ�δ�����������ÿ���������ʱ��ԭ�����ù�������Ҫ����������һ�Σ���Ȼ�ͻᶪʧԭ�������ԡ�������<font color="#003399"><strong>�õ����š��������š���������</strong></font>ʱһ��Ҫ��ѡ��<strong>ͼƬ����</strong>���ԡ�</font></div></td>
  </tr>
  <tr> 
    <td height="46" colspan="2">
<div align="center"> 
          <input name="Submitfgsfd" type="submit" id="Submitfgsfd" value=" ȷ �� ">
        &nbsp;&nbsp;&nbsp;&nbsp; 
        <input name="Submit2fasd" type="button" id="Submit2fasd" onClick="window.close();" value=" ȡ �� ">
      </div></td>
  </tr>
 </form>
</table>
</body>
</html>
<%
Set Conn = Nothing
%>