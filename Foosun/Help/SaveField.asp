<% Option Explicit %>
<!--#include file="../../Inc/Cls_DB.asp" -->
<!--#include file="../../Inc/Const.asp" -->
<%
Dim DBC,Conn,HelpConn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
DBC.ConnStr = "DBQ=" + Server.MapPath("Foosun_help.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
Set HelpConn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../Inc/Session.asp" -->
<!--#include file="../../Inc/CheckPopedom.asp" -->
<%
'==============================================================================
'������ƣ�FoosunHelp System Form FoosunCMS
'��ǰ�汾��Foosun Content Manager System 3.0 ϵ��
'���¸��£�2004.12
'==============================================================================
'��ҵע����ϵ��028-85098980-601,602 ����֧�֣�028-85098980-605��607,�ͻ�֧�֣�608
'��Ʒ��ѯQQ��159410,394226379,125114015,655071
'����֧��:���г���ʹ�����⣬�����ʵ�bbs.foosun.net���ǽ���ʱ�ش���
'���򿪷�����Ѷ������ & ��Ѷ���������
'��̳֧�֣���Ѷ������̳(http://bbs.foosun.net)
'�ٷ���վ��www.Foosun.net  ��ʾվ�㣺test.cooin.com    
'��վ����ר����www.cooin.com
'==============================================================================
'��Ѱ汾����������ҳ������Ȩ��Ϣ�������ϱ�վLOGO��������
'==============================================================================
Dim Action,HelpID
Dim FuncName,FileName,PageField,HelpContent,HelpSingleContent

Action = Lcase(Request.Form("Action"))
HelpID = Request.Form("HelpID")
'������Ҫ����,����
FuncName = replace(Request.Form("FuncName")," ","")
FileName = replace(Request.Form("FileName")," ","")
PageField = Request.Form("NewPageField")
HelpContent = Request.Form("HelpContent")
HelpSingleContent = Request.Form("HelpSingleContent")



Dim strErrMsg
If FuncName="" Then strErrMsg = strErrMsg & "ҳ�湦��û������\n"
If FileName="" or (Instr(Lcase(FileName),".asp")=0 and Instr(Lcase(FileName),".htm")=0) Then strErrMsg = strErrMsg & "ҳ���ַ����ȷ\n"
If PageField="" or Len(PageField)=1 Then strErrMsg = strErrMsg & "�ؼ������ݴ���\n"

If HelpContent="" Then strErrMsg = strErrMsg & "û�о���İ�����Ϣ����\n"
If HelpSingleContent="" Then strErrMsg = strErrMsg & "û�зǳ���Ҫ�ĵİ�����Ϣ����\n"

If strErrMsg<>"" Then
	Response.write "<script language=javascript>alert('"&strErrMsg&"');history.back();</script>"
	Response.end
End If

if session("FuncName") <> FuncName then session("FuncName") = FuncName
if session("FileName") <> FileName then session("FileName") = FileName

Dim strSQL,HelpRs,PageFieldArray,iTemp
Set HelpRs = Server.CreateObject(G_FS_RS)

Select Case Action
	Case "addnew" : Call AddNew()
	Case "modify" : Call ModiHelp()
End Select
set Conn = Nothing

'���������ʾ��Ϣ
If strErrMsg="" Then
	Response.write "<script language=javascript>alert('�����ɹ�');location='SearchManage.asp?FileName="&FileName&"&FuncName="&FuncName&"';</script>"
	Response.end
End If
Response.write "<script language=javascript>alert('"&strErrMsg&"');location='SearchManage.asp?FileName="&FileName&"&FuncName="&FuncName&"';</script>"

Sub AddNew()
	if Not JudgePopedomTF(Session("Name"),"P070801") then Call ReturnError1()
	strSQL = "select * From [Fs_Help] Where FuncName='"&Replace(FuncName,"'","''")&"' and FileName='"& Replace(FileName,"'","''") &"' and PageField='"&Replace(PageField,"'","''")&"'"
	HelpRs.open strSQL,HelpConn,1,3
	If not HelpRs.eof Then
		strErrMsg = strErrMsg & "��"&PageField&"���İ����Ѿ�����\n"
		Exit Sub
	End If
	HelpRs.addnew
	HelpRs("FuncName") = FuncName
	HelpRs("FileName") = FileName
	HelpRs("PageField") = PageField
	HelpRs("HelpContent") = HelpContent
	HelpRs("HelpSingleContent") = HelpSingleContent
	HelpRs("SvTime") = now
	HelpRs.Update
	HelpRs.Close
End Sub

Sub ModiHelp()
	if Not JudgePopedomTF(Session("Name"),"P070802") then Call ReturnError1()
	If HelpID=0 Then
		strErrMsg = strErrMsg & "��" & HelpID & "���İ���û���ҵ�\n"
		Exit Sub
	End If

	strSQL = "select * From [Fs_Help] Where ID="& HelpID &""
	HelpRs.open strSQL,HelpConn,1,3
	If HelpRs.eof Then
		strErrMsg = strErrMsg & "��"&PageField&"���İ���������\n"
		Exit Sub
	End If
	HelpRs("FuncName") = FuncName
	HelpRs("FileName") = FileName
	HelpRs("PageField") = PageField
	HelpRs("HelpContent") = HelpContent
	HelpRs("HelpSingleContent") = HelpSingleContent
	HelpRs("SvTime") = now
	HelpRs.Update
	HelpRs.Close
End Sub
Set HelpConn = nothing
%>
