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
'软件名称：FoosunHelp System Form FoosunCMS
'当前版本：Foosun Content Manager System 3.0 系列
'最新更新：2004.12
'==============================================================================
'商业注册联系：028-85098980-601,602 技术支持：028-85098980-605、607,客户支持：608
'产品咨询QQ：159410,394226379,125114015,655071
'技术支持:所有程序使用问题，请提问到bbs.foosun.net我们将及时回答您
'程序开发：风讯开发组 & 风讯插件开发组
'论坛支持：风讯在线论坛(http://bbs.foosun.net)
'官方网站：www.Foosun.net  演示站点：test.cooin.com    
'网站建设专区：www.cooin.com
'==============================================================================
'免费版本请在新闻首页保留版权信息，并做上本站LOGO友情连接
'==============================================================================
Dim Action,HelpID
Dim FuncName,FileName,PageField,HelpContent,HelpSingleContent

Action = Lcase(Request.Form("Action"))
HelpID = Request.Form("HelpID")
'过滤需要处理,暂略
FuncName = replace(Request.Form("FuncName")," ","")
FileName = replace(Request.Form("FileName")," ","")
PageField = Request.Form("NewPageField")
HelpContent = Request.Form("HelpContent")
HelpSingleContent = Request.Form("HelpSingleContent")



Dim strErrMsg
If FuncName="" Then strErrMsg = strErrMsg & "页面功能没有数据\n"
If FileName="" or (Instr(Lcase(FileName),".asp")=0 and Instr(Lcase(FileName),".htm")=0) Then strErrMsg = strErrMsg & "页面地址不正确\n"
If PageField="" or Len(PageField)=1 Then strErrMsg = strErrMsg & "关键字数据错误\n"

If HelpContent="" Then strErrMsg = strErrMsg & "没有具体的帮助信息数据\n"
If HelpSingleContent="" Then strErrMsg = strErrMsg & "没有非常概要的的帮助信息数据\n"

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

'操作后的提示信息
If strErrMsg="" Then
	Response.write "<script language=javascript>alert('操作成功');location='SearchManage.asp?FileName="&FileName&"&FuncName="&FuncName&"';</script>"
	Response.end
End If
Response.write "<script language=javascript>alert('"&strErrMsg&"');location='SearchManage.asp?FileName="&FileName&"&FuncName="&FuncName&"';</script>"

Sub AddNew()
	if Not JudgePopedomTF(Session("Name"),"P070801") then Call ReturnError1()
	strSQL = "select * From [Fs_Help] Where FuncName='"&Replace(FuncName,"'","''")&"' and FileName='"& Replace(FileName,"'","''") &"' and PageField='"&Replace(PageField,"'","''")&"'"
	HelpRs.open strSQL,HelpConn,1,3
	If not HelpRs.eof Then
		strErrMsg = strErrMsg & "“"&PageField&"”的帮助已经存在\n"
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
		strErrMsg = strErrMsg & "“" & HelpID & "”的帮助没有找到\n"
		Exit Sub
	End If

	strSQL = "select * From [Fs_Help] Where ID="& HelpID &""
	HelpRs.open strSQL,HelpConn,1,3
	If HelpRs.eof Then
		strErrMsg = strErrMsg & "“"&PageField&"”的帮助不存在\n"
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
