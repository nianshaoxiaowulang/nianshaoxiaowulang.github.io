<% Option Explicit %>
<!--#include file="../Inc/Cls_DB.asp" -->
<!--#include file="../Inc/Const.asp" -->
<!--#include file="../Inc/Md5.asp" -->
<!--#include file="../Inc/Function.asp" -->
<%
	Dim DBC,conn
	Set DBC = new databaseclass
	Set conn = DBC.openconnection()
	Set DBC = nothing
	Dim I,RsConfigObj
	Set RsConfigObj = Conn.Execute("Select SiteName,Copyright from FS_Config")
%>
<HTML><HEAD><TITLE><%=RsConfigObj("SiteName")%> >> 会员登陆</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<META content="MSHTML 6.00.2600.0" name=GENERATOR>
<META content=JavaScript name=vs_defaultClientScript>
<LINK href="Css/UserCSS.css" type=text/css rel=stylesheet>
</HEAD>
<script language="javascript" src="Comm/MyScript.js"></script>
<BODY leftMargin=0 topMargin=10 MS_POSITIONING="GridLayout">
<div align="center">
  <script language="JavaScript" src="top.js" type="text/JavaScript"></script>
</div>
<TABLE cellSpacing=2 width="98%" align=center border=0>
  <TBODY>
    <TR> 
      <TD vAlign=top> <TABLE cellSpacing=0 cellPadding=0 width="98%" align=center 
                  border=0>
          <TBODY>
            <TR> 
              <TD width="100%"> <TABLE width="100%" border=0>
                  <TBODY>
                    <TR> 
                      <TD width=26><IMG 
                              src="images/Favorite.OnArrow.gif" border=0></TD>
                      <TD 
class=f4>会员登陆</TD>
                    </TR>
                  </TBODY>
                </TABLE></TD>
            </TR>
            <TR> 
              <TD width="100%"> <TABLE cellSpacing=0 cellPadding=0 width="100%" 
border=0>
                  <TBODY>
                    <TR> 
                      <TD bgColor=#ff6633 height=4><IMG height=1 src="" 
                              width=1></TD>
                    </TR>
                  </TBODY>
                </TABLE></TD>
            </TR>
            <TR> 
              <TD width="100%" height="37"> <div align="left"> 
                  <form name=LoginForm method="post" action="CheckLogin.asp" onSubmit="return CheckLogindata()">
                    <br>
                    <table width="316" border="0" align="center" cellpadding="6" cellspacing="0">
                      <tr class="txt"> 
                        <td> <div align="center" class="td"><font color="#000000">用户名</font> 
                            <input name="MemName" type="text" class="input" id="MemName" style="CURSOR: hand">
                          </div></td>
                      </tr>
                      <tr class="txt"> 
                        <td> <div align="center" class="td"><font color="#000000">密　码</font> 
                            <input name="Password" type="password" class="input" id="Password" style="CURSOR: hand">
                          </div></td>
                      </tr>
                      <tr> 
                        <td><div align="center"> 
                            <input type="submit" name="Submit" value="登陆">
                            　 
                            <input type="reset" name="Submit2" value="重置">
                            <input name="Url" type="hidden" id="Url" value="<%=Request("Url")%>">
                          </div></td>
                      </tr>
                      <tr> 
                        <td><div align="center"><a href="Register.asp" class="txt">免费注册</a> 
                            　<a href="User_GetPassword.asp" class="txt">忘记密码</a></div></td>
                      </tr>
                    </table>
                    <div align="center"></div>
                  </form>
                </div></TD>
            </TR>
          </TBODY>
        </TABLE></TD>
    </TR>
  </TBODY>
</TABLE>
  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="0">
  <tr> 
    <td> <hr size="1" noshade color="#FF6600"> <div align="center"> 
        <% = RsConfigObj("Copyright") %>
      </div></td>
  </tr>
</table>
</BODY></HTML>
