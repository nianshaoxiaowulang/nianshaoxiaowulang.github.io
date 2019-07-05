<%
'**************************************
'**    Inc_menu.asp
'**
'** 文件说明：留言本导航栏
'** 修改日期：2005-04-07
'**************************************
%>

<table border="0" cellpadding="5" cellspacing="0" style="border-collapse: collapse" width="100%" id="AutoNumber1">
  <tr>
    <td width="100%" align="right" height="36" bgcolor='#F1F1F1'><IMG SRC="images/blank.gif" WIDTH="100" HEIGHT="36"></td>
  </tr>
  <tr>
    <td id="menu1" width="100%" align="right" height="25"<%if pagename="查看留言" or pagename="管理留言" or pagename="批量管理留言" then%> bgcolor='<%=maincolor%>'<%else%> onMouseOver="document.all.menu1.bgColor='<%=maincolor%>'" onMouseOut="document.all.menu1.bgColor='#F1F1F1'"<%end if%>><a href="index.asp"><B>查看留言</B></a></td>
  </tr>
  <tr>
    <td id="menu2" width="100%" align="right" height="25"<%if pagename="写留言" then%> bgcolor='<%=maincolor%>'<%else%> onMouseOver="document.all.menu2.bgColor='<%=maincolor%>'" onMouseOut="document.all.menu2.bgColor='#F1F1F1'"<%end if%>><a href="new.asp"><B>写留言</B></a></td>
  </tr>
  <tr>
    <td id="menu3" width="100%" align="right" height="25"<%if pagename="搜索留言" or pagename="搜索结果" then%> bgcolor='<%=maincolor%>'<%else%> onMouseOver="document.all.menu3.bgColor='<%=maincolor%>'" onMouseOut="document.all.menu3.bgColor='#F1F1F1'"<%end if%>
    ><a href="search.asp?act=fillform"><B>搜索留言</B></a></td>
  </tr><tr>
    <td id="menu4" width="100%" align="right" height="25"<%if pagename="留言帮助" then%> bgcolor='<%=maincolor%>'<%else%> onMouseOver="document.all.menu4.bgColor='<%=maincolor%>'" onMouseOut="document.all.menu4.bgColor='#F1F1F1'"<%end if%>><a href="help.asp"><B>帮助</B></a></td>
  </tr>
</table>