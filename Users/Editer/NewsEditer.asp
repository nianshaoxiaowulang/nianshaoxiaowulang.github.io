<% Option Explicit %>
<!--#include file="../../Inc/Cls_DB.asp" -->
<!--#include file="../../Inc/Const.asp" -->
<%
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
Dim LimitUpFileFlag
LimitUpFileFlag = Request("LimitUpFileFlag")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>可视编辑器</title>
</head>
<link rel="stylesheet" href="Editer.css">
<script language="JavaScript" src="Editer.js"></script>
<script language="JavaScript" src="PublicJS.js"></script>
<body onLoad="return LoadEditFile();">
<table height="120" id="Toolbar" width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td class="ToolSet">
<table height="30" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="30" align="center"><img src="image/Code.gif" width="24" height="24" class="Btn" title="隐藏代码窗口" onClick="DisplayCodeWindow();" ></td>
          <td width="30" align="center"><img src="image/undo.gif" class="Btn" title="撤消" onClick="Format('undo')" ></td>
          <td width="30" align="center"><img src="image/redo.gif" class="Btn" title="恢复" onClick="Format('redo')" ></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
          <td width="30" align="center"><img src="image/NewDoc.gif" width="23" height="22" class="Btn" title="新建文档" onClick="NewPage()" ></td>
          <td align="center"> <select onFocus="SaveCurrPage();" onChange="ChangePage(this.value);" name="PageNumSelect">
              <option value="1" selected>1</option>
            </select> </td>
          <td width="30" align="center"><img onClick="DeletePage();" alt="删除文档" class="Btn" src="image/DelDoc.gif" width="23" height="22"></td>
          <td width="1" align="center"> <div align="center" class="ToolSeparator"></div></td>
          <td width="30" align="center"><img src="image/find.gif" class="Btn" title="查找 / 替换" onClick="SearchStr();" ></td>
          <td width="30" align="center"><img src="image/calculator.gif" class="Btn" title="计算器" onClick="Calculator()" ></td>
          <td width="30" align="center"><img title="插入当前日期" onClick="InsertDate()" class="Btn" src="image/date.gif" ></td>
          <td width="30" align="center"><img title="插入当前时间" onClick="InsertTime()" class="Btn" src="image/time.gif" ></td>
          <td width="30" align="center"><img title="删除所有HTML标识" onClick="DelAllHtmlTag()" class="Btn" src="image/geshi.gif" ></td>
          <td width="30" align="center"><img title="删除文字格式" onClick="Format('removeformat')" class="Btn" src="image/clear.gif" ></td>
          <td width="1" align="center"> <div align="center" class="ToolSeparator"></div></td>
          <td width="30" align="center"><img title="插入超级连接" onClick="Format('CreateLink')" class="Btn" src="image/url.gif" ></td>
          <td width="30" align="center"><img title="取消超级链接" onClick="Format('unLink')" class="Btn" src="image/nourl.gif" ></td>
          <td width="30" align="center"><img title="插入网页" onClick="InsertPage()" class="Btn" src="image/htm.gif" ></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
          <td width="30" align="center"><img title="插入栏目框" onClick="InsertFrame()" class="Btn" src="image/fieldset.gif" ></td>
          <td width="30" align="center"><img title="插入Excel表格" onClick="InsertExcel()" class="Btn" src="image/excel.gif" ></td>
          <td width="30" align="center"><img title="插入滚动文本" onClick="InsertMarquee()" class="Btn" src="image/Marquee.gif" ></td>
          <td width="30" align="center"><img title="插入图片，支持格式为：jpg、gif、bmp、png等" onClick="InsertPicture('<% = LimitUpFileFlag %>')" class="Btn" src="image/img.gif" ></td>
          <td width="30" align="center"><img title="插入flash多媒体文件" onClick="InsertFlash('<% = LimitUpFileFlag %>')" class="Btn" src="image/flash.gif" ></td>
          <td width="30" align="center"><img title="插入视频文件，支持格式为：avi、wmv、asf、mpg" onClick="InsertVideo('<% = LimitUpFileFlag %>')" class="Btn" src="image/wmv.gif" ></td>
          <td width="30" align="center"><img title="插入RealPlay文件，支持格式为：rm、ra、ram" onClick="InsertRM('<% = LimitUpFileFlag %>')" class="Btn" src="image/rm.gif" ></td>
          <td width="30" align="center"><img src="image/PicAlign.gif" width="23" height="22" class="Btn" title="图文并排" onClick="PicAndTextArrange()" ></td>
		</tr>
      </table></td>
  </tr>
  <tr> 
    <td height="30" class="ToolSet"> 
      <table height="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="26" align="center"><img src="Image/Inserttable.gif"  class="Btn" title="插入表格" onClick="InsertTable()"></td>
          <td width="26" align="center"><img src="Image/inserttable1.gif" width="23" height="22"  class="Btn" title="插入行" onClick="InsertRow()"></td>
          <td width="26" align="center"><img src="Image/delinserttable1.gif" width="23" height="22"  class="Btn" title="删除行" onClick="DeleteRow()"></td>
		  <td width="26" align="center"><img src="Image/inserttablec.gif" width="23" height="22"  class="Btn" title="插入列" onClick="InsertColumn()"></td>
          <td width="26" align="center"><img src="Image/delinserttablec.gif" width="23" height="22"  class="Btn" title="删除列" onClick="DeleteColumn()"></td>
		  <td style="display:none;" width="26" align="center"><img src="Image/insterttable2.gif" width="23" height="22"  class="Btn" title="插入单元格" onClick="InsertCell()"></td>
          <td style="display:none;" width="26" align="center"><img src="Image/delinsterttable2.gif" width="23" height="22"  class="Btn" title="删除单元格" onClick="DeleteCell()"></td>
		  <td width="26" align="center"><img src="image/MargeTD.gif" width="23" height="22"  class="Btn" title="合并列" onClick="MergeColumn()"></td>
		  <td width="26" align="center"><img src="Image/Hbtable.gif" width="23" height="22"  class="Btn" title="合并行" onClick="MergeRow()"></td>
		  <td width="23" align="center"><img src="Image/cftable.gif" width="23" height="22"  class="Btn" title="拆分行" onClick="SplitRows()"></td>
		  <td width="23" align="center"><img src="Image/SplitTD.gif" width="23" height="22"  class="Btn" title="拆分列" onClick="SplitColumn()"></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
		  <td width="30" align="center"><img src="image/sline.gif" width="23" height="22" class="Btn" title="插入特殊水平线" onClick="SpecialHR()" ></td>
          <td width="30" align="center"><img src="image/line.gif" width="23" height="22" class="Btn" title="插入普通水平线" onClick="InsertHR();" ></td>
          <td width="30" align="center"><img title="插入换行符号" onClick="InsertBR()" class="Btn" src="image/chars.gif" ></td>
		</tr>
      </table></td>
  </tr>
  <tr> 
    <td class="ToolSet">
<table height="30" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td align="center"> 			<select name="select2" class="ToolSelectStyle" onchange="Format('fontname',this[this.selectedIndex].value);this.selectedIndex=0;EditArea.focus();">
              <option selected>字体</option>
              <option value="宋体">宋体</option>
              <option value="黑体">黑体</option>
              <option value="楷体_GB2312">楷体</option>
              <option value="仿宋_GB2312">仿宋</option>
              <option value="隶书">隶书</option>
              <option value="幼圆">幼圆</option>
              <option value="Arial">Arial</option>
              <option value="Arial Black">Arial Black</option>
              <option value="Arial Narrow">Arial Narrow</option>
              <option value="Brush Script	MT">Brush Script MT</option>
              <option value="Century Gothic">Century Gothic</option>
              <option value="Comic Sans MS">Comic Sans MS</option>
              <option value="Courier">Courier</option>
              <option value="Courier New">Courier New</option>
              <option value="MS Sans Serif">MS Sans Serif</option>
              <option value="Script">Script</option>
              <option value="System">System</option>
              <option value="Times New Roman">Times New Roman</option>
              <option value="Verdana">Verdana</option>
              <option value="Wide Latin">Wide Latin</option>
              <option value="Wingdings">Wingdings</option>
            </SELECT></td>
          <td align="center">			<select name="select3" class="ToolSelectStyle" onchange="Format('fontsize',this[this.selectedIndex].value);this.selectedIndex=0;EditArea.focus();">
              <option selected>字号</option>
              <option value="7">一号</option>
              <option value="6">二号</option>
              <option value="5">三号</option>
              <option value="4">四号</option>
              <option value="3">五号</option>
              <option value="2">六号</option>
              <option value="1">七号</option>
            </SELECT></td>
          <td width="30" align="center"><img title="加粗" onClick="Format('bold')" class="Btn" src="image/bold.gif" ></td>
          <td width="30" align="center"><img title="斜体" onClick="Format('italic')" class="Btn" src="image/italic.gif" ></td>
          <td width="30" align="center"><img title="下划线" onClick="Format('underline')" class="Btn" src="image/underline.gif" ></td>
          <td width="30" align="center"><img title="文字背景色" onClick="TextBGColor()" class="Btn" src="image/fgbgcolor.gif" ></td>
          <td width="1"> <div style="z-index:1;left:478px;top:38px;" align="center" class="ToolSeparator"></div></td>
          <td width="30" align="center">			<select name="select" class="ToolSelectStyle" onchange="Format('FormatBlock',this[this.selectedIndex].value);this.selectedIndex=0;EditArea.focus();">
              <option selected>段落样式</option>
              <option value="&lt;P&gt;">普通</option>
              <option value="&lt;H1&gt;">标题一</option>
              <option value="&lt;H2&gt;">标题二</option>
              <option value="&lt;H3&gt;">标题三</option>
              <option value="&lt;H4&gt;">标题四</option>
              <option value="&lt;H5&gt;">标题五</option>
              <option value="&lt;H6&gt;">标题六</option>
              <option value="&lt;p&gt;">段落</option>
              <option value="&lt;dd&gt;">定义</option>
              <option value="&lt;dt&gt;">术语定义</option>
              <option value="&lt;dir&gt;">目录列表</option>
              <option value="&lt;menu&gt;">菜单列表</option>
              <option value="&lt;PRE&gt;">已编排格式</option>
            </SELECT></td>
          <td width="30" align="center"><img title="减少缩进量"onClick="Format('outdent');" class="Btn" src="image/outdent.gif" ></td>
          <td width="30" align="center"><img title="增加缩进量" onClick="Format('indent')" class="Btn" src="image/indent.gif" ></td>
          <td width="30" align="center"><img src="image/abspos.gif" width="23" height="22" class="Btn" title="绝对或相对位置" onClick="Pos();" ></td>
          <td width="30" align="center"><img title="编号" onClick="Format('insertorderedlist')" class="Btn" src="image/num.gif" ></td>
          <td width="30" align="center"><img title="项目符号" onClick="Format('insertunorderedlist')" class="Btn" src="image/list.gif" ></td>
          <td width="30" align="center"><img title="左对齐" onClick="Format('justifyleft')" class="Btn" src="image/Aleft.gif" ></td>
          <td width="30" align="center"><img title="居中" onClick="Format('justifycenter')" class="Btn" src="image/Acenter.gif" ></td>
          <td width="30" align="center"><img title="右对齐" onClick="Format('justifyright')" class="Btn" src="image/Aright.gif" ></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
		  <td width="30" align="center"><img title="关于" onClick="AbortInfo()" class="Btn" src="image/Abort.gif" ></td>
		</tr>
      </table></td>
  </tr>
  <tr> 
    <td height="30"><table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" class="ToolSet">
        <tr>
          <td id="ShowObject">&nbsp;</td>
          <td width="30"><div align="center"></div></td>
          <td width="30"><div align="center"></div></td>
		</tr>
      </table></td>
  </tr>
</table>
<iframe name="EditArea" class="Composition" ID="EditArea" MARGINHEIGHT="1" MARGINWIDTH="1" width="100%" scrolling="yes"></iframe>
<textarea onFocus="SetCode();" onBlur="SetHtml();" style="width:100%;Height:100;display:none;" id="CodeEditArea" name="CodeEditArea"></textarea>
</body>
</html>
<script language="JavaScript">
var NewsContentArray=new Array('');
setTimeout('SetNewsContentArray();SetBodyStyle();',800);
function SetNewsContentArray()
{
	var AlreadyExistsNewsContent=unescape(parent.document.NewsForm.Content.value);
	if (AlreadyExistsNewsContent!='')
	{
		var TempArray;
		TempArray=AlreadyExistsNewsContent.split('[Page]');
		for (var i=0;i<TempArray.length;i++)
		{
			NewsContentArray[i+1]=TempArray[i];
		}
		SetNewsContent();
	}
	else
	{
		EditArea.document.body.innerHTML='';
	}
}
function SetNewsContent()
{
	var PageSelectObj=document.all.PageNumSelect;
	if (NewsContentArray.length>=2)
	{
		PageSelectObj.options.remove(0);
		for (var i=1;i<NewsContentArray.length;i++)
		{
			var AddOption = document.createElement("OPTION");
			AddOption.text=i;
			AddOption.value=i;
			PageSelectObj.add(AddOption);
		}
		PageSelectObj.options(0).selected=true;
		EditArea.document.body.innerHTML=NewsContentArray[1];
	}
	ShowTableBorders();
}
var BodyHeight=document.body.clientHeight;
var EditAreaHeight=BodyHeight-document.all.Toolbar.height-10;
var CodeWindowHeight=100;
var HtmlWindowHeight=EditAreaHeight;
function SetEditAreaHeight()
{
	document.all.EditArea.height=HtmlWindowHeight;
	document.all.CodeEditArea.height=CodeWindowHeight;
}
SetEditAreaHeight();
window.onresize=SetEditAreaHeight;
function NewPage()
{
	var PageSelectObj=document.all.PageNumSelect;
	var PageNum=PageSelectObj.options.length;
	NewsContentArray[parseInt(PageSelectObj.options(PageSelectObj.selectedIndex).value)]=EditArea.document.body.innerHTML;
	EditArea.document.body.innerHTML='';
	var CurrPage=PageNum+1;
	var AddOption = document.createElement("OPTION");
	AddOption.text=CurrPage;
	AddOption.value=CurrPage;
	document.all.PageNumSelect.add(AddOption);
	document.all.PageNumSelect.options(document.all.PageNumSelect.length-1).selected=true;
	EditArea.focus();
	document.all.CodeEditArea.value=EditArea.document.body.innerHTML;
	ShowTableBorders();
}
function ChangePage(PageIndex)
{
	var CurrPage=parseInt(PageIndex);
	EditArea.document.body.innerHTML=NewsContentArray[CurrPage];
	EditArea.focus();
	document.all.CodeEditArea.value=EditArea.document.body.innerHTML;
	ShowTableBorders();
}
function SaveCurrPage()
{
	var SelectObj=document.all.PageNumSelect;
	var PageIndex=parseInt(SelectObj.options(SelectObj.selectedIndex).value);
	NewsContentArray[PageIndex]=EditArea.document.body.innerHTML;
	ShowTableBorders();
}
function DeletePage()
{
	var PageNum=document.all.PageNumSelect.options.length,i=0;
	var CurrPage=parseInt(document.all.PageNumSelect.value);
	if (PageNum==1) return;
	if (CurrPage!=PageNum)
	{
		EditArea.document.body.innerHTML=NewsContentArray[CurrPage+1];
		for (i=CurrPage+1;i<=PageNum;i++)
		{
			NewsContentArray[i-1]=NewsContentArray[i];
		}
		NewsContentArray[PageNum]='';
		for (i=document.all.PageNumSelect.selectedIndex+1;i<document.all.PageNumSelect.options.length;i++)
		{
			document.all.PageNumSelect.options(i).value=parseInt(document.all.PageNumSelect.options(i).value)-1;
			document.all.PageNumSelect.options(i).text=parseInt(document.all.PageNumSelect.options(i).text)-1;
		}
		document.all.PageNumSelect.options(CurrPage).selected=true;
		document.all.PageNumSelect.options.remove(CurrPage-1);
	}
	else
	{
		EditArea.document.body.innerHTML=NewsContentArray[CurrPage-1];
		NewsContentArray[CurrPage]='';
		document.all.PageNumSelect.options(CurrPage-2).selected=true;
		document.all.PageNumSelect.options.remove(CurrPage-1);
	}
	EditArea.focus();
	document.all.CodeEditArea.value=EditArea.document.body.innerHTML;
	ShowTableBorders();
}
function SetCode()
{
	document.all.CodeEditArea.value=EditArea.document.body.innerHTML;
	ShowTableBorders();
}
function SetHtml()
{
	EditArea.document.body.innerHTML=document.all.CodeEditArea.value;
	ShowTableBorders();
}
function DisplayCodeWindow()
{
	if (document.all.CodeEditArea.style.display=='')
	{
		document.all.CodeEditArea.style.display='none';
		document.all.EditArea.style.display='';
		document.all.EditArea.height=EditAreaHeight;
	}
	else
	{
		document.all.EditArea.height=HtmlWindowHeight-CodeWindowHeight;
		document.all.CodeEditArea.height=CodeWindowHeight;
		document.all.CodeEditArea.style.display='';
	}
	ShowTableBorders();
}
function DisplayHtmlWindow()
{
	if (document.all.EditArea.style.display=='')
	{
		document.all.EditArea.style.display='none';
		document.all.CodeEditArea.style.display='';
		document.all.CodeEditArea.height=EditAreaHeight;
	}
	else
	{
		document.all.CodeEditArea.height=CodeWindowHeight;
		document.all.EditArea.height=EditAreaHeight;
		document.all.EditArea.style.display='';
	}
}
function SetBodyStyle()
{
	EditArea.document.body.runtimeStyle.fontSize='9pt';
}
function SearchObject()
{
	EditArea.focus();
	UpdateToolbar();
}
</script>