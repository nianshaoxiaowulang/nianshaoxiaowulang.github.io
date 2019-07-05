<% Option Explicit %>
<!--#include file="../../Inc/Cls_DB.asp" -->
<!--#include file="../../Inc/Const.asp" -->
<%
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../Inc/Session.asp" -->
<%
Dim LimitUpFileFlag,SysDoMain
SysDoMain = confimsn("DoMain")
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
<script language="JavaScript" src="../SysJS/PublicJS.js"></script>
<script language="javascript" event="onerror(msg, url, line)" for="window">return true;</script>
<body onLoad="return LoadEditFile();">
<table height="120" id="Toolbar" width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td class="ToolSet">
<table height="30" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="30" align="center"><img src="../Images/Editer/undo.gif" class="Btn" title="撤消" onClick="Format('undo')" ></td>
          <td width="30" align="center"><img src="../Images/Editer/redo.gif" class="Btn" title="恢复" onClick="Format('redo')" ></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
          <td width="30" align="center"><img src="../Images/Editer/NewDoc.gif" width="23" height="22" class="Btn" title="新建文档" onClick="NewPage()" ></td>
          <td align="center"> <select class="ToolSelectStyle" onFocus="SaveCurrPage();" onChange="ChangePage(this.value);" name="PageNumSelect">
              <option value="1" selected>1</option>
            </select> </td>
          <td width="30" align="center"><img onClick="DeletePage();" alt="删除文档" class="Btn" src="../Images/Editer/DelDoc.gif" width="23" height="22"></td>
          <td width="1" align="center"> <div align="center" class="ToolSeparator"></div></td>
          <td width="30" align="center"><img src="../Images/Editer/find.gif" class="Btn" title="查找 / 替换" onClick="SearchStr();" ></td>
          <td width="30" align="center"><img src="../Images/Editer/calculator.gif" class="Btn" title="计算器" onClick="Calculator()" ></td>
          <td width="30" align="center"><img title="插入当前日期" onClick="InsertDate()" class="Btn" src="../Images/Editer/date.gif" ></td>
          <td width="30" align="center"><img title="插入当前时间" onClick="InsertTime()" class="Btn" src="../Images/Editer/time.gif" ></td>
          <td width="30" align="center"><img title="删除所有HTML标识" onClick="DelAllHtmlTag()" class="Btn" src="../Images/Editer/geshi.gif" ></td>
          <td width="30" align="center"><img title="删除文字格式" onClick="Format('removeformat')" class="Btn" src="../Images/Editer/clear.gif" ></td>
          <td width="1" align="center"> <div align="center" class="ToolSeparator"></div></td>
		  <td width="30" align="center"><img title="插入超级连接" onClick="Format('CreateLink')" class="Btn" src="../Images/Editer/url.gif" ></td>
          <td width="30" align="center"><img title="取消超级链接" onClick="Format('unLink')" class="Btn" src="../Images/Editer/nourl.gif" ></td>
          <td width="30" align="center"><img title="插入网页" onClick="InsertPage()" class="Btn" src="../Images/Editer/htm.gif" ></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
          <td width="30" align="center"><img title="插入栏目框" onClick="InsertFrame()" class="Btn" src="../Images/Editer/fieldset.gif" ></td>
          <td width="30" align="center"><img title="插入Excel表格" onClick="InsertExcel()" class="Btn" src="../Images/Editer/excel.gif" ></td>
          <td width="30" align="center"><img title="插入滚动文本" onClick="InsertMarquee()" class="Btn" src="../Images/Editer/Marquee.gif" ></td>
          <td width="30" align="center"><img title="插入图片，支持格式为：jpg、gif、bmp、png等" onClick="InsertPicture('<% = LimitUpFileFlag %>')" class="Btn" src="../Images/Editer/img.gif" ></td>
          <td width="30" align="center"><img title="插入flash多媒体文件" onClick="InsertFlash('<% = LimitUpFileFlag %>')" class="Btn" src="../Images/Editer/flash.gif" ></td>
          <td width="30" align="center"><img title="插入视频文件，支持格式为：avi、wmv、asf、mpg" onClick="InsertVideo('<% = LimitUpFileFlag %>')" class="Btn" src="../Images/Editer/wmv.gif" ></td>
          <td width="30" align="center"><img title="插入RealPlay文件，支持格式为：rm、ra、ram" onClick="InsertRM('<% = LimitUpFileFlag %>')" class="Btn" src="../Images/Editer/rm.gif" ></td>
          <td width="30" align="center"><img src="../Images/Editer/PicAlign.gif" width="23" height="22" class="Btn" title="图文并排" onClick="PicAndTextArrange()" ></td>
          <td width="30" align="center"><img src="../Images/Editer/TextPaste.jpg" class="Btn" title="文本粘贴" onClick="PureTextPaste()" ></td> 
		</tr>
      </table></td>
  </tr>
  <tr> 
    <td height="30" class="ToolSet"> 
      <table height="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="26" align="center"><img src="../Images/Editer/Inserttable.gif"  class="Btn" title="插入表格" onClick="InsertTable()"></td>
          <td width="26" align="center"><img src="../Images/Editer/inserttable1.gif" width="23" height="22"  class="Btn" title="插入行" onClick="InsertRow()"></td>
          <td width="26" align="center"><img src="../Images/Editer/delinserttable1.gif" width="23" height="22"  class="Btn" title="删除行" onClick="DeleteRow()"></td>
		  <td width="26" align="center"><img src="../Images/Editer/inserttablec.gif" width="23" height="22"  class="Btn" title="插入列" onClick="InsertColumn()"></td>
          <td width="26" align="center"><img src="../Images/Editer/delinserttablec.gif" width="23" height="22"  class="Btn" title="删除列" onClick="DeleteColumn()"></td>
		  <td style="display:none;" width="26" align="center"><img src="../Images/Editer/insterttable2.gif" width="23" height="22"  class="Btn" title="插入单元格" onClick="InsertCell()"></td>
          <td style="display:none;" width="26" align="center"><img src="../Images/Editer/delinsterttable2.gif" width="23" height="22"  class="Btn" title="删除单元格" onClick="DeleteCell()"></td>
		  <td width="26" align="center"><img src="../Images/Editer/MargeTD.gif" width="23" height="22"  class="Btn" title="合并列" onClick="MergeColumn()"></td>
		  <td width="26" align="center"><img src="../Images/Editer/Hbtable.gif" width="23" height="22"  class="Btn" title="合并行" onClick="MergeRow()"></td>
		  <td width="23" align="center"><img src="../Images/Editer/cftable.gif" width="23" height="22"  class="Btn" title="拆分行" onClick="SplitRows()"></td>
		  <td width="23" align="center"><img src="../Images/Editer/SplitTD.gif" width="23" height="22"  class="Btn" title="拆分列" onClick="SplitColumn()"></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
		  <td width="30" align="center"><img src="../Images/Editer/sline.gif" width="23" height="22" class="Btn" title="插入特殊水平线" onClick="SpecialHR()" ></td>
          <td width="30" align="center"><img src="../Images/Editer/line.gif" width="23" height="22" class="Btn" title="插入普通水平线" onClick="InsertHR();" ></td>
          <td width="30" align="center"><img title="插入换行符号" onClick="InsertBR()" class="Btn" src="../Images/Editer/chars.gif" ></td>
		  <td width="30" align="center"><img title="插入分页符号" onClick="InsertMorePage()" class="Btn" src="../Images/Editer/InsertPage.gif" ></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
          <td width="30" align="center"><img src="../Images/Editer/down.gif" width="23" height="22" class="Btn" title="插入下载" onClick="InsertDownLoad('<% = SysDoMain %>')" ></td>
		  <td width="30" align="center"><img src="../Images/Editer/ToAddress.gif" width="23" height="22" class="Btn" title="选择图片作为图片新闻的图片" onClick="SetPicNews()" ></td>
		</tr>
      </table></td>
  </tr>
  <tr> 
    <td class="ToolSet">
<table height="30" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td align="center">			<select name="select2" class="ToolSelectStyle" onchange="Format('fontname',this[this.selectedIndex].value);this.selectedIndex=0;EditArea.focus();">
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
          <td width="30" align="center"><img title="加粗" onClick="Format('bold')" class="Btn" src="../Images/Editer/bold.gif" ></td>
          <td width="30" align="center"><img title="斜体" onClick="Format('italic')" class="Btn" src="../Images/Editer/italic.gif" ></td>
          <td width="30" align="center"><img title="下划线" onClick="Format('underline')" class="Btn" src="../Images/Editer/underline.gif" ></td>
		  <td width="30" align="center"><img src="../Images/Editer/TextColor.gif" width="23" height="22" class="Btn" title="文字颜色" onClick="TextColor()" ></td>
		  <td width="30" align="center"><img title="文字背景色" onClick="TextBGColor()" class="Btn" src="../Images/Editer/fgbgcolor.gif" ></td>
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
          <td width="30" align="center"><img title="减少缩进量"onClick="Format('outdent');" class="Btn" src="../Images/Editer/outdent.gif" ></td>
          <td width="30" align="center"><img title="增加缩进量" onClick="Format('indent')" class="Btn" src="../Images/Editer/indent.gif" ></td>
          <td width="30" align="center"><img src="../Images/Editer/abspos.gif" width="23" height="22" class="Btn" title="绝对或相对位置" onClick="Pos();" ></td>
          <td width="30" align="center"><img title="编号" onClick="Format('insertorderedlist')" class="Btn" src="../Images/Editer/num.gif" ></td>
          <td width="30" align="center"><img title="项目符号" onClick="Format('insertunorderedlist')" class="Btn" src="../Images/Editer/list.gif" ></td>
          <td width="30" align="center"><img title="左对齐" onClick="Format('justifyleft')" class="Btn" src="../Images/Editer/Aleft.gif" ></td>
          <td width="30" align="center"><img title="居中" onClick="Format('justifycenter')" class="Btn" src="../Images/Editer/Acenter.gif" ></td>
          <td width="30" align="center"><img title="右对齐" onClick="Format('justifyright')" class="Btn" src="../Images/Editer/Aright.gif" ></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
		  <td width="30" align="center"><img title="关于" onClick="AbortInfo()" class="Btn" src="../Images/Editer/Abort.gif" ></td>
		</tr>
      </table></td>
  </tr>
  <tr> 
    <td height="30"><table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" class="ToolSet">
        <tr> 
          <td id="ShowObject">&nbsp;</td>
          <td width="30"><div align="center"><img src="../Images/Editer/tablemodify.gif" width="23" height="22"  class="Btn" title="属性" onClick="ExeEditAttribute();"></div></td>
          <td width="30"><div align="center"><img src="../Images/Editer/delLable.gif" width="23" height="22"  class="Btn" title="删除标签" onClick="DeleteHTMLTag();"></div></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td><iframe name="EditArea" class="Composition" ID="EditArea" MARGINHEIGHT="1" MARGINWIDTH="1" height="100%;" width="100%" scrolling="yes"></iframe></td>
  </tr>
  <tr> 
    <td height="20" id="SetModeArea"> <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="60" height="20" align="center" class="ModeBarBtnOff" id=Editer_CODE onClick="setMode('CODE');"><img src="../Images/Editer/CodeMode.GIF" width="50" height="15"></td>
          <td width="60" height="20" align="center" class="ModeBarBtnOff" id=Editer_VIEW onClick="setMode('VIEW');"><img src="../Images/Editer/PreviewMode.gif" width="50" height="15"></td>
          <td width="60" height="20" align="center" class="ModeBarBtnOn" id=Editer_EDIT onClick="setMode('EDIT');"><img src="../Images/Editer/EditMode.GIF" width="50" height="15"></td>
          <td width="60" height="20" align="center" class="ModeBarBtnOff" id=Editer_TEXT onClick="setMode('TEXT');"><img src="../Images/Editer/TextMode.GIF" width="50" height="15"></td>
          <td height="20">&nbsp;</td>
          <td style="display:none;" width="30" height="20" align="center" onClick="AddHeight();"><img class="Btn" src="../Images/Editer/AddHeight.gif" width="23" height="22"></td>
          <td style="display:none;" width="30" height="20" align="center" onClick="MinusHeight();"><img class="Btn" src="../Images/Editer/MinusHeight.gif" width="23" height="22"></td>
        </tr>
      </table></td>
  </tr>
</table>
</body>
</html>
<script language="JavaScript">
var NewsContentArray=new Array('');
var DocumentReadyTF=false;
function SetPicNews()
{
	if (parent.document.NewsForm.IsPicNews.checked==true)
	{
		var SelectObj=GetSelection();
		var ObjContent=CreateRange(SelectObj).item(0);
		if(SelectObj.type=="Control")
			if (ObjContent.tagName.toUpperCase()=='IMG')
				parent.document.NewsForm.PicPath.value=ObjContent.src;
			else
			{
				alert("你选择的不是图片，请选择图片！");
				return;
			}
		else
		{
			alert("你选择的不是图片，请选择图片！");
			return;
		}
	}
	else
	alert('请先将此新闻转为图片新闻,才可指定图片');
}
function document.onreadystatechange()
{
	if (document.readyState!="complete") return;
	if (DocumentReadyTF) return;
	DocumentReadyTF = true;
	var i,j,s,curr;
	for (i=0; i<document.body.all.length;i++)
	{
		curr=document.body.all[i];
		if (curr.className=="Btn") InitBtn(curr);
	}
	ShowTableBorders();
	SetNewsContentArray();
}
function SetNewsContentArray()
{
	parent.document.NewsForm.Content.value=unescape(parent.document.NewsForm.Content.value)
	var AlreadyExistsNewsContent=parent.document.NewsForm.Content.value;
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
var EditAreaHeight=BodyHeight-document.all.Toolbar.height-document.all.SetModeArea.height-8;
var CodeWindowHeight=100;
var HtmlWindowHeight=EditAreaHeight;
function SetEditAreaHeight()
{
	document.all.EditArea.height=HtmlWindowHeight;
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
	ShowTableBorders();
}
function ChangePage(PageIndex)
{
	var CurrPage=parseInt(PageIndex);
	EditArea.document.body.innerHTML=NewsContentArray[CurrPage];
	EditArea.focus();
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
	ShowTableBorders();
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