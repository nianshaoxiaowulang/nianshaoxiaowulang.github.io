bInitialized=false;
var ObjPopupMenu = null;
ObjPopupMenu = window.createPopup();
var SelectedTD=null;
var SelectedTR=null;
var SelectedTBODY=null;
var SelectedTable=null;
function InitBtn(btn) 
{
	btn.onmouseover = ImageBtnMouseOver;
	btn.onmouseout = BtnMouseOut;
	btn.onmousedown = BtnMouseDown;
	btn.onmouseup = BtnMouseOut;
	btn.ondragstart = YCancelEvent;
	btn.onselectstart = YCancelEvent;
	btn.onselect = YCancelEvent;
	btn.disabled=false;
	return true;
}

function ImageBtnMouseOver() 
{
	var image = event.srcElement;
	image.className = "ToolBtnMouseOver";
	event.cancelBubble = true;
}

function BtnMouseOut() 
{
	var image = event.srcElement;
	image.className = "Btn";
	event.cancelBubble = true;
}

function BtnMouseDown() 
{
	var image = event.srcElement;
	image.className = "ToolBtnMouseDown";
	event.cancelBubble = true;
	event.returnValue=false;
	return false;
}

function YCancelEvent() 
{
	event.returnValue=false;
	event.cancelBubble=true;
	return false;
}
function LoadEditFile(FileName)
{
	EditArea.document.body.contentEditable="true";
	EditArea.document.onmouseup=new Function("return SearchObject(EditArea.event);");
	//EditArea.document.onkeyup=new Function("return SearchObject(EditArea.event);");
	EditArea.document.oncontextmenu=new Function("return ShowMouseRightMenu(EditArea.event);");
	EditArea.focus();
}

var CurrMode='EDIT';
function setMode(NewMode)   
{  
	if (NewMode!=CurrMode)   
	{   
		if (NewMode=='TEXT')
		{
			if (!confirm("警告！切换到纯文本模式会丢失您所有的HTML格式，您确认切换吗？")) return false;
		}
		var sBody='';
		switch(CurrMode)
		{
			case "CODE":
				if (NewMode=="TEXT") sBody=EditArea.document.body.innerText;
				else sBody=EditArea.document.body.innerText;
				break;
			case "TEXT":
				sBody=EditArea.document.body.innerText;
				sBody=HTMLEncode(sBody);
				break;
			case "EDIT":
			case "VIEW":
				if (NewMode=="TEXT") sBody=EditArea.document.body.innerText;
				else sBody=EditArea.document.body.innerHTML;
				break;
		}
		document.all["Editer_CODE"].className='ModeBarBtnOff';
		document.all["Editer_EDIT"].className='ModeBarBtnOff';
		document.all["Editer_TEXT"].className='ModeBarBtnOff';
		document.all["Editer_VIEW"].className='ModeBarBtnOff';
		document.all["Editer_"+NewMode].className='ModeBarBtnOn';
		switch (NewMode)
		{
			case "CODE":
				EditArea.document.designMode="On";
				EditArea.document.open();
				EditArea.document.write("<head><link href=\"../../CSS/FS_css.css\" type=\"text/css\" rel=\"stylesheet\"></head><body MONOSPACE>");
				EditArea.document.body.innerText=sBody;
				EditArea.document.body.contentEditable="true";
				EditArea.document.close();
				DisabledAllBtn(true);
				break;
			case "EDIT":
				EditArea.document.designMode="On";
				EditArea.document.open();
				EditArea.document.write("<head><link href=\"../../CSS/FS_css.css\" type=\"text/css\" rel=\"stylesheet\"></head><body MONOSPACE>"+sBody);
				EditArea.document.body.contentEditable="true";
				EditArea.document.execCommand("2D-Position",true,true);
				EditArea.document.execCommand("MultipleSelection", true, true);
				EditArea.document.execCommand("LiveResize", true, true);
				EditArea.document.close();
				ShowTableBorders();
				DisabledAllBtn(false);
				break;
			case "TEXT":
				EditArea.document.designMode="On";
				EditArea.document.open();
				EditArea.document.write("<head><link href=\"../../CSS/FS_css.css\" type=\"text/css\" rel=\"stylesheet\"></head><body MONOSPACE>");
				EditArea.document.body.innerText=sBody;
				EditArea.document.body.contentEditable="true";
				EditArea.document.close();
				DisabledAllBtn(true);
				break;
			case "VIEW":
				EditArea.document.designMode="off";
				EditArea.document.open();
				EditArea.document.write("<head><link href=\"../../CSS/FS_css.css\" type=\"text/css\" rel=\"stylesheet\"></head><body MONOSPACE>"+sBody);
				EditArea.document.body.contentEditable="false";
				EditArea.document.close();
				DisabledAllBtn(true);
				break;
		}
		CurrMode=NewMode;
		if (NewMode!='EDIT') EmptyShowObject(true);
		else {EmptyShowObject(false);LoadEditFile();}
	}
	EditArea.focus();
}
function setTempletMode(NewMode)   
{  
	if (NewMode!=CurrMode)   
	{   
		if (NewMode=='TEXT')
		{
			if (!confirm("警告！切换到纯文本模式会丢失您所有的HTML格式，您确认切换吗？")) return false;
		}
		var sBody='';
		switch(CurrMode)
		{
			case "CODE":
				if (NewMode=="TEXT") sBody=EditArea.document.body.innerText;
				else sBody=EditArea.document.body.innerText;
				break;
			case "TEXT":
				sBody=EditArea.document.body.innerText;
				sBody=HTMLEncode(sBody);
				break;
			case "EDIT":
			case "VIEW":
				if (NewMode=="TEXT") sBody=EditArea.document.body.innerText;
				else sBody=EditArea.document.body.innerHTML;
				break;
		}
		document.all["Editer_CODE"].className='ModeBarBtnOff';
		document.all["Editer_EDIT"].className='ModeBarBtnOff';
		document.all["Editer_TEXT"].className='ModeBarBtnOff';
		document.all["Editer_VIEW"].className='ModeBarBtnOff';
		document.all["Editer_"+NewMode].className='ModeBarBtnOn';
		var documentElementStr='<html>'+EditArea.document.documentElement.innerHTML+'</html>';
		switch (NewMode)
		{
			case "CODE":
				EditArea.document.designMode="On";
				EditArea.document.open();
				EditArea.document.write(documentElementStr);
				EditArea.document.body.innerText=sBody;
				EditArea.document.body.contentEditable="true";
				EditArea.document.close();
				DisabledAllBtn(true);
				break;
			case "EDIT":
				EditArea.document.designMode="On";
				EditArea.document.open();
				EditArea.document.write(documentElementStr);
				EditArea.document.body.innerHTML=sBody;
				EditArea.document.body.contentEditable="true";
				EditArea.document.close();
				ShowTableBorders();
				DisabledAllBtn(false);
				break;
			case "TEXT":
				EditArea.document.designMode="On";
				EditArea.document.open();
				EditArea.document.write(documentElementStr);
				EditArea.document.body.innerText=sBody;
				EditArea.document.body.contentEditable="true";
				EditArea.document.close();
				DisabledAllBtn(true);
				break;
			case "VIEW":
				EditArea.document.designMode="off";
				EditArea.document.open();
				EditArea.document.write(documentElementStr);
				EditArea.document.body.innerHTML=sBody;
				EditArea.document.body.contentEditable="false";
				EditArea.document.close();
				DisabledAllBtn(true);
				break;
		}
		CurrMode=NewMode;
		if (NewMode!='EDIT') EmptyShowObject(true);
		else {EmptyShowObject(false);LoadEditFile();}
	}
	EditArea.focus();
}
function EmptyShowObject(Flag)
{
	document.all.ShowObject.disabled=Flag;
}
function HTMLEncode(text)
{
	text = text.replace(/&/g, "&amp;") ;
	text = text.replace(/"/g, "&quot;") ;
	text = text.replace(/</g, "&lt;") ;
	text = text.replace(/>/g, "&gt;") ;
	text = text.replace(/'/g, "&#146;") ;
	text = text.replace(/\ /g,"&nbsp;");
	text = text.replace(/\n/g,"<br>");
	text = text.replace(/\t/g,"&nbsp;&nbsp;&nbsp;&nbsp;");
	return text;
}
function ShowMouseRightMenu(event)
{
	var width=86;
	var height=0;
	var lefter=event.clientX;
	var topper=event.clientY;
	var ObjPopDocument=ObjPopupMenu.document;
	var ObjPopBody=ObjPopupMenu.document.body;
	var MenuStr='';
	MenuStr+=FormatMenuRow("selectall", "全选","SelectAll.gif");
	MenuStr+=FormatMenuRow("cut", "剪切","Cut.gif");
	MenuStr+=FormatMenuRow("copy", "复制","Copy.gif");
	MenuStr+=FormatMenuRow("paste", "粘贴","Paste.gif");
	MenuStr+=FormatMenuRow("delete", "删除","Del.gif");
	height+=100;
	MenuStr="<TABLE border=0 cellpadding=0 cellspacing=0 class=Menu width=86><tr><td width=86 class=RightBg><TABLE border=0 cellpadding=0 cellspacing=0>"+MenuStr
	MenuStr=MenuStr+"<\/TABLE><\/td><\/tr><\/TABLE>";
	ObjPopDocument.open();
	ObjPopDocument.write("<head><link href=\"MenuCSS.css\" type=\"text/css\" rel=\"stylesheet\"></head><body scroll=\"no\" onConTextMenu=\"event.returnValue=false;\">"+MenuStr);
	ObjPopDocument.close();
	height+=5;
	if(lefter+width > document.body.clientWidth) lefter=lefter-width;
	ObjPopupMenu.show(lefter, topper, width, height, EditArea.document.body);
	return false;
}
function GetMenuRowStr(DisabledStr, MenuOperation, MenuImage, MenuDescripion)
{
	var MenuRowStr='';
	MenuRowStr="<tr><td align=center valign=middle><TABLE border=0 cellpadding=0 cellspacing=0 width=81><tr "+DisabledStr+"><td valign=middle height=20 class=MouseOut onMouseOver=this.className='MouseOver'; onMouseOut=this.className='MouseOut';";
	if (DisabledStr==''){
		MenuRowStr += " onclick=\"parent."+MenuOperation+";parent.ObjPopupMenu.hide();\"";
	}
	MenuRowStr+=">"
	if (MenuImage!="")
	{
		MenuRowStr+="&nbsp;<img border=0 src='../Images/Editer/"+MenuImage+"' width=20 height=20 align=absmiddle "+DisabledStr+">&nbsp;";
	}
	else
	{
		MenuRowStr+="&nbsp;";
	}
	MenuRowStr+=MenuDescripion+"<\/td><\/tr><\/TABLE><\/td><\/tr>";
	return MenuRowStr;

}
function FormatMenuRow(MenuStr,MenuDescription,MenuImage)
{
	var DisabledStr='';
	var ShowMenuImage='';
	if (!EditArea.document.queryCommandEnabled(MenuStr))
	{
		DisabledStr="disabled";
	}
	var MenuOperation="Format('"+MenuStr+"')";
	if (MenuImage)
	{
		ShowMenuImage=MenuImage;
	}
	return GetMenuRowStr(DisabledStr,MenuOperation,ShowMenuImage,MenuDescription)
}
function SearchObject()
{
	UpdateToolbar();
}
function MouseRightMenuItem(CommandString, CommandId)
{
	this.CommandString = CommandString;
	this.CommandId = CommandId;
}
function GetEditAreaSelectionType()
{
	return EditArea.document.selection.type;
}
var ContextMenu = new Array();
function ExeEditAttribute()
{
	OpenWindow('AttributeWindow.htm',360,120,window)
	EditArea.focus();
}
function InsertHTMLStr(Str)
{
	EditArea.focus();
	if (EditArea.document.selection.type.toLowerCase() != "none")
	{
		EditArea.document.selection.clear() ;
	}
	EditArea.document.selection.createRange().pasteHTML(Str) ; 
	EditArea.focus();
	ShowTableBorders();
}
//Editer Btn Click Event Function Begin.
function InsertPicture(LimitUpFileFlag)
{
	//alert(LimitUpFileFlag);
	//return;
	var ReturnValue=OpenWindow('Picture.asp?LimitUpFileFlag='+LimitUpFileFlag,420,180,window);
	if (ReturnValue!='')
	{
		var TempArray=ReturnValue.split("$$$");
		InsertHTMLStr(TempArray[0]);
	}
}
function QueryCommand(CommandID)
{
	var State=EditArea.QueryStatus(CommandID)
	if (State==3) return true;
	else return false;
}
function Format(Operation,Val) 
{
	EditArea.focus();
	if (Val=="RemoveFormat")
	 {
		Operation=Val;
		Val=null;
	}
	if (Val==null) EditArea.document.execCommand(Operation);
	else EditArea.document.execCommand(Operation,"",Val);
	EditArea.focus();
}
function TextBGColor()
{
	EditArea.focus();
	var EditRange = EditArea.document.selection.createRange();
	var RangeType = EditArea.document.selection.type;
	if (RangeType!="Text")
	{
		alert("请先选择一段文字！");
		return;
	}
	var ReturnValue=OpenWindow('SelectColor.htm',230,190,window);
	if (ReturnValue!=null)
	{
		EditRange.pasteHTML("<span style='background-color:"+ReturnValue+"'>"+EditRange.text+"</span> ");
		EditRange.select();
	}
	EditArea.focus();
}
function Print(CommandID)
{
	EditArea.focus();
	//alert(EditArea.QueryStatus(CommandID));
	if (EditArea.QueryStatus(CommandID)!=3) EditArea.ExecCommand(CommandID,0);
	EditArea.focus();
}
function InsertTable()
{
	var ReturnValue=OpenWindow('InsertTable.htm',290,110,window);
	InsertHTMLStr(ReturnValue);
	EditArea.focus();
}
function InsertPage()
{
	var ReturnValue=OpenWindow('Page.htm',320,110,window);
	InsertHTMLStr(ReturnValue);
	EditArea.focus();
}
function InsertExcel()
{
	EditArea.focus();
	var TempStr="<object classid='clsid:0002E510-0000-0000-C000-000000000046' id='Spreadsheet1' codebase='file:\\Bob\software\office2000\msowc.cab' width='100%' height='250'><param name='EnableAutoCalculate' value='-1'><param name='DisplayTitleBar' value='0'><param name='DisplayToolbar' value='-1'><param name='ViewableRange' value='1:65536'></object>";
	InsertHTMLStr(TempStr);
	EditArea.focus();
}
function InsertMarquee()
{
	EditArea.focus();
	var ReturnValue=OpenWindow('Marquee.htm',260,50,window); 
	InsertHTMLStr(ReturnValue);
	EditArea.focus();
}
function Calculator()
{
	EditArea.focus();
	var ReturnValue=OpenWindow('Calculator.htm',160,180,window);
	if (ReturnValue!=null)
	{
		var TempArray,ParameterA,ParameterB;
		TempArray=ReturnValue.split("*")
		ParameterA=TempArray[0];
		InsertHTMLStr(ParameterA);
	}
	EditArea.focus();
}
function InsertDate()
{
	EditArea.focus();
	var NowDate = new Date();
	var FormateDate=NowDate.getYear()+"年"+(NowDate.getMonth() + 1)+"月"+NowDate.getDate() +"日";
	InsertHTMLStr(FormateDate);
	EditArea.focus();
}
function InsertTime()
{
	EditArea.focus();
	var NowDate=new Date();
	var FormatTime=NowDate.getHours() +":"+NowDate.getMinutes()+":"+NowDate.getSeconds();
	InsertHTMLStr(FormatTime);
	EditArea.focus();
}
function InsertFrame()
{
	EditArea.focus();
	var ReturnVlaue =OpenWindow('Frame.htm',280,118,window);
	if (ReturnVlaue != null)
	{
		InsertHTMLStr(ReturnVlaue);
	}
	EditArea.focus();
}
function InsertBR(Index)
{
	EditArea.focus();
	InsertHTMLStr('<br>');
	EditArea.focus();
}
function InsertMorePage(Index)
{
	EditArea.focus();
	InsertHTMLStr('[Page]');
	EditArea.focus();
}
function DelAllHtmlTag()
{
	var TempStr;
	TempStr=EditArea.document.body.innerHTML;
	var re=/<\/*[^<>]*>/ig
	TempStr=TempStr.replace(re,'');
	EditArea.document.body.innerHTML=TempStr;
}
function AbortInfo()
{
  var arr = OpenWindow('Abort.htm',220,100,window);
}
function InsertFlash(LimitUpFileFlag)
{
  var ReturnValue = OpenWindow('Flash.asp?LimitUpFileFlag='+LimitUpFileFlag,380,100,window); 
  if (ReturnValue!='')
  {
    var TempArray=ReturnValue.split("$$$");
    InsertHTMLStr(TempArray[0]);
  }
  EditArea.focus();
}
function InsertVideo(LimitUpFileFlag)
{
  var ReturnValue=OpenWindow('Video.asp?LimitUpFileFlag='+LimitUpFileFlag,400,100,window);
  if (ReturnValue!='')
  {
    var TempArray=ReturnValue.split("$$$");
    InsertHTMLStr(TempArray[0]);
  }
  EditArea.focus();
}
function InsertRM(LimitUpFileFlag)
{
  var ReturnValue=OpenWindow('RM.asp?LimitUpFileFlag='+LimitUpFileFlag,400,100,window);  
  if (ReturnValue!='')
  {
    var TempArray=ReturnValue.split("$$$");
    InsertHTMLStr(TempArray[0]);
  }
  EditArea.focus();
}
function SpecialHR()
{
	EditArea.focus();
	var ReturnValue = OpenWindow('SpecialHR.htm',320,120,window); 
	if (ReturnValue!= null) InsertHTMLStr(ReturnValue);
	EditArea.focus();
}
function InsertHR()
{
	EditArea.focus();
	InsertHTMLStr('<hr>');
	EditArea.focus();
}
var BorderShown=1;
function ShowTableBorders()
{
	AllTables=EditArea.document.body.getElementsByTagName("TABLE");
	for(var i=0;i<AllTables.length;i++)
	{
		if ((AllTables[i].border==null)||(AllTables[i].border=='0'))
		{
			AllTables[i].runtimeStyle.borderTop=AllTables[i].runtimeStyle.borderLeft="1px dotted #709FCB";
			AllRows = AllTables[i].rows;
			for(var y=0;y<AllRows.length;y++)
			{
				AllCells=AllRows[y].cells;
				for(var x=0;x<AllCells.length;x++)
				{
					AllCells[x].runtimeStyle.borderRight=AllCells[x].runtimeStyle.borderBottom="1px dotted #709FCB";
				}
			}
		}
		else
		{
			AllTables[i].runtimeStyle.borderTop='';
			AllRows=AllTables[i].rows;
			for(var y=0;y<AllRows.length;y++)
			{
				AllCells=AllRows[y].cells;
				for(var x=0;x<AllCells.length;x++)
				{
					AllCells[x].runtimeStyle.borderRight=AllCells[x].runtimeStyle.borderBottom='';
				}
			}
		}
	}
  BorderShown=BorderShown?0:1;
}
function ImageSelected()
{
	EditArea.focus();
	if (EditArea.document.selection.type=="Control")
	{
		var oControlRange=EditArea.document.selection.createRange();
		if (oControlRange(0).tagName.toUpperCase()=="IMG")
		{
			selectedImage=EditArea.document.selection.createRange()(0);
			return true;
		}	
	}
}
function TextColor()
{	
	EditArea.focus();
	var EditRange = EditArea.document.selection.createRange();
	var RangeType = EditArea.document.selection.type;
	if (RangeType!="Text")
	{
		alert("请先选择一段文字！");
		return;
	}
	var ReturnValue=OpenWindow('SelectColor.htm',230,190,window);
	if (ReturnValue!=null)
	{
		EditRange.pasteHTML("<font color='"+ReturnValue+"'>"+EditRange.text+"</font>");
		EditRange.select();
	}
	EditArea.focus();
}
function PicAndTextArrange()
{
	if(ImageSelected())
	{
		sPrePos=selectedImage.style.position;
		var ReturnValue=OpenWindow('SelectPicStyle.htm',380,130,window);
		if(ReturnValue)
		{
			for(key in ReturnValue)
			if(key=='style') for(sub_key in ReturnValue.style) selectedImage.style[sub_key]=ReturnValue.style[sub_key];
			else selectedImage[key]=ReturnValue[key];
			if(!ReturnValue.align) selectedImage.removeAttribute('align');
			if(sPrePos.match(/^absolute$/i) && !selectedImage.style.position.match(/^absolute$/i))
			{
				sFired = selectedImage.parentElement;
				while(!sFired.tagName.match(/^table$|^body$/i))
				sFired = sFired.parentElement;
				if(sFired.tagName.match(/^table$/i) && sFired.style.position.match(/absolute/i));
				sFired.outerHTML=selectedImage.outerHTML;
			}
			else
			{
				if(!sPrePos.match(/^absolute$/i) && selectedImage.style.position.match(/^absolute$/i)) selectedImage.outerHTML='<table style="position: absolute;"><tr><td>' + selectedImage.outerHTML + '</td></tr></table>';
			}
		}
	}
	else alert('请选择图片');
}
function GetAllAncestors()
{
	var p = GetParentElement();
	var a = [];
	while (p && (p.nodeType==1)&&(p.tagName.toLowerCase()!='body'))
	{
		a.push(p);
		p=p.parentNode;
	}
	a.push(EditArea.document.body);
	return a;
}
function GetParentElement()
{
	var sel=GetSelection();
	var range=CreateRange(sel);
	switch (sel.type)
	{
		case "Text":
		case "None":
			return range.parentElement();
		case "Control":
			return range.item(0);
		default:
			return EditArea.document.body;
	}
}
function GetSelection()
{
	return EditArea.document.selection;
}
function CreateRange(sel)
{
	return sel.createRange();
}
function UpdateToolbar()
{
	var ancestors=null;
	ancestors=GetAllAncestors();
	document.all.ShowObject.innerHTML='&nbsp;';
	for (var i=ancestors.length;--i>=0;)
	{
		var el = ancestors[i];
		if (!el) continue;
		var a=document.createElement("span");
		a.href="#";
		a.el=el;
		a.editor=this;
		if (i==0)
		{
			a.className='AncestorsMouseUp';
			EditControl=a.el;
		}
		else a.className='AncestorsStyle';
		a.onmouseover=function()
		{
			if (this.className=='AncestorsMouseUp') this.className='AncestorsMouseUpOver';
			else if (this.className=='AncestorsStyle') this.className='AncestorsMouseOver';
		};
		a.onmouseout=function()
		{
			if (this.className=='AncestorsMouseUpOver') this.className='AncestorsMouseUp';
			else if (this.className=='AncestorsMouseOver') this.className='AncestorsStyle';
		};
		a.onmousedown=function(){this.className='AncestorsMouseDown';};
		a.onmouseup=function(){this.className='AncestorsMouseUpOver';};
		a.ondragstart=YCancelEvent;
		a.onselectstart=YCancelEvent;
		a.onselect=YCancelEvent;
		a.onclick=function()
		{
			this.blur();
			SelectNodeContents(this);
			return false;
		};
		var txt='<'+el.tagName.toLowerCase();
		a.title=el.style.cssText;
		if (el.id) txt += "#" + el.id;
		if (el.className) txt += "." + el.className;
		txt=txt+'>';
		a.appendChild(document.createTextNode(txt));
		document.all.ShowObject.appendChild(a);
	}
}
function SelectNodeContents(Obj,pos)
{
	var node=Obj.el;
	EditControl=node;
	for (var i=0;i<document.all.ShowObject.children.length;i++)
	{
		if (document.all.ShowObject.children(i).className=='AncestorsMouseUp') document.all.ShowObject.children(i).className='AncestorsStyle';
	}
	//Obj.className='AncestorsMouseUp';
	EditArea.focus();
	var range;
	var collapsed=(typeof pos!='undefined');
	range = EditArea.document.body.createTextRange();
	range.moveToElementText(node);
	(collapsed) && range.collapse(pos);
	range.select();
}
function DeleteHTMLTag()
{
	var AvailableDeleteTagName='p,a,div,span';
	if (EditControl!=null)
	{
		var DeleteTagName=EditControl.tagName.toLowerCase();
		if (AvailableDeleteTagName.indexOf(DeleteTagName)!=-1)
		{
			EditControl.parentElement.innerHTML=EditControl.innerHTML;
		}
	}
	UpdateToolbar();
	ShowTableBorders();
}
function InsertRow()
{
	if (CursorInTableCell())
	{
		var SelectColsNum=0;
		var AllCells=SelectedTR.cells;
		for (var i=0;i<AllCells.length;i++)
		{
		 	SelectColsNum=SelectColsNum+AllCells[i].getAttribute('colSpan');
		}
		var NewTR=SelectedTable.insertRow(SelectedTR.rowIndex);
		for (i=0;i<SelectColsNum;i++)
		{
		 	NewTD=NewTR.insertCell();
			NewTD.innerHTML="&nbsp;";
		}
	}
	ShowTableBorders();	
}
function InsertColumn()
{
   	if (CursorInTableCell())
	{
		var MoveFromEnd=(SelectedTR.cells.length-1)-(SelectedTD.cellIndex);
		var AllRows=SelectedTable.rows;
		for (i=0;i<AllRows.length;i++)
		{
			RowCount=AllRows[i].cells.length-1;
			Position=RowCount-MoveFromEnd;
			if (Position<0)
			{
				Position=0;
			}
			NewCell=AllRows[i].insertCell(Position);
			NewCell.innerHTML="&nbsp;";
		}
		ShowTableBorders();
	}	
}
function DeleteRow()
{
	if (CursorInTableCell())
	{
		SelectedTable.deleteRow(SelectedTR.rowIndex);
	}
}
function DeleteColumn()
{
   	if (CursorInTableCell())
	{
		var MoveFromEnd=(SelectedTR.cells.length-1)-(SelectedTD.cellIndex);
		var AllRows=SelectedTable.rows;
		for (var i=0;i<AllRows.length;i++)
		{
			var EndOfRow=AllRows[i].cells.length-1;
			var Position=EndOfRow-MoveFromEnd;
			if (Position<0) Position=0;
			var AllCellsInRow=AllRows[i].cells;
			if (AllCellsInRow[Position].colSpan>1) AllCellsInRow[position].colSpan=AllCellsInRow[position].colSpan-1;
			else AllRows[i].deleteCell(Position);
		}
	}
}
function MergeColumn()
{
	if (CursorInTableCell())
	{
		var RowSpanTD=SelectedTD.getAttribute('rowSpan');
		AllRows=SelectedTable.rows;
		if (SelectedTR.rowIndex+1!=AllRows.length)
		{
			var AllCellsInNextRow=AllRows[SelectedTR.rowIndex+SelectedTD.rowSpan].cells;
			var AddRowSpan=AllCellsInNextRow[SelectedTD.cellIndex].getAttribute('rowSpan');
			var MoveTo=SelectedTD.rowSpan;
			if (!AddRowSpan) AddRowSpan=1;
			SelectedTD.rowSpan=SelectedTD.rowSpan+AddRowSpan;
			AllRows[SelectedTR.rowIndex+MoveTo].deleteCell(SelectedTD.cellIndex);
		}
		else alert('请重新选择');
	}
	ShowTableBorders();
}
function MergeRow()
{
	if (CursorInTableCell())
	{
		var ColSpanTD=SelectedTD.getAttribute('colSpan');
		var AllCells=SelectedTR.cells;
		if (SelectedTD.cellIndex+1!=SelectedTR.cells.length)
		{
			var AddColspan=AllCells[SelectedTD.cellIndex+1].getAttribute('colSpan');
			SelectedTD.colSpan=ColSpanTD+AddColspan;
			SelectedTR.deleteCell(SelectedTD.cellIndex+1);
		}	
	}
}
function SplitRows()
{
	if (!CursorInTableCell()) return;
	var AddRowsNoSpan=1;
	var NsLeftColSpan=0;
	for (var i=0;i<SelectedTD.cellIndex;i++) NsLeftColSpan+=SelectedTR.cells[i].colSpan;
	var AllRows=SelectedTable.rows;
	while (SelectedTD.rowSpan>1&&AddRowsNoSpan>0)
	{
		var NextRow=AllRows[SelectedTR.rowIndex+SelectedTD.rowSpan-1];
		SelectedTD.rowSpan-=1;
		var NcLeftColSpan=0;
		var Position=-1;
		for (var n=0;n<NextRow.cells.length;n++)
		{
			NcLeftColSpan+=NextRow.cells[n].getAttribute('colSpan');
			if (NcLeftColSpan>NsLeftColSpan)
			{
				Position=n;
				break;
			}
		}
		var NewTD=NextRow.insertCell(Position);
		NewTD.innerHTML="&nbsp;";
		AddRowsNoSpan-=1;
	}
	for (var n=0;n<AddRowsNoSpan;n++)
	{
		var numCols=0
		allCells=SelectedTR.cells
		for (var i=0;i<allCells.length;i++) numCols=numCols+allCells[i].getAttribute('colSpan')
		var newTR=SelectedTable.insertRow(SelectedTR.rowIndex+1)
		for (var j=0;j<SelectedTR.rowIndex;j++)
		{
			for (var k=0;k<AllRows[j].cells.length;k++)
			{
				if ((AllRows[j].cells[k].rowSpan>1)&&(AllRows[j].cells[k].rowSpan>=SelectedTR.rowIndex-AllRows[j].rowIndex+1)) AllRows[j].cells[k].rowSpan+=1;
			}
		}
		for (i=0;i<allCells.length;i++)
		{
			if (i!=SelectedTD.cellIndex) SelectedTR.cells[i].rowSpan+=1;
			else
			{
				NewTD=newTR.insertCell();
				NewTD.colSpan=SelectedTD.colSpan;
				NewTD.innerHTML="&nbsp;";
			}
		}
	}
	ShowTableBorders();
}
function SplitColumn()
{
	if (!CursorInTableCell()) return;
	var AddColsNoSpan=1;
	var NewCell=null;
	var NsLeftColSpan=0;
	var NsLeftRowSpanMoreOne=0;
	for (var i=0;i<SelectedTD.cellIndex;i++)
	{
		NsLeftColSpan+=SelectedTR.cells[i].colSpan;
		if (SelectedTR.cells[i].rowSpan>1) NsLeftRowSpanMoreOne+=1;
	}
	var AllRows=SelectedTable.rows;
	while (SelectedTD.colSpan>1&&AddColsNoSpan>0)
	{
		NewCell=SelectedTR.insertCell(SelectedTD.cellIndex+1);
		NewCell.innerHTML="&nbsp;"
		selectedTD.colSpan-=1;
		AddColsNoSpan-=1;
	}
	for (i=0;i<AllRows.length;i++)
	{
		var ncLeftColSpan=0;
		var position=-1;
		for (var n=0;n<AllRows[i].cells.length;n++)
		{
			ncLeftColSpan+=AllRows[i].cells[n].getAttribute('colSpan');
			if (ncLeftColSpan+NsLeftRowSpanMoreOne>NsLeftColSpan)
			{
				position=n;
				break;
			}
		}
		if (SelectedTR.rowIndex!=i)
		{
			if (position!=-1) AllRows[i].cells[position+NsLeftRowSpanMoreOne].colSpan+=AddColsNoSpan;
		}
		else
		{
			for (var n=0;n<AddColsNoSpan;n++)
			{
				NewCell=AllRows[i].insertCell(SelectedTD.cellIndex+1)
				NewCell.innerHTML="&nbsp;"
				NewCell.rowSpan=SelectedTD.rowSpan;
			}
		}
	}
	ShowTableBorders();
}
function InsertDownLoad(SysDoMain)
{
	var SelectionType=GetEditAreaSelectionType().toLowerCase(),ReturnValue='';
	switch (SelectionType)
	{
		case 'text':
			ReturnValue=OpenWindow('InsertDownLoadFrame.asp?FileName=InsertDownLoad.asp&PageTitle=插入下载',420,180,window);
			if (ReturnValue!='')
			{
				var SelectionObj=EditArea.document.selection.createRange();
				InsertHTMLStr('<a href="'+SysDoMain+'/Down.asp?FileUrl='+ReturnValue+'">'+SelectionObj.text+'</a>');
			}
			break;
		case 'none':
			ReturnValue=OpenWindow('InsertDownLoadFrame.asp?FileName=InsertDownLoad.asp&PageTitle=插入下载',420,180,window);
			if (ReturnValue!='')
			{
				InsertHTMLStr('<a href="'+SysDoMain+'/Down.asp?FileUrl='+ReturnValue+'">下载</a>');
			}
			break;
		default:
			alert('此处不允许插入');
	}
}
function InsertHref(Operation)
{
	EditArea.focus();
	EditArea.document.execCommand(Operation,true);
	EditArea.focus();
}
function Pos()    //有待完善
{
	var ObjReference=null;
	var RangeType=EditArea.document.selection.type;
	if (RangeType!="Control")
	{
		alert('你选择的不是对象！');
		return;
	}
	var SelectedRange= EditArea.document.selection.createRange();
	for (var i=0; i<SelectedRange.length; i++)
	{
		ObjReference = SelectedRange.item(i);
		if (ObjReference.style.position != 'absolute') 
		{
			ObjReference.style.position='absolute';
		}
		else
		{
			ObjReference.style.position='static';
		}
	}
	EditArea.content = false;
}
function CursorInTableCell()
{
	if (EditArea.document.selection.type!="Control")
	{
		var SelectedElement=EditArea.document.selection.createRange().parentElement();
		while (SelectedElement.tagName.toUpperCase()!="TD"&&SelectedElement.tagName.toUpperCase()!="TH")
		{
			SelectedElement=SelectedElement.parentElement;
			if (SelectedElement==null) break;
		}
		if (SelectedElement)
		{
			SelectedTD=SelectedElement;
			SelectedTR=SelectedTD.parentElement;
			SelectedTBODY=SelectedTR.parentElement;
			SelectedTable=SelectedTBODY.parentElement;
			return true;
		}
	}
}
function SearchStr()
{
  var Temp=window.showModalDialog("Search.htm", window, "dialogWidth:320px; dialogHeight:170px; help: no; scroll: no; status: no");
}
function AddHeight()
{
	alert(document.all.EditArea.height);
	EditAreaHeight=EditAreaHeight+60;
	//document.all.EditArea.height=EditAreaHeight;
	parent.document.all.NewsContent.height=parseInt(parent.document.all.NewsContent.height)+60;
}
function MinusHeight()
{
	document.all.EditArea.height=parseInt(BodyHeight-document.all.Toolbar.height-document.all.SetModeArea.height)-60;
	parent.document.all.NewsContent.height=parseInt(parent.document.all.NewsContent.height)-60;
}
//Editer Btn Click Event Function End.
function DisabledAllBtn(Flag)
{
	var AllBtnArray=document.body.getElementsByTagName('IMG'),CurrObj=null;
	for (var i=0;i<AllBtnArray.length;i++)
	{
		CurrObj=AllBtnArray[i];
		if (CurrObj.className=='Btn') CurrObj.disabled=Flag;
	}
	AllBtnArray=document.body.getElementsByTagName('SELECT');
	for (var i=0;i<AllBtnArray.length;i++)
	{
		CurrObj=AllBtnArray[i];
		if (CurrObj.className=='ToolSelectStyle') CurrObj.disabled=Flag;
	}
}

function PureTextPaste()
{
	EditArea.focus();
	var sText = HTMLEncode(clipboardData.getData("Text")) ;
	InsertHTMLStr(sText);
	EditArea.focus();
} 