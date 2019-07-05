//Open Window
function OpenWindow(Url,Width,Height,WindowObj)
{
	var ReturnStr=showModalDialog(Url,WindowObj,'dialogWidth:'+Width+'pt;dialogHeight:'+Height+'pt;status:no;help:no;scroll:no;');
	return ReturnStr;
}
//Open Modal Window
function OpenWindowAndSetValue(Url,Width,Height,WindowObj,SetObj)
{
	var ReturnStr=showModalDialog(Url,WindowObj,'dialogWidth:'+Width+'pt;dialogHeight:'+Height+'pt;status:no;help:no;scroll:no;');
	if (ReturnStr!='007007007007') SetObj.value=ReturnStr;
	return ReturnStr;
}
//Open Editer Window
function OpenEditerWindow(Url,WindowName,Width,Height)
{
	window.open(Url,WindowName,'toolbar=0,location=0,maximize=1,directories=0,status=1,menubar=0,scrollbars=0,resizable=1,top=50,left=50,width='+Width+',height='+Height);
}
//Send Data To Server
function SendDataToServer(Url)
{
	var HTTP = new ActiveXObject("Microsoft.XMLHTTP");
	var ReturnValue=HTTP.open("POST", Url, false);
	HTTP.send("");
	return HTTP.responseText;
}
//Button MouseOver Event
function BtnMouseOver(Obj)
{
	if (event.type!='mouseout')
	{
		Obj.className='BtnMouseOver';
		if (Obj.tagName.toLowerCase()=='td' || Obj.tagName.toLowerCase()=='img') window.status=Obj.alt;
		else window.status=Obj.title;
	}
	else
	{
		window.status=top.LoginStr;
		Obj.className='BtnMouseOut';
	}
}
//Check number or not and alarm user.
function CheckNumber(Obj,DescriptionStr)
{
	if (Obj.value!='' && (isNaN(Obj.value) || Obj.value<0))
	{
		alert(DescriptionStr+"应填有效数字！");
		Obj.value="";
		Obj.focus();
	}
}
//Check English Str
function CheckEnglishStr(Obj,DescriptionStr)
{
	var TempStr=Obj.value,i=0,ErrorStr='',CharAscii;
	if (TempStr!='')
	{
		for (i=0;i<TempStr.length;i++)
		{
			CharAscii=TempStr.charCodeAt(i);
			if (CharAscii>=255||CharAscii<=31)
			{
				ErrorStr=ErrorStr+TempStr.charAt(i);
			}
			else
			{
				if (!CheckClassErrorStr(CharAscii))
				{
					ErrorStr=ErrorStr+TempStr.charAt(i);
				}
			}
		}
		if (ErrorStr!='')
		{
			alert(DescriptionStr+'发现非法字符:'+ErrorStr);
			Obj.focus();
			return false;
		}
		if (!(((TempStr.charCodeAt(0)>=48)&&(TempStr.charCodeAt(0)<=57))||((TempStr.charCodeAt(0)>=65)&&(TempStr.charCodeAt(0)<=90))||((TempStr.charCodeAt(0)>=97)&&(TempStr.charCodeAt(0)<=122))))
		{
			alert(DescriptionStr+'首字符只能够为数字或者字母');
			Obj.focus();
			return false;
		}
	}
	return true;
}
function CheckClassErrorStr(CharAsciiCode)
{
	var TempArray=new Array(34,47,92,42,58,60,62,63,124);
	for (var i=0;i<TempArray.length;i++)
	{
		if (CharAsciiCode==TempArray[i]) return false;
	}
	return true;
}

//
function ChooseSpecial(Special)
{
	var TempArray,TempStr;
	TempArray=Special.split("***");
	if (TempArray[0] != '')
	{
		if (document.all.SpecialID.value.search(TempArray[1])==-1)
	
		   {		if (document.all.SpecialIDText.value=='') 	document.all.SpecialIDText.value=TempArray[0];
					else document.all.SpecialIDText.value = document.all.SpecialIDText.value + ',' + TempArray[0];
					if (document.all.SpecialID.value=='') 	document.all.SpecialID.value=TempArray[1];
					else document.all.SpecialID.value = document.all.SpecialID.value + ',' + TempArray[1];
			}
	}
	if ((TempArray[0] == '')&&(TempArray[1] == 'Clean'))
	{
		document.all.SpecialID.value = '';
		document.all.SpecialIDText.value = '';
	}
	return;
}

function Dosusite(Source)
{
	var TempArray,TempStr;
	TempArray=Source.split("***");
	if (TempArray[0] != '')
	{
		if (document.NewsForm.TxtSourceText.value.indexOf(TempArray[0])<0)
		{
			if (typeof(TempArray[1])=='undefined') TempStr=TempArray[0];
			else TempStr='<a href='+TempArray[1].replace(/[\"\']/,'')+'>'+TempArray[0]+'</a>';
			if (document.NewsForm.TxtSourceText.value=='') 	document.NewsForm.TxtSourceText.value=TempArray[0];
			else document.NewsForm.TxtSourceText.value = document.NewsForm.TxtSourceText.value + ',' + TempArray[0];
			if (document.NewsForm.TxtSource.value=='') 	document.NewsForm.TxtSource.value=TempArray[0];
			else document.NewsForm.TxtSource.value = document.NewsForm.TxtSource.value + ',' + TempArray[0];
		}
	}
	if ((TempArray[0] == '')&&(TempArray[1] == 'Clean'))
	{
		document.NewsForm.TxtSource.value = '';
		document.NewsForm.TxtSourceText.value = '';
	}
	return;
}

function Dokesite(KeyWords)
{
	if (KeyWords!='')
	{
		if (document.NewsForm.KeywordText.value.search(KeyWords)==-1)
		{
			if (document.NewsForm.KeyWords.value=='') document.NewsForm.KeyWords.value=KeyWords;
			else document.NewsForm.KeyWords.value=document.NewsForm.KeyWords.value+','+KeyWords;
			if (document.NewsForm.KeywordText.value=='') document.NewsForm.KeywordText.value=KeyWords;
			else document.NewsForm.KeywordText.value=document.NewsForm.KeywordText.value+','+KeyWords;
		}
	}
	if (KeyWords == 'Clean')
	{
		document.NewsForm.KeyWords.value = '';
		document.NewsForm.KeywordText.value = '';
	}
	return;
}

function Doauthsite(Author)
{
	var TempArray,TempStr;
	TempArray=Author.split("***");
	if (TempArray[0] != '')
	{
		if (document.NewsForm.AuthorText.value.indexOf(TempArray[0])<0)
		{
			if (typeof(TempArray[1])=='undefined') TempStr=TempArray[0];
			else TempStr='<a href='+TempArray[1].replace(/[\"\']/,'')+'>'+TempArray[0]+'</a>';
			if (document.NewsForm.AuthorText.value=='') 	document.NewsForm.AuthorText.value=TempArray[0];
			else document.NewsForm.AuthorText.value = document.NewsForm.AuthorText.value + ',' + TempArray[0];
			if (document.NewsForm.Author.value=='') 	document.NewsForm.Author.value=TempArray[0];
			else document.NewsForm.Author.value = document.NewsForm.Author.value + ',' + TempArray[0];
		}
	}
	if ((TempArray[0] == '')&&(TempArray[1] == 'Clean'))
	{
		document.NewsForm.Author.value = '';
		document.NewsForm.AuthorText.value = '';
	}
	return;
}

function Editsite(Editer1)
{
	var TempArray,TempStr;
	TempArray=Editer1.split("***");
	if (TempArray[0] != '')
	{
		if (document.NewsForm.EditerText.value.indexOf(TempArray[0])<0)
		{
			if (typeof(TempArray[1])=='undefined') TempStr=TempArray[0];
			else TempStr='<a href='+TempArray[1].replace(/[\"\']/,'')+'>'+TempArray[0]+'</a>';
			if (document.NewsForm.EditerText.value=='') 	document.NewsForm.EditerText.value=TempArray[0];
			else document.NewsForm.EditerText.value = document.NewsForm.EditerText.value + ',' + TempArray[0];
			if (document.NewsForm.Editer.value=='') 	document.NewsForm.Editer.value=TempArray[0];
			else document.NewsForm.Editer.value = document.NewsForm.Editer.value + ',' + TempArray[0];
		}
	}
	if ((TempArray[0] == '')&&(TempArray[1] == 'Clean'))
	{
		document.NewsForm.Editer.value = '';
		document.NewsForm.EditerText.value = '';
	}
	return;
}

function ChooseSystem(DownSystem)
{
	if (DownSystem != '')
		{	
			if (document.DownForm.SystemType.value.search(DownSystem)==-1)
			{
				if (document.DownForm.SystemType.value=='') document.DownForm.SystemType.value=DownSystem;
				else document.DownForm.SystemType.value = document.DownForm.SystemType.value + '/' + DownSystem;
			}
		}
	if (DownSystem == 'Clean') document.DownForm.SystemType.value = '';
	return;
}

//////////////////////////////////////////////////////////////////
var MouseOverObj=null;
var MouseOverPageLocation='';
function document.onmouseover()
{
	MouseOverObj=event.srcElement;
	MouseOverPageLocation=location.href;
	//var DocumentBodyObj=MouseOverObj;
	//while ((DocumentBodyObj.parentElement)&&(DocumentBodyObj.tagName!='BODY')) DocumentBodyObj=DocumentBodyObj.parentElement;
	//if ((DocumentBodyObj)&&(DocumentBodyObj.tagName=='BODY')) DocumentBodyObj.focus();
}
function document.onkeydown()
{
	var ParentObj=null,OverPageLocationStr='',Loc=0,HrefStr='',KeyWord='';
	if (top.dialogArguments) {ParentObj=top.dialogArguments.top.GetFSHelpObject();}
	else ParentObj=top.GetFSHelpObject();
	if (!ParentObj) return;
	HrefStr=ParentObj.location.href;
	if (event.ctrlKey==true)
	{
		if (!((event.keyCode==49)||(event.keyCode==97))) return;
		var DocSelObj=document.selection;
		if (DocSelObj.type=='Text') KeyWord=DocSelObj.createRange().text;
		else KeyWord=escape(AnalyKeyWord());
		Loc=MouseOverPageLocation.lastIndexOf('?');
		if (Loc!=-1) OverPageLocationStr=MouseOverPageLocation.slice(0,Loc);
		else OverPageLocationStr=MouseOverPageLocation;
		Loc=OverPageLocationStr.lastIndexOf('/');
		if (Loc!=-1) OverPageLocationStr=OverPageLocationStr.slice(Loc+1,OverPageLocationStr.length);
		OverPageLocationStr=escape(OverPageLocationStr);
		Loc=HrefStr.lastIndexOf('?');
		if (Loc==-1) HrefStr=HrefStr+'?KeyWord='+KeyWord+'&Page='+OverPageLocationStr;
		else
		{
			HrefStr=HrefStr.slice(0,HrefStr.lastIndexOf('?'))+'?KeyWord='+KeyWord+'&Page='+OverPageLocationStr;
		}
		ParentObj.location.href=HrefStr;
	}
}

function AnalyKeyWord()
{
	var returnValue='',TempObj=MouseOverObj;
	returnValue=GetKeyWord(MouseOverObj,0);
	if (returnValue=='')
	{
		while ((TempObj.children)&&(TempObj.children.length==1)) TempObj=TempObj.children(0);
		return GetKeyWord(TempObj,1);
	}
	else return returnValue;
}

function GetKeyWord(Obj,flag)
{
	var TagString='',returnValue='';
	if (!Obj) return;
	TagString=Obj.tagName;
	switch (TagString)
	{
		case 'INPUT':
			if (Obj.type=='button') returnValue=Obj.value;
			else
			{
				if (Obj.id) returnValue=Obj.id;
				else returnValue=Obj.name;
			}
			break;
		case 'SELECT':
			if (Obj.id) returnValue = Obj.id;
			else returnValue = Obj.name;
			break;
		case 'TEXTAREA':
			if (Obj.id) returnValue = Obj.id;
			else returnValue = Obj.name;
			break;
		case 'IMG':
			if (Obj.alt) returnValue = Obj.alt;
			else returnValue = Obj.title;
			break;
		default :
			if (flag) returnValue=Obj.innerText;
			else returnValue='';
			break;
	}
	return returnValue;
}