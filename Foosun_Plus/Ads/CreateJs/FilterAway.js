var lstart=0
var loop=true
var speed=50 
var pr_step=3 
var newspeed=800
var newspeed2=0

function makeObj(obj,nest)
{
	nest=(!nest) ? '':'document.'+nest+'.'
	this.css=(document.layers) ? eval(nest+'document.'+obj):eval(obj+'.style')
	this.scrollHeight=(document.layers) ? 
 	this.css.document.height:eval(obj+'.offsetHeight')
	this.scrollWidth=(document.layers) ? 
 	this.css.document.width:eval(obj+'.offsetWidth')
	this.up=goUp
	this.obj = obj + "Object"
	eval(this.obj + "=this")
	return this
}

function goUp(speed)
{
	if(parseInt(this.css.top)>-(this.scrollHeight-0))
	{
		this.css.top=parseInt(this.css.top)-pr_step-1
		setTimeout(this.obj+".up("+speed+")",35)
	}
	else
	 {
 		if(navigator.appName == "Netscape")
		{
		tome=setInterval(this.obj+".setClipne()",50)
		}
		else
		{
		tome=setInterval('setClipie()',50)
		tmp=FilterAwayT.style.clip;
		}
		
	}
}

function setClipne()
{
	if(temp==0)
	{
	clearInterval(tome);
	document.FilterAwayT.document.FilterAwayF.visibility="hide";
	document.FilterAwayT.visibility="hide";
	}
}

function setClipie()
{
	newspeed=newspeed-pr_step;
	newspeed2=newspeed2+pr_step;
	temp="rect(0px "+newspeed+"px 600px "+newspeed2+"px)";
	this.css.clip=temp;
	if(newspeed<newspeed2)
	{
 	clearInterval(tome);
 	FilterAwayF.style.display="none"
 	FilterAwayT.style.display="none"
 	}
}

function slideInit()
{
	oSlide=makeObj('FilterAwayF','FilterAwayT')
	oSlide.css.top=lstart
	oSlide.up(speed)
}
function myload()
{
setTimeout("slideInit()",4000);
}

myload();