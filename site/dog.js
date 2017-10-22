

var agent = navigator.userAgent.toLowerCase();
var is_ns  = ((agent.indexOf('mozilla')!=-1) && ((agent.indexOf('spoofer')==-1) && (agent.indexOf('compatible') == -1)));
var is_ie   = (agent.indexOf("msie") != -1);


var xDog =0;
var yDog =0;

var stepDog =12;
var xMouse  = 1000 ;
var yMouse  = 0;

var flag    = 1
var L_DOG = 73
var dxDog
var dyDog
var dxMouse
var dyMouse

var _all = '';
var _style = '';
if(!is_ns) 
{
 _all='all.';
 _style='.style';
}

var blockD=eval('document.'+_all+'blockD'+_style);
var blockG=eval('document.'+_all+'blockG'+_style);


function init()
{
	if(document.layers)
	{
		blockD.visibility="visible";
		blockG.visibility="visible";
	}
	if(document.all && !(document.getElementById))
	{
		blockD.visibility="visible";
		blockG.visibility="visible";
	}
	if(document.getElementById && is_ns)
	{
		blockG=document.getElementById('blockG');
		blockG.style.visibility="visible";
		blockD=document.getElementById('blockD');
		blockD.style.visibility="visible";
	}
	if(document.getElementById && is_ie)
	{
		blockG=document.getElementById('blockG');
		blockG.style.visibility="visible";
		blockD=document.getElementById('blockD');
		blockD.style.visibility="visible";
	}
}



function dogLeft (object)
{
  if (is_ns) return object.left
  else return  object.pixelLeft
}
function dogTop (object)
{
  if (is_ns) return object.top
  else return  object.pixelTop
}

function move (ob,x,y)
{

	if (document.getElementById)
	{
	  ob.style.left = x ;
	  ob.style.top  = y ;
	 }
	 else if (is_ns) 
	 {
	  ob.moveTo( x, y);
	 } 
	 else 
	 {
	  ob.pixelLeft = x ;
	  ob.pixelTop  = y ;
	 }
}

function show(ob,bShow)
{
	if (document.getElementById)
	 {
	  if (bShow)   ob.style.visibility = "VISIBLE" ;
	  else ob.style.visibility  = "HIDDEN" ;
	 }
	 else 
	 {
	  if (bShow)   ob.visibility = "VISIBLE" ;
	  else ob.visibility  = "HIDDEN" ;
	}
}

function animateDog() 
{

   if ( (xDog <= xMouse + stepDog/2) && (xDog >= xMouse - stepDog/2) 
      && (yDog  <= yMouse + stepDog/2) && (yDog  >= yMouse - stepDog/2)  ) 
    {
     	//ordre = "Pied" ;
    }
    else
    {
	dxMouse = xDog-xMouse 
	dyMouse = yDog-yMouse 
	
	dxDog = Math.sqrt((stepDog*stepDog*dxMouse*dxMouse)/(dxMouse*dxMouse+dyMouse*dyMouse))
	dyDog = Math.sqrt((stepDog*stepDog*dyMouse*dyMouse)/(dyMouse*dyMouse+dxMouse*dxMouse))



	xDog = (xDog <= xMouse) ? xDog+dxDog  : xDog-dxDog  
        yDog = (yDog <= yMouse ) ? yDog+dyDog  : yDog-dyDog  

	

       // Change direction 
       if ( (xMouse <= xDog+stepDog ) && (xMouse >= xDog-stepDog )) 
	{
	
	show (blockD,0);
	show (blockG,1);
	
		  
	}
	else
	{
		if (xDog+stepDog < xMouse ) 
		{
			show (blockG,0);
			show (blockD,1);
		}
		else if (xDog-stepDog > xMouse ) 
		{
			show (blockD,0);
			show (blockG,1);
		}
	}
     }
  	
     move (blockD,xDog,yDog);
     move (blockG,xDog,yDog);


  id1 = setTimeout('animateDog()', 30);
}




 function moveHandler (evt)
 {
   if (document.layers )
   {
       xMouse = evt.pageX +20;
       yMouse = evt.pageY +20;
    }
   
   if(document.all && !(document.getElementById))
   {
       xMouse = event.x + 20;
       yMouse = event.y + 20;
   }
   if (document.getElementById)
   {
       xMouse = window.event.x + document.body.scrollLeft+20;
       yMouse = window.event.y + document.body.scrollTop+20;
   }
 }
 
 if (document.layers)
     document.captureEvents(Event.MOUSEMOVE);
   document.onmousemove = moveHandler;



window.onresize=init;
init();
animateDog(); 


var is_open=false
function open_pict(pct,w,h)
{
	w=w+25
	h=h+30
	if(is_open)
	{
		winEx.close();
	}
		winEx = window.open(pct,"winEx","width="+w+",height="+h+", status=yes, resizable=no,scrollbars=no , BGCOLOR=#000000");
	is_open = true;
}