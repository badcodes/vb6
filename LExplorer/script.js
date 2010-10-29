var currentpos,timer;
function initialize()
{ 
timer=setInterval("scrollwindow()",50);
}
function sc(){
clearInterval(timer);
}
function scrollwindow()
{
currentpos=document.body.scrollTop;
window.scroll(0,++currentpos);
if(currentpos!=document.body.scrollTop)
sc();
}
document.onclick=sc
document.ondblclick=initialize

function click(e) {
if (document.all) {
if (event.button==2||event.button==3) {
setInterval("window.status=''",10);
}
}
if (document.layers) {
if (e.which == 3) {
setInterval("window.status=''",10);
}
}
}
if (document.layers) {
document.captureEvents(Event.MOUSEDOWN);
}
document.onmousedown=click;