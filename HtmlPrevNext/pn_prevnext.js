
var pnFileList = new Array();
//,"test1.htm","test2.htm","test3.htm","test4.htm");
var pnTopFile = "";
var pnImagePrev = "pn_prev.gif";
var pnImageNext	= "pn_next.gif";
var pnImageTop = "pn_top.gif";
var pnHRColor = "#ff9900";


var pnImageWidth = "34";
var pnImageHeight = "18";
var pnTableStart = "<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 border=0 width=100%>";
var pnTDStart = "<TD width=34 align=right >";
var pnTDEnd = "</TD>";
var pnTDHR = "<TD height=3 align=center colspan=4><HR color=#ff9900 size=2></TD>";

function pnNewLinkTD(href,img) {
   return  pnTDStart + "<A href='" + href + "'><Img border=0 src='" + img + "' vspace=0 hspace=0></a>" + pnTDEnd;
}
function pnInitFileList() {
    for(var i=0;i<arguments.length;i++) {
        pnFileList.push(arguments[i]);
    }
}
function pnInitTopFile(filename) {
    pnTopFile = filename;
}
function pnInitTable() {
        var filename = document.location.href.replace(/^.*\/\//g,"");
	var index = -1;
	for(var i=0;i<pnFileList.length;i++) {
            if(filename == pnFileList[i]) {
                index = i;
                break;
            }
	}
        if(index == -1) {
            var newName = document.location.href.replace(/^.*\//g,"");
            if(newName != filename) {
                for(var i=0;i<pnFileList.length;i++) {
                    if(newName == pnFileList[i]) {
                        index = i;
                        break;
                    }
                }
            }
        }
        if(index <0 || index>pnFileList.length-1) 
            return;

        var TRLinks = "<TR VALIGN=top ALIGN=right><TD align=right>&nbsp;</TD>";	
        TRLinks += "<TD VALIGN=bottom ALIGN=right>[" + (index+1) + "/" + pnFileList.length + "]</TD>";
        if(index>0) 
            TRLinks += pnNewLinkTD(pnFileList[index-1],pnImagePrev);
	if(pnTopFile) 
            TRLinks += pnNewLinkTD(pnTopFile,pnImageTop);
        if(index<pnFileList.length-1) 
            TRLinks += pnNewLinkTD(pnFileList[index+1],pnImageNext);
        TRLinks += "</TR>";
        
        var tableTop = document.createElement("div");
        tableTop.innerHTML = pnTableStart + TRLinks + "<TR VALIGN=top>" + pnTDHR + "</TR></TABLE>";
        var tableBot = document.createElement("div");
        tableBot.innerHTML = pnTableStart + "<TR VALIGN=bottom>" + pnTDHR + "</TR>" + TRLinks + "</TABLE>";

        document.body.insertBefore(tableTop,document.body.firstChild);
        document.body.appendChild(tableBot);
}

//window.addEventListener("load",pnInitTable,false);
//document.attachEvent("onload",pnInitTable);
window.onload=pnInitTable;




