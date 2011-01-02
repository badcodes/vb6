<!--
	var datafile="D:\\WorkBench\\VB\\RSSBHO\\feeds.txt";
	var myfeeds=new Array();

	function feed(href,title){
	this.href=href;
	this.title=title;
	}

	function getfeed(Datafile){
	myfeeds.length=0;
	var fso = new ActiveXObject("Scripting.FileSystemObject");
	if (!fso.FileExists(Datafile)) {alert(Datafile+"not found!");return;}
	var ts=fso.OpenTextFile(Datafile,1);
	var tmpstr,re;
	var i=0;
	while (!ts.AtEndOfStream){
		tmpstr=ts.ReadLine();
		re = new RegExp("^(.*)\,(.*)$","ig");
		re.exec(tmpstr);
		myfeeds[i]=new feed(RegExp.$1,RegExp.$2);
		i=i+1;
		}
	ts.close();
	}

function loadfeed() {
	var feedscount=myfeeds.length;
	var fd=document.getElementById("feedtable");
	var ftHtml;
	if (feedscount>0) ftHtml='<table  border="0" cellpadding="4" width="100%" style="border-collapse: collapse" bordercolor="#111111" cellspacing="0">';
	for (i=0;i<feedscount;i++){
		var fdid=i%3;
		var ih='<td valign="top" align="right"><a href="'+myfeeds[i].href+'" id="feed'+i+'" class="rss-url" > '+myfeeds[i].title+'</a></td><td valign="top" align="right"><img class="imgcmd" src="images\\edit.gif" onclick="editfeed(\'feed'+i+'\')"><img class="imgcmd" src="images\\remove.gif" onclick="removefeed(\'feed'+i+'\')"></td>';
		if (fdid==0) ih="<tr>"+ih;
		if (fdid==2) ih=ih+"</tr>";
		ftHtml=ftHtml+ih;
	}
	if (feedscount>0) fd.innerHTML=ftHtml+'</table>';
}
function savefeed(dstfile) {
	var fd;
	var al;
	var als;
	var fso = new ActiveXObject("Scripting.FileSystemObject");
	var ts=fso.OpenTextFile(dstfile,2,true);
		fd=document.getElementById("feedtable");
		als=fd.getElementsByTagName('a');
		for (j=0;j<als.length;j++) {
			ts.WriteLine ("rssfeed:"+als(j).href+","+als(j).innerText);
		}
	ts.Close();
}
function editfeed(feedid) {
	var al=document.getElementById(feedid);
	var nfeed=new feed(al.href,al.innerText);
	nfeed.href=window.prompt("Href of the Feed",nfeed.href);
	nfeed.title=window.prompt("Title of the Feed",nfeed.title);
	al.href=nfeed.href;
	al.innerText=nfeed.title;
	savefeed(datafile);
}
function removefeed(feedid){
	var al=document.getElementById(feedid);
	al.parentElement.removeChild(al);
	savefeed(datafile);
	getfeed(datafile);
	loadfeed();

}
function addfeed(){
	var nfeed=new feed();
	nfeed.href=window.prompt("Href of the Feed");
	nfeed.title=window.prompt("Title of the Feed");
	if ((nfeed.href == null)||(nfeed.title == null)) return;
	var fso=new ActiveXObject("Scripting.FileSystemObject");
	var ts=fso.OpenTextFile(datafile,8,true);
	ts.WriteLine (nfeed.href+","+nfeed.title);
	ts.Close();
	getfeed(datafile);
	loadfeed();
}

function SwitchDisplay(){
	var ftb=document.getElementById("feedtable");
	if (ftb.className=="show") ftb.className="hide";
	else ftb.className="show";
	window.event.returnValue=false;
}

//-->
