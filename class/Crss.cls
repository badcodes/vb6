VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Crss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type imageRSS
title As String
link As String
URL As String
End Type

Private Type itemRSS
title As String
link As String
description As String
author As String
category As String
date As String
image As imageRSS
comment As String
End Type

Private Type channelRSS
title As String
link As String
image As imageRSS
description As String
language As String
generator As String
copyright As String
date As String
category As String
itemCount As Integer
End Type

Private Type rssInfo
version As String
channel As channelRSS
End Type

Private Type xmlInfo
version As String
encoding As String
End Type

Private Type RSSXML
xml As xmlInfo
Rss As rssInfo
End Type

Const CDSTART = "<![CDATA["
Const CDEND = "]]>"


Public Function CreateHtmlSource(urlXML As String, spathTMPL As String) As String


CreateHtmlSource = ""
If urlXML = "" Then
    MsgBox "URL Error!"
    Exit Function
End If
If Dir(spathTMPL) = "" Then
    spathTMPL = App.Path + "\" + "rsstmpl.htm"
    If Dir(spathTMPL) = "" Then MsgBox spathTMPL + " Not Found!": Exit Function

End If

Dim docXMl As New DOMDocument

docXMl.Load urlXML

Do
DoEvents
Loop Until docXMl.readyState = 4

If docXMl.hasChildNodes = False Then
MsgBox urlXML + " not appeared to be a rss resource."
Exit Function
End If
'If LCase(docXMl.firstChild.baseName) <> "xml" Then
'MsgBox urlXML + " not appeared to be a rss resource."
'Exit Function
'End If
If docXMl.selectNodes("rss").length < 1 And docXMl.selectNodes("rdf:RDF").length < 1 Then
MsgBox urlXML + " not appeared to be a rss resource."
Exit Function
End If


Dim rRssXML As RSSXML
Dim nlTemp As IXMLDOMNodeList
Dim nTemp As IXMLDOMNode


On Error Resume Next

'READ RSS HEAD
With rRssXML.xml

Set nTemp = docXMl.firstChild
.version = nTemp.Attributes.getNamedItem("version").Text
.encoding = nTemp.Attributes.getNamedItem("encoding").Text
End With


With rRssXML.Rss


Set nlTemp = docXMl.selectNodes("rss")
If nlTemp.length < 1 Then Set nlTemp = docXMl.selectNodes("rdf:RDF")
If nlTemp.length < 1 Then Exit Function
Set nTemp = nlTemp(0)
.version = nTemp.Attributes.getNamedItem("version").Text
End With

'READ CHANNEL HEAD

With rRssXML.Rss.channel

Set nTemp = nTemp.selectSingleNode("channel")
.title = nTemp.selectSingleNode("title").Text
.link = nTemp.selectSingleNode("link").Text

.description = nTemp.selectSingleNode("description").xml
If InStr(.description, "&gt;") > 0 Then
    .description = nTemp.selectSingleNode("description").Text
Else
.description = xmlInnerHtml(.description)
End If

.category = nTemp.selectSingleNode("category").Text
.copyright = nTemp.selectSingleNode("copyright").Text
.generator = nTemp.selectSingleNode("generator").Text
.date = nTemp.selectSingleNode("pubDate").Text
If .date = "" Then .date = nTemp.selectSingleNode("dc:date").Text

.language = nTemp.selectSingleNode("language").Text

Set nlTemp = docXMl.getElementsByTagName("item")

.itemCount = nlTemp.length

End With

With rRssXML.Rss.channel.image

Set nTemp = nTemp.selectSingleNode("image")
.link = nTemp.selectSingleNode("link").Text
.title = nTemp.selectSingleNode("title").Text
.URL = nTemp.selectSingleNode("url").Text
End With

'READ HTML TEMPLATE


Dim i As Integer


Dim elmHtml As IHTMLElement
Dim wbHtml As New HTMLDocument
Dim htmldoc As IHTMLDocument2
Set htmldoc = wbHtml.createDocumentFromUrl(spathTMPL, "")

Do
DoEvents
Loop Until wbHtml.readyState = "complete" And htmldoc.readyState = "complete"

'HTML HEAD
Set elmHtml = Nothing
Set elmHtml = htmldoc.createElement("meta")
elmHtml.setAttribute "http-equiv", "Content-Type"
elmHtml.setAttribute "content", "text/html; charset=" + rRssXML.xml.encoding

htmldoc.All.tags("head")(0).appendChild elmHtml

htmldoc.title = rRssXML.Rss.channel.title

'XML Head
Set elmHtml = Nothing
Set elmHtml = htmldoc.getElementsByTagName("a")("rss-url")
elmHtml.setAttribute "href", urlXML
elmHtml.innerText = urlXML

'RSS HEAD

With rRssXML.Rss

Set elmHtml = Nothing
Set elmHtml = htmldoc.getElementsByTagName("a")("channel-link")
elmHtml.setAttribute "href", .channel.link
elmHtml.innerText = .channel.title

Set elmHtml = Nothing
Set elmHtml = htmldoc.getElementById("channel-desc")
elmHtml.innerHTML = .channel.description

Set elmHtml = Nothing
Set elmHtml = htmldoc.getElementById("channel-date")
elmHtml.innerText = .channel.date

Set elmHtml = Nothing
Set elmHtml = htmldoc.getElementById("channel-category")
elmHtml.innerText = .channel.category

Set elmHtml = Nothing
Set elmHtml = htmldoc.getElementById("channel-copyright")
elmHtml.innerText = .channel.copyright

Set elmHtml = Nothing
Set elmHtml = htmldoc.getElementById("channel-generator")
elmHtml.innerText = .channel.generator

Set elmHtml = Nothing
Set elmHtml = htmldoc.getElementById("channel-image")
If .channel.image.URL = "" Then elmHtml.parentElement.removeChild elmHtml
If .channel.image.URL <> "" Then
    elmHtml.setAttribute "src", .channel.image.URL
    elmHtml.setAttribute "href", .channel.image.link
    elmHtml.innerText = .channel.image.title
End If

Set elmHtml = Nothing
Set elmHtml = htmldoc.getElementById("channel-title")
elmHtml.innerText = .channel.title

Set elmHtml = Nothing
Set elmHtml = htmldoc.getElementById("channel-itemcount")
elmHtml.innerText = .channel.itemCount

End With


Dim iOrd As Integer


Dim itemElement As IHTMLElement

Set elmHtml = Nothing
Set elmHtml = htmldoc.getElementById("item")
Set itemElement = htmldoc.createElement("DIV")

itemElement.innerHTML = elmHtml.innerHTML
itemElement.className = elmHtml.className

elmHtml.className = ""
elmHtml.innerHTML = ""


For iOrd = 1 To rRssXML.Rss.channel.itemCount


Dim iItemRss As itemRSS

With iItemRss

.title = ""
.link = ""
.description = ""
.author = ""
.category = ""
.date = ""
.image.link = ""
.image.title = ""
.image.URL = ""
.comment = ""

Set nTemp = nlTemp(iOrd - 1)

.author = nTemp.selectSingleNode("author").Text
.category = nTemp.selectSingleNode("category").Text
.comment = nTemp.selectSingleNode("comment").Text
.description = nTemp.selectSingleNode("description").xml
If InStr(.description, "&gt;") > 0 Then
    .description = nTemp.selectSingleNode("description").Text
Else
.description = xmlInnerHtml(.description)
End If
.link = nTemp.selectSingleNode("link").Text
.date = nTemp.selectSingleNode("pubDate").Text
If .date = "" Then .date = nTemp.selectSingleNode("dc:date").Text
.title = nTemp.selectSingleNode("title").Text

Set nTemp = nTemp.selectSingleNode("image")

.image.link = nTemp.selectSingleNode("link").Text
.image.title = nTemp.selectSingleNode("title").Text
.image.URL = nTemp.selectSingleNode("url").Text


Set elmHtml = Nothing
Set elmHtml = itemElement.All("item-link")
elmHtml.setAttribute "href", .link
elmHtml.innerText = .title

Set elmHtml = Nothing
Set elmHtml = itemElement.All("item-desc")
elmHtml.innerHTML = .description

Set elmHtml = Nothing
Set elmHtml = itemElement.All("item-date")
elmHtml.innerText = .date

Set elmHtml = Nothing
Set elmHtml = itemElement.All("item-category")
elmHtml.innerText = .category

Set elmHtml = Nothing
Set elmHtml = itemElement.All("item-author")
elmHtml.innerText = .author

Set elmHtml = Nothing
Set elmHtml = itemElement.All("item-comment")
elmHtml.innerText = .comment

Set elmHtml = Nothing
Set elmHtml = itemElement.All("item-image")
If .image.URL = "" Then elmHtml.parentElement.removeChild elmHtml
If .image.URL <> "" Then
    elmHtml.setAttribute "src", .image.URL
    elmHtml.setAttribute "href", .image.link
    elmHtml.innerText = .image.title
End If


Set elmHtml = Nothing
Set elmHtml = itemElement.All("item-title")
elmHtml.innerText = .title

Set elmHtml = Nothing
Set elmHtml = itemElement.All("item-ordinal")
elmHtml.innerText = LTrim(Str(iOrd))


End With


Set elmHtml = Nothing
Set elmHtml = htmldoc.getElementById("item")

elmHtml.insertAdjacentHTML "BeforeEnd", itemElement.outerHTML

Next

CreateHtmlSource = htmldoc.firstChild.outerHTML

End Function

Private Function xmlInnerHtml(strXML As String)
If strXML = "" Then Exit Function
strXML = LTrim(strXML)
strXML = RTrim(strXML)
Dim pStart As Integer
Dim pEnd As Integer
If Left(strXML, 1) = "<" Then pStart = InStr(strXML, ">")
If pStart < 0 Then pStart = 0
pEnd = Len(strXML)
If Right(strXML, 1) = ">" Then pEnd = InStrRev(strXML, "<")
If pEnd < 0 Then pEnd = 0
If pStart >= pEnd Then
    xmlInnerHtml = strXML
Else
    xmlInnerHtml = Mid(strXML, pStart + 1, pEnd - pStart - 1)
End If

xmlInnerHtml = Replace(xmlInnerHtml, CDSTART, "")
xmlInnerHtml = Replace(xmlInnerHtml, CDEND, "")

End Function


