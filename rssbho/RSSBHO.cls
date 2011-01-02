VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RSS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type imageRSS
title As String
link As String
url As String
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
RSS As rssInfo
End Type


'Const nameCss = "style.css"
Const nameHtml = "rss.htm"
Const nameTemplate = "rsstmpl.htm"
'Const nameXml = "Xml.txt"
'Const nameXmlhtml = "Xml.htm"

Dim spathHtml As String
'Dim spathXml As String
'Dim spathXmlHtml As String
'Dim spathCss As String
Dim spathTmpl As String

Implements IObjectWithSiteTLB.IObjectWithSite
Dim WithEvents m_ie As InternetExplorer
Attribute m_ie.VB_VarHelpID = -1
Dim m_Site As IUnknownVB
Const CDSTART = "<![CDATA["
Const CDEND = "]]>"





Private Sub IObjectWithSite_GetSite(ByVal priid As IObjectWithSiteTLB.GUIDPtr, ppvObj As IObjectWithSiteTLB.VOIDPtr)
    m_Site.QueryInterface priid, ppvObj
End Sub

Private Sub IObjectWithSite_SetSite(ByVal pSite As IObjectWithSiteTLB.IUnknownVB)
    Set m_Site = pSite
    Set m_ie = pSite
End Sub
Public Sub xml_rschange(tagrs As tagREADYSTATE, urlXml As Variant)
Select Case tagrs
 Case READYSTATE_LOADING
    m_ie.StatusText = "Loading " + urlXml
 Case READYSTATE_LOADED
    m_ie.StatusText = urlXml + " Loaded."
 Case READYSTATE_INTERACTIVE
    m_ie.StatusText = urlXml + " be Interative."
 Case READYSTATE_COMPLETE
    m_ie.StatusText = "LiNRss BHO V" + Str(App.Major) + "." + Str(App.Major) + "." + Str(App.Revision)
 End Select

End Sub


Private Function viewrss(MYIE As InternetExplorer, urlXml As Variant, flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant) As Boolean

Dim docXMl As New DOMDocument



docXMl.Load urlXml

Dim tagrs As tagREADYSTATE
Dim oldtag As tagREADYSTATE
Do
DoEvents
tagrs = docXMl.readyState
If oldtag <> tagrs Then
    xml_rschange tagrs, urlXml
    oldtag = tagrs
End If
Loop Until tagrs = READYSTATE_COMPLETE


If docXMl.hasChildNodes = False Then viewrss = False: Exit Function
If docXMl.selectNodes("rss").length < 1 And docXMl.selectNodes("rdf:RDF").length < 1 Then viewrss = False: Exit Function


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


With rRssXML.RSS


Set nlTemp = docXMl.selectNodes("rss")
If nlTemp.length < 1 Then Set nlTemp = docXMl.selectNodes("rdf:RDF")
If nlTemp.length < 1 Then Exit Function
Set nTemp = nlTemp(0)
.version = nTemp.Attributes.getNamedItem("version").Text
End With

'READ CHANNEL HEAD

With rRssXML.RSS.channel

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

With rRssXML.RSS.channel.image

Set nTemp = nTemp.selectSingleNode("image")
.link = nTemp.selectSingleNode("link").Text
.title = nTemp.selectSingleNode("title").Text
.url = nTemp.selectSingleNode("url").Text
End With

'READ HTML TEMPLATE

Dim fso As New FileSystemObject
Dim ts As TextStream

Dim i As Integer

spathTmpl = fso.BuildPath(App.Path, nameTemplate)
spathHtml = fso.BuildPath(App.Path, nameHtml)
'spathCss = fso.BuildPath(App.Path, nameCss)
'spathXml = fso.BuildPath(App.Path, nameXml)
'spathXmlHtml = fso.BuildPath(App.Path, nameXmlhtml)

If fso.FileExists(spathTmpl) = False Then viewrss = False: Exit Function
'Set ts = fso.OpenTextFile(spathXml, ForWriting, True)
'ts.write docXMl.xml
'ts.Close
'Set ts = fso.OpenTextFile(spathXmlHtml, ForWriting, True)
'ts.write m_ie.document.firstChild.innerHTML
'ts.Close



Dim elmHtml As IHTMLElement
Dim wbHtml As New HTMLDocument
Dim htmldoc As IHTMLDocument2
Set htmldoc = wbHtml.createDocumentFromUrl(spathTmpl, "")
Do
DoEvents
Loop Until wbHtml.readyState = "complete" And htmldoc.readyState = "complete"

'HTML HEAD
Set elmHtml = Nothing
Set elmHtml = htmldoc.createElement("meta")
elmHtml.setAttribute "http-equiv", "Content-Type"
elmHtml.setAttribute "content", "text/html; charset=" + rRssXML.xml.encoding
htmldoc.title = rRssXML.RSS.channel.title + " - " + aboutapp
htmldoc.All.tags("head")(0).appendChild elmHtml

'Set elmHtml = Nothing
'Set elmHtml = htmldoc.createStyleSheet(spathCss)
'htmldoc.All.tags("head")(0).appendChild elmHtml
'htmldoc.title = rRssXML.rss.channel.title

'XML Head
Set elmHtml = Nothing
Set elmHtml = htmldoc.getElementsByTagName("a")("rss-url")
elmHtml.setAttribute "href", urlXml
elmHtml.innerText = urlXml

'Set elmHtml = Nothing
'Set elmHtml = htmldoc.getElementsByTagName("a")("xml-source")
'elmHtml.setAttribute "href", spathXml
'elmHtml.innerText = "XML"

'RSS HEAD

With rRssXML.RSS

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
If .channel.image.url = "" Then elmHtml.parentElement.removeChild elmHtml
If .channel.image.url <> "" Then
    elmHtml.setAttribute "src", .channel.image.url
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


For iOrd = 1 To rRssXML.RSS.channel.itemCount


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
.image.url = ""
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
.image.url = nTemp.selectSingleNode("url").Text

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
If .image.url = "" Then elmHtml.parentElement.removeChild elmHtml
If .image.url <> "" Then
    elmHtml.setAttribute "src", .image.url
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


Dim fnum As Integer
fnum = FreeFile
Open spathHtml For Output As fnum
Print #fnum, htmldoc.firstChild.outerHTML
Close #fnum


MYIE.Navigate2 spathHtml, flags, TargetFrameName, PostData, Headers
viewrss = True

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

Private Sub m_ie_BeforeNavigate2(ByVal pDisp As Object, url As Variant, flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
If LCase(Right(url, 8)) = "|rssfeed" Then
    Cancel = True
    Dim urlXml As String
    Dim tfRss As Boolean
    urlXml = Left(url, Len(url) - 8)
    tfRss = viewrss(m_ie, urlXml, flags, TargetFrameName, PostData, Headers)
    If tfRss = False Then m_ie.Navigate2 urlXml, flags, TargetFrameName, PostData, Headers
End If
End Sub

Function aboutapp() As String
aboutapp = App.ProductName + " V" + Str(App.Major) + "." + Str(App.Minor) + " By " + App.LegalCopyright + " " + App.CompanyName
End Function

