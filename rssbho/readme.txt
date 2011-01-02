  


如何打印: 请在您的浏览器中选择 文件 菜单中的 打印 选项
--------------------------------------------------------------
此文章打印自ZDNet China。
--------------------------------------------------------------

用浏览器辅助对象来增加功能
Builder.com 
23/4/2004 
URL: http://www.zdnet.com.cn/developer/webdevelop/story/0,2000081602,39236051,00.htm 
将下面的代码添加到模块中

 

在内联网设置中，你通常对用户用来访问本地Web应用程序的浏览器有控制权。通常，这种浏览器是Internet Explorer（IE）。

这不一定是坏事，因为IE给开发者提供了大量创建强劲内联网应用程序的工具。本文并不是在推销IE，但是因为IE确实存在，而且是本地内联网应用程序的常用工具，所以有必要提到一些对Web开发人员有用的技术。

其中一种技术就是浏览器辅助对象（BHO）。BHO是在IE进程空间中运行的、能在可利用的窗口和模板中执行任何指令的组件对象模型（COM）对象。BHO通过常规脚本之外的Web应用程序可以给你提供额外功能。

假设你已经创建了需要用户输入的不同程序，而其中一些程序需要相同的信息。那这就要求用户输入数据许多次，而这可能会产生错误数据。

但是如果当载入页面时，你可以自动输入数据，并且你可以将这些信息储存在一个用户计算机上，创建一个中央信息库，那又会怎样呢？其中可以列入这个范畴的一条信息是用户的私人信息，如地址、电话号码等。当包含表格输入的页面载入时，BHO将自动将数据键入到表格字段中。

我将使用Visual Basic 来创建这个组件。但是，为了提供接口让IE能和组件交流，我必须参照一个显示IObjectWithSite接口的类型库。因为这不是十分容易，所以我不得不创建一个。我将使用对象描述语言（ODL）以及和VB一起装载的mktyplib工具来实现它。创建一个叫作VBBHO.ODL的文本文件，键入如下代码：

 [
uuid(CF9D9B76-EC4B-470D-99DC-AEC6F36A9261),
helpstring("VB IObjectWithSite Interface"),
version(1.0)
]
library IObjectWithSiteTLB
{
importlib("stdole2.tlb");
typedef [public] long GUIDPtr;
typedef [public] long VOIDPtr;
[
uuid(00000000-0000-0000-C000-000000000046),
odl
]
interface IUnknownVB
{
HRESULT QueryInterface(
[in] GUIDPtr priid,
[out] VOIDPtr *pvObj
);
long AddRef();
long Release();
}
[
uuid(FC4801A3-2BA9-11CF-A229-00AA003D7352),
odl
]
interface IObjectWithSite:IUnknown
{
typedef IObjectWithSite *LPOBJECTWITHSITE;
HRESULT SetSite([in] IUnknownVB* pSite);
HRESULT GetSite([in] GUIDPtr priid, [in, out] VOIDPtr* ppvObj);
}
}; 

 

保存这个文件，并用mktyplib工具创建类型库文件。打开命令提示符，定位到包含MKTYPLIB.EXE的目录地址，然后键入mktyplib c:\[path to ODL file]\vbbho.odl。在VB中创建一个新的ActiveX DLL应用程序，将这个工程命名为VBBHO，将类模块则命名为MyBHO。打开程序索引，点击浏览按钮，接着添加我们刚刚创建的VBBHO.TLB文件。还有，要参照微软XML2.6版（Microsoft XML v2.6）或更新版本、微软Internet 控件（Microsoft Internet Controls）以及微软HTML对象库（Microsoft HTML Object Library）。

 

将下面的代码添加到模块中


将下面的代码添加到你的类模块中：

Option Explicit
Option Base 0

Implements IObjectWithSiteTLB.IObjectWithSite
Dim WithEvents m_ie As InternetExplorer
Dim m_Site As IUnknownVB

Private Sub IObjectWithSite_GetSite(ByVal priid As
IObjectWithSiteTLB.GUIDPtr,
 ppvObj As IObjectWithSiteTLB.VOIDPtr)
    m_Site.QueryInterface priid, ppvObj
End Sub

Private Sub IObjectWithSite_SetSite(ByVal pSite As
IObjectWithSiteTLB.IUnknownVB)
    Set m_Site = pSite
    Set m_ie = pSite
End Sub

Private Sub m_ie_DocumentComplete(ByVal pDisp As Object,
URL As Variant)
On Error GoTo ErrorHandler
    Dim HTMLDoc As MSHTML.HTMLDocument
    Dim HTMLElement As MSHTML.HTMLInputElement
    Dim ElementCollection As Object
    Dim DOMDoc As MSXML2.DOMDocument
    Dim i As Integer, l As Integer
    Dim m_lError As Long, m_sError As String
    m_lError = 0
    Set HTMLDoc = m_ie.document
    Set ElementCollection = HTMLDoc.getElementsByName("myInput")
    l = ElementCollection.length
    If l > 0 Then
        Set DOMDoc = New MSXML2.DOMDocument
        DOMDoc.Load App.Path & "/data.xml"
        If DOMDoc.parseError.errorCode <> 0 Then
            App.LogEvent "DOM Error: " & DOMDoc.parseError.errorCode
 & vbCrLf
& DOMDoc.parseError.reason
            GoTo ExitCall
        End If
        Dim sField As String
        For i = 1 To l
            Set HTMLElement = ElementCollection.Item("myInput", i - 1)
            On Error Resume Next
            HTMLElement.setAttribute "value",
 DOMDoc.selectSingleNode(HTMLElement.getAttribute("field")).Text
            On Error GoTo ErrorHandler
        Next
    End If
ExitCall:
    Set HTMLDoc = Nothing
    Set HTMLElement = Nothing
    Set ElementCollection = Nothing
    Set DOMDoc = Nothing
    If m_lError <> 0 Then
        App.LogEvent "There was an error in VBBHO.MyBHO: " & vbCrLf &
 m_lError & vbCrLf & m_sError
    End If
    Exit Sub
ErrorHandler:
    m_lError = Err.Number
    m_sError = Err.Description
    Err.Clear
    GoTo ExitCall
End Sub 

 

当IE开始运行时，它创建了一个对象实例，并将首先调用SetSite()方法。一个指向InternetExplorer对象的指针被传送进去并同时保存在m_Site和m_ie成员变量当中。m_Site成员变量用来传回GetSite()方法中的一个参数值。

 

这组编码中最重要的部分是m_ie成员变量的DocumentComplete事件。当这个事件启动时（也就是当页面完成装载时），每个myInput INPUT元素进行循环，值属性被设置。输入标示符如下：

<INPUT TYPE="text" NAME="myInput" field="//personal_info/first_name"> 

  

这条代码还装载了一个叫作data.xml的文件，这个文件位于和组件相同的目录地址。下面就是那个文件的XML：

<?xml version='1.0'?>
<xml>
    <personal_info>
        <first_name>John</first_name>
        <last_name>Public</last_name>
        <age>99</age>
    </personal_info>
</xml> 

 

为了使IE启动完整的组件，你需要在HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects下面添加一个访问注册表的键，如果Browser Helper Objects键不存在，你需要添加它。编辑组件，并在HKEY_CLASSES_ROOT\VBBHO.MyBHO的注册表中找到组件的CLSID文件。复制组件的CLSID文件并把它添加到Browser Helper Objects键。你也可以在SetSite()方法中的组件代码中增加一个MsgBox调用，来确保它的装载。

 

一旦你确定它正在装载，下面的HTML就是用来测试你的组件：

<HTML>
<BODY>
<FORM>
<INPUT TYPE="text" NAME="myInput" field="//personal_info/first_name"><BR>
<INPUT TYPE="text" NAME="myInput" field="//personal_info/last_name"><BR>
<INPUT TYPE="text" NAME="myInput" field="//personal_info/age"><BR>
</FORM>
</BODY>
</HTML> 

 

登录到IE中的这个页面，你可以看到这些自动装载XML文件信息的INPUT字段。

 

注意：编辑注册表是不安全的。在做任何改动之前，务必要对注册表进行备份，这样如果有错误出现，你就可以对它进行恢复。 

责任编辑：李宁

欢迎评论或投稿 
 
 
