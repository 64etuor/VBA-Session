Attribute VB_Name = "Module1"
'--------------------------------------
'�� �̹� ���Ǵ� �Թ��� �������, ���� �����̳� ����ó�� ���� �ִ��� �����ϰ� �غ��߽��ϴ�.
' ����ó���� ��� ������ GoogleTranslate �Լ� ���۹���� �Ʒ� �����̾� Ŭ������ Ȯ���ϼ���!
' https://www.oppadu.com/����-live-72��/
'---------------------------------------
 
'�� GoogleTranslate �Լ� �ۼ��ϱ�
' Function GoogleTranslate(originaltext, sFrom, sTo)
' Function GoogleTranslate(OriginalText, Optional sFrom = "auto", Optional sTo = "ko")
'
'�� ���۹����� ��û�� URL �����
' strURL = "https://translate.google.com/m?sl=��߾��&tl=�������&q=�����ҹ���"
'
'�� EncodeURL �Լ��� �����ڵ��� ���� ó���ϱ�
' OriginalText = EncodeURL(OriginalText)
'
'�� GetHTTP �Լ��� URL ��� �޾ƿ���
' strResult = GetHTTP(strURL)
'
'�� Splitter �Լ��� <div class="result-container"> ~~~ <div> ���� �ܾ� ��ȯ�ϱ�
' strResult = splitter(strResult,"<div class=""result-container"">", "</div>")
'
'�� GoogleTraslate �Լ� ��� ��ȯ �� ����
' GoogleTranslate = strResult
'
'�� (����) ������� Error 413 �ڵ� ���� ���, �ִ� �������ڼ� 5000�� �ʰ��̹Ƿ�, ���� ��ȯ �� ����
' If InStr(1, strResult, "Error 413") > 0 Then GoogleTranslate = "#Request Too Large!": Exit Function

Function GoogleTranslate(OriginalText)
 
sFrom = "ko"
sTo = "en"

OriginalText = ENCODEURL(OriginalText)
strURL = "https://translate.google.com/m?sl=" & sFrom & "&tl=" & sTo & "&q=" & OriginalText
oResult = GetHttp(strURL)
 
sResult = Splitter(oResult, "<div class=""result-container"">", "</div>")
If Len(sResult) = 0 Then
sResult = Splitter(oResult, "result-container>", "</DIV>")
End If
 
GoogleTranslate = sResult
 
End Function
 
 
Function GetHttp(URL, Optional formText As String, _
Optional isWinHttp As Boolean = False, _
Optional RequestHeader As Variant, _
Optional includeMeta As Boolean = False, _
Optional RequestType As String = "GET", _
Optional returnInnerHTML As Boolean = True)
 
'###############################################################
'�����ο��� VBA ����������Լ� (https://www.oppadu.com)
'�� GetHttp �Լ�
'�� ������ �����͸� �޾ƿɴϴ�.
'�� �μ� ����
'_____________URL : �����͸� ��ũ���� �� ������ �ּ��Դϴ�.
'_____________formText : Encoding �� FormText �������� ������ �� ���, Send String�� �������� �߰��մϴ�.
'_____________isWinHttp : WinHTTP �� ��û���� �����Դϴ�. Redirect�� �ʿ��� ��� True�� �Է��Ͽ� WinHttp ��û�� �����մϴ�.
'_____________RequestHeader : RequestHeader�� �迭�� �Է��մϴ�. �ݵ�� ¦��(�� �־� �̷����) ������ �ԷµǾ�� �մϴ�.
'_____________includeMeta : TRUE �� ��� HTML �������� ResponseText�� ���� �Է��մϴ�. Meta���� ���ԵǾ� HTML�� �ۼ��Ǹ� innerText�� ����� �� �����ϴ�. �⺻���� False �Դϴ�.
'_____________RequestType : ��û����Դϴ�. �⺻���� "GET"�Դϴ�.
'_____________ReturnInnerHTML : TRUE �� ��� InnerHTML�� �⺻���� ��ȯ�մϴ�. �⺻���� TRUE �Դϴ�.
 
'�� ��� ����
'Dim HtmlResult As Object
'Set htmlResult = GetHttp("https://www.naver.com")
'msgbox htmlResult.body.innerHTML
'###############################################################
 




Dim oHTMLDoc As Object: Dim objHTTP As Object
Dim HTMLDoc As Object
Dim i As Long: Dim blnAgent As Boolean: blnAgent = False
Dim sUserAgent As String: sUserAgent = "Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.183 Mobile Safari/537.36"
 
Application.DisplayAlerts = False
 
If Left(URL, 4) <> "http" Then URL = "http://" & URL
 
Set oHTMLDoc = CreateObject("HtmlFile")
Set HTMLDoc = CreateObject("HtmlFile")
 
' 2023-02-22 | ���� | ������ ���� ���� ���� �߻� �� (�Ϻ� ����) ServerXMLHTTP -> XMLHTTP ��û���� ����
' XMLHTTP ��û ��, TimeOut ���� �Ұ� (�⺻�� ����)
' https://stackoverflow.com/questions/11605613/differences-between-xmlhttp-and-serverxmlhttp
On Error GoTo SendError:
'------------------------------
If isWinHttp = False Then
Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
Else
Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
End If
 
objHTTP.setTimeouts 1200000, 1200000, 1200000, 1200000 '���� ���ð� 120��
 
' 2023-02-22 | ���� | ������ ���� ���� ���� �߻� �� (�Ϻ� ����) ServerXMLHTTP -> XMLHTTP ��û���� ����
SendRestart:
'------------------------------
objHTTP.Open RequestType, URL, False
If Not IsMissing(RequestHeader) Then
Dim vRequestHeader As Variant
For Each vRequestHeader In RequestHeader
Dim uHeader As Long: Dim Lheader As Long: Dim steps As Long
uHeader = UBound(vRequestHeader): Lheader = LBound(vRequestHeader)
If (uHeader - Lheader) Mod 2 = 0 Then GetHttp = CVErr(xlValue): Exit Function
For i = Lheader To uHeader Step 2
If vRequestHeader(i) = "User-Agent" Then blnAgent = True
objHTTP.setRequestHeader vRequestHeader(i), vRequestHeader(i + 1)
Next
Next
End If
If blnAgent = False Then objHTTP.setRequestHeader "User-Agent", sUserAgent
 
objHTTP.send formText
 
If includeMeta = False Then
With oHTMLDoc
.Open
.Write objHTTP.responseText
.Close
End With
Else
oHTMLDoc.body.innerhtml = objHTTP.responseText
End If
 
If returnInnerHTML = True Then
GetHttp = oHTMLDoc.body.innerhtml
Else
Set GetHttp = oHTMLDoc
End If
Set oHTMLDoc = Nothing
Set objHTTP = Nothing
 
Application.DisplayAlerts = True
 
' 2023-02-22 | ���� | ������ ���� ���� ���� �߻� �� (�Ϻ� ����) ServerXMLHTTP -> XMLHTTP ��û���� ����
Exit Function
 
SendError:
Set objHTTP = CreateObject("MSXML2.XMLHTTP")
On Error GoTo 0
Resume SendRestart:
'------------------------------
 
End Function
 
Function ENCODEURL(varText As Variant, Optional blnEncode = True)
 
'###############################################################
'�����ο��� VBA ����������Լ� (https://www.oppadu.com)
'�� EncodeURL �Լ�
'�� �ѱ�/����, Ư����ȣ�� ���Ե� ���ڿ��� �� URL ǥ�� �ּҷ� ��ȯ�մϴ�.
'�� �μ� ����
'_____________varTest : ǥ�� URL �ּҷ� ��ȯ�� ���ڿ��Դϴ�.
'_____________blnEncode : TRUE �� ��� ������� ����մϴ�.
'�� ��� ����
's = "http://www.google.com/search=���"
's = ENCODEURL(s)
'MsgBox s
'###############################################################
 
Static objHtmlfile As Object
 
If objHtmlfile Is Nothing Then
Set objHtmlfile = CreateObject("htmlfile")
With objHtmlfile.parentWindow
.execScript "function encode(s) {return encodeURIComponent(s)}", "jscript"
End With
End If
 
If blnEncode Then
ENCODEURL = objHtmlfile.parentWindow.encode(varText)
End If
 
End Function
 
Function Splitter(v As Variant, Cutter As String, Optional Trimmer As String)
 
'###############################################################
'�����ο��� VBA ����������Լ� (https://www.oppadu.com)
'�� Splitter �Լ�
'�� Cutter ~ Timmer ������ ���ڸ� �����մϴ�. (Timmer�� ��ĭ�� ��� Cutter ���� ���ڿ��� �����մϴ�.)
'�� �μ� ����
'_____________v : ���ڿ��Դϴ�.
'_________Cutter : ���ڿ� ������ ������ �ؽ�Ʈ�Դϴ�.
'_________Trimmer : ���ڿ� ������ ������ �ؽ�Ʈ�Դϴ�. (�����μ�)
'�� ��� ����
'Dim s As String
's = "{sa;b132@drama#weekend;aabbcc"
's = Splitter(s, "@", "#")
'msgbox s '--> "drama"�� ��ȯ�մϴ�.
'###############################################################
 
Dim vaArr As Variant
 
On Error GoTo EH:
 
vaArr = Split(v, Cutter)(1)
If Not IsMissing(Trimmer) Then vaArr = Split(vaArr, Trimmer)(0)
 
Splitter = vaArr
 
Exit Function
 
EH:
Splitter = ""
 
End Function

