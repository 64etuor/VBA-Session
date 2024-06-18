Attribute VB_Name = "Module1"
'--------------------------------------
'● 이번 강의는 입문자 대상으로, 변수 설정이나 오류처리 없이 최대한 간단하게 준비했습니다.
' 오류처리를 모두 포함한 GoogleTranslate 함수 제작방법은 아래 프리미엄 클래스를 확인하세요!
' https://www.oppadu.com/엑셀-live-72강/
'---------------------------------------
 
'① GoogleTranslate 함수 작성하기
' Function GoogleTranslate(originaltext, sFrom, sTo)
' Function GoogleTranslate(OriginalText, Optional sFrom = "auto", Optional sTo = "ko")
'
'② 구글번역을 요청할 URL 만들기
' strURL = "https://translate.google.com/m?sl=출발언어&tl=도착언어&q=번역할문장"
'
'③ EncodeURL 함수로 유니코드언어 오류 처리하기
' OriginalText = EncodeURL(OriginalText)
'
'④ GetHTTP 함수로 URL 결과 받아오기
' strResult = GetHTTP(strURL)
'
'⑤ Splitter 함수로 <div class="result-container"> ~~~ <div> 사이 단어 반환하기
' strResult = splitter(strResult,"<div class=""result-container"">", "</div>")
'
'⑥ GoogleTraslate 함수 결과 반환 후 종료
' GoogleTranslate = strResult
'
'⑦ (선택) 결과값에 Error 413 코드 있을 경우, 최대 번역글자수 5000자 초과이므로, 오류 반환 후 종료
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
'오빠두엑셀 VBA 사용자지정함수 (https://www.oppadu.com)
'▶ GetHttp 함수
'▶ 웹에서 데이터를 받아옵니다.
'▶ 인수 설명
'_____________URL : 데이터를 스크랩할 웹 페이지 주소입니다.
'_____________formText : Encoding 된 FormText 형식으로 보내야 할 경우, Send String에 쿼리문을 추가합니다.
'_____________isWinHttp : WinHTTP 로 요청할지 여부입니다. Redirect가 필요할 경우 True로 입력하여 WinHttp 요청을 전송합니다.
'_____________RequestHeader : RequestHeader를 배열로 입력합니다. 반드시 짝수(한 쌍씩 이루어진) 개수로 입력되어야 합니다.
'_____________includeMeta : TRUE 일 경우 HTML 문서위로 ResponseText를 강제 입력합니다. Meta값이 포함되어 HTML이 작성되며 innerText를 사용할 수 없습니다. 기본값은 False 입니다.
'_____________RequestType : 요청방식입니다. 기본값은 "GET"입니다.
'_____________ReturnInnerHTML : TRUE 일 경우 InnerHTML을 기본으로 반환합니다. 기본값은 TRUE 입니다.
 
'▶ 사용 예제
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
 
' 2023-02-22 | 수정 | 윈도우 인증 접속 문제 발생 시 (일부 버전) ServerXMLHTTP -> XMLHTTP 요청으로 변경
' XMLHTTP 요청 시, TimeOut 세팅 불가 (기본값 설정)
' https://stackoverflow.com/questions/11605613/differences-between-xmlhttp-and-serverxmlhttp
On Error GoTo SendError:
'------------------------------
If isWinHttp = False Then
Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
Else
Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
End If
 
objHTTP.setTimeouts 1200000, 1200000, 1200000, 1200000 '응답 대기시간 120초
 
' 2023-02-22 | 수정 | 윈도우 인증 접속 문제 발생 시 (일부 버전) ServerXMLHTTP -> XMLHTTP 요청으로 변경
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
 
' 2023-02-22 | 수정 | 윈도우 인증 접속 문제 발생 시 (일부 버전) ServerXMLHTTP -> XMLHTTP 요청으로 변경
Exit Function
 
SendError:
Set objHTTP = CreateObject("MSXML2.XMLHTTP")
On Error GoTo 0
Resume SendRestart:
'------------------------------
 
End Function
 
Function ENCODEURL(varText As Variant, Optional blnEncode = True)
 
'###############################################################
'오빠두엑셀 VBA 사용자지정함수 (https://www.oppadu.com)
'▶ EncodeURL 함수
'▶ 한글/영문, 특수기호가 포함된 문자열을 웹 URL 표준 주소로 변환합니다.
'▶ 인수 설명
'_____________varTest : 표준 URL 주소로 변환할 문자열입니다.
'_____________blnEncode : TRUE 일 경우 결과값을 출력합니다.
'▶ 사용 예제
's = "http://www.google.com/search=사과"
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
'오빠두엑셀 VBA 사용자지정함수 (https://www.oppadu.com)
'▶ Splitter 함수
'▶ Cutter ~ Timmer 사이의 문자를 추출합니다. (Timmer가 빈칸일 경우 Cutter 이후 문자열을 추출합니다.)
'▶ 인수 설명
'_____________v : 문자열입니다.
'_________Cutter : 문자열 절삭을 시작할 텍스트입니다.
'_________Trimmer : 문자열 절삭을 종료할 텍스트입니다. (선택인수)
'▶ 사용 예제
'Dim s As String
's = "{sa;b132@drama#weekend;aabbcc"
's = Splitter(s, "@", "#")
'msgbox s '--> "drama"를 반환합니다.
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

