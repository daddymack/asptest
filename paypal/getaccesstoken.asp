<%@ Language="VBScript"%>
<%option explicit%>
<!--#include file="utils.asp"-->
<script language="JScript" runat="server" src='../_private/js/json2.js'></script>
<%
    Const SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS = 13056
    Const SXH_OPTION_SELECT_CLIENT_SSL_CERT = 3        
    dim client_id 
    client_id = "AXdp3xCcrIVM1AfsMrSbzL9sj3b6v7FzwMSznxWJyYXeYJooy86TKDS4Nu2E"
    dim secret
    secret = "EFqeLxBe6u1fKG4Xcmj8S15owbtSzQuWGUdoWZYw7CJ1MilRXSpsKQvJQ-Go"
            
    dim strURL
                
    strURL = "https://api.sandbox.paypal.com/v1/oauth2/token"

    Dim HttpReq
    
    set HttpReq = CreateObject("MSXML2.ServerXMLHTTP") 

    dim data
    data = "{""grant_type"":""client_credentials""}"

    HttpReq.setOption(2) = SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS
    HttpReq.setOption(3) = "LOCAL_MACHINE\My\futuregreatdoctor_api1.gmail.com"

    HttpReq.Open "POST", strURL, false
    'HttpReq.setRequestHeader "Content-type","application/x-www-form-urlencoded"
    HttpReq.setRequestHeader "Content-type","application/json"
    HttpReq.setRequestHeader "Accept", "application/json"
    HttpReq.setRequestHeader "Accept-Language", "en_US"
    'HttpReq.setRequestHeader "Authorization", "Basic " & Base64Encode(client_id & ":" & secret)
    HttpReq.setRequestHeader "X-PAYPAL-SECURITY-USERID", "futuregreatdoctor_api1.gmail.com"
    HttpReq.setRequestHeader "X-PAYPAL-SECURITY-PASSWORD", "WUSKMK9ZBQLB8KNY"
'vXMLHttp.setRequestHeader "X-PAYPAL-REQUEST-DATA-FORMAT", paypal_request_format
'vXMLHttp.setRequestHeader "X-PAYPAL-RESPONSE-DATA-FORMAT", paypal_response_format
    HttpReq.setRequestHeader "X-PAYPAL-APPLICATION-ID", "APP-79J7049294519430U"
    'HttpReq.setRequestHeader "X-PAYPAL-SECURITY-SIGNATURE", "AFcWxV21C7fd0v3bYYYRCpSSRl31A8EgR2HWfamuKj.DQwDsY1Yfjyiu"
    HttpReq.send data
    
    dim resp, result
    resp = HttpReq.responseText

	        set HttpReq=Nothing
	        if resp <> "" Then
	            set result=JSON.parse(resp)
            end if
 %>
