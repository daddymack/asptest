<%@ Language="VBScript"%>
<%option explicit%>
<!--#include file="../utils.asp"-->
<%
            
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

    HttpReq.Open "POST", strURL, false
    'HttpReq.setRequestHeader "Content-type","application/x-www-form-urlencoded"
    HttpReq.setRequestHeader "Content-type","application/json"
    HttpReq.setRequestHeader "Accept", "application/json"
    HttpReq.setRequestHeader "Accept-Language", "en_US"
    HttpReq.setRequestHeader "Authorization", "Basic " & Base64Encode(client_id & ":" & secret)
    HttpReq.send data
    
 %>
