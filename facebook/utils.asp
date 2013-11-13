<%
function curPageURL()
 dim s, protocol, port

 if Request.ServerVariables("HTTPS") = "on" then 
   s = "s"
 else 
   s = ""
 end if  
 
 protocol = strleft(LCase(Request.ServerVariables("SERVER_PROTOCOL")), "/") & s 

 if Request.ServerVariables("SERVER_PORT") = "80" then
   port = ""
 else
   port = ":" & Request.ServerVariables("SERVER_PORT")
 end if  

 curPageURL = protocol & "://" & Request.ServerVariables("SERVER_NAME") &_ 
              port & Request.ServerVariables("SCRIPT_NAME")
end function

function strLeft(str1,str2)
 strLeft = Left(str1,InStr(str1,str2)-1)
end function

function GetTextFromUrl(url)
  Dim oXMLHTTP
 
  Set oXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")

  oXMLHTTP.Open "GET", url, False
  oXMLHTTP.Send

  If oXMLHTTP.Status = 200 Then

    GetTextFromUrl = oXMLHTTP.responseText

  End If
end function
 %>