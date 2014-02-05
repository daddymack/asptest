<%@ Language="VBScript"%>
<%option explicit%>
<!--#include file="../_private/standard_funcs.asp"-->
<script language="JScript" runat="server" src='../_private/js/json2.js'></script>
<%
    Const SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS = 13056
    Const SXH_OPTION_SELECT_CLIENT_SSL_CERT = 3        
    
    dim strURL
    strURL = "https://api.sandbox.paypal.com/v1/payments/payment"

    Dim HttpReq
    set HttpReq = CreateObject("MSXML2.ServerXMLHTTP") 
    'set HttpReq = CreateObject("Microsoft.XMLHTTP") 
    

    dim data
    data = "{""intent"":""sale""," & _
            """payer"":{" & _
    """payment_method"":""credit_card""," & _
    """funding_instruments"":[" & _
      "{" & _
        """credit_card"":{" & _
          """number"":""4417119669820331""," & _
          """type"":""visa""," & _
          """expire_month"":11," & _
          """expire_year"":2018," & _
          """cvv2"":""874""," & _
          """first_name"":""Betsy""," & _
          """last_name"":""Buyer""," & _
          """billing_address"":{" & _
            """line1"":""111 First Street""," & _
            """city"":""Saratoga""," & _
            """state"":""CA""," & _
            """postal_code"":""95070""," & _
            """country_code"":""US""" & _
          "}" & _
        "}" & _
      "}" & _
    "]" & _
  "}," & _
  """transactions"":[" & _
    "{" & _
      """amount"":{" & _
        """total"":""7.47""," & _
        """currency"":""USD""," & _
        """details"":{" & _
          """subtotal"":""7.41""," & _
          """tax"":""0.03""," & _
          """shipping"":""0.03""" & _
        "}" & _
      "}," & _
      """description"":""This is the payment transaction description.""" & _
    "}" & _
  "]" & _
"}"

    HttpReq.setOption(2) = SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS
    HttpReq.setOption(3) = "LOCAL_MACHINE\My\futuregreatdoctor-facilitator_api1.gmail.com"

    HttpReq.Open "POST", strURL, false
    HttpReq.setRequestHeader "Content-type","application/json"
    HttpReq.setRequestHeader "Accept", "application/json"
    HttpReq.setRequestHeader "Accept-Language", "en_US"
    HttpReq.setRequestHeader "Authorization", Session("paypal_token")
    
    HttpReq.send data
    
    dim resp, result
    resp = HttpReq.responseText
    set HttpReq=Nothing
	response.Write resp
 %>
