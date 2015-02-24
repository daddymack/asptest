<% option explicit %>
<!--#include file="FedexAccountInfo.asp"-->
<%
    Dim subscriberzip
    Dim subscribercountry
    Dim ShipmentDate
    Dim xmlReq
    Dim objhttp
    Dim outstr
    Dim NodeList
 
    subscriberzip = "92840" 'Request.QueryString("zip")
    subscribercountry = "US"
    ShipmentDate = "8/30/2014"'Request.QueryString("date")
 
    xmlReq = "<?xml version=""1.0"" encoding=""UTF-8""?>" &_
    "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">" &_
    "<soapenv:Body>" &_
    "<RateRequest xmlns=""http://fedex.com/ws/rate/v9"">" &_
    "<WebAuthenticationDetail>" &_
    "<UserCredential>" &_
    "<Key>" & FedExkey & "</Key>" &_
    "<Password>" & FedExPassword & "</Password>" &_
    "</UserCredential>" &_
    "</WebAuthenticationDetail>" &_
    "<ClientDetail>" &_
    "<AccountNumber>" & FedExAccountNumber & "</AccountNumber>" &_
    "<MeterNumber>" & FedExMeterNumber & "</MeterNumber>" &_
    "</ClientDetail>" &_
    "<TransactionDetail>" &_
    "<CustomerTransactionId>TEST</CustomerTransactionId>" &_
    "</TransactionDetail>" &_
    "<Version>" &_
    "<ServiceId>crs</ServiceId>" &_
    "<Major>9</Major>" &_
    "<Intermediate>0</Intermediate>" &_
    "<Minor>0</Minor>" &_
    "</Version>" &_
    "<ReturnTransitAndCommit>1</ReturnTransitAndCommit>" &_
    "<CarrierCodes>FDXE</CarrierCodes>" &_
    "<VariableOptions>SATURDAY_DELIVERY</VariableOptions>" &_
    "<RequestedShipment>" &_
    "<ShipTimestamp>" & ShipmentDate & "T09:00:00-00:00</ShipTimestamp>" &_
    "<DropoffType>REGULAR_PICKUP</DropoffType>" &_
    "<PackagingType>YOUR_PACKAGING</PackagingType>" &_
    "<Shipper>" &_
    "<Address>" &_
    "<PostalCode>96790</PostalCode>" &_
    "<CountryCode>US</CountryCode>" &_
    "</Address>" &_
    "</Shipper>" &_
    "<Recipient>" &_
    "<Address>" &_
    "<PostalCode>" & subscriberzip & "</PostalCode>" &_
    "<CountryCode>US</CountryCode>" &_
    "</Address>" &_
    "</Recipient>" &_
    "<ShippingChargesPayment>" &_
    "<PaymentType>SENDER</PaymentType>" &_
    "<Payor>" &_
    "<AccountNumber>" & FedExAccountNumber & "</AccountNumber>" &_
    "<CountryCode>US</CountryCode>" &_
    "</Payor>" &_
    "</ShippingChargesPayment>" &_
    "<RateRequestTypes>LIST</RateRequestTypes>" &_
    "<PackageCount>1</PackageCount>" &_
    "<PackageDetail>INDIVIDUAL_PACKAGES</PackageDetail>" &_
    "<RequestedPackageLineItems>" &_
    "<SequenceNumber>1</SequenceNumber>" &_
    "<Weight>" &_
    "<Units>LB</Units>" &_
    "<Value>10.0</Value>" &_
    "</Weight>" &_
    "</RequestedPackageLineItems>" &_
    "</RequestedShipment>" &_
    "</RateRequest>" &_
    "</soapenv:Body>" &_
    "</soapenv:Envelope>"
 
    set objHttp = Server.createobject("Msxml2.ServerXMLHTTP")
 
    'For live
    objHttp.open "POST", "https://wsbeta.fedex.com:443/web-services/rate", false
 
    'For test
    'objHttp.open "POST", https://gatewaybeta.fedex.com:443/web-services/rate, false
 
    OBJHTTP.setRequestHeader "Referer", "Your Company name"
    OBJHTTP.setRequestHeader "Host", "wsbeta.fedex.com"
    OBJHTTP.setRequestHeader "Accept", "image/gif, image/jpeg,image/pjpeg, text/plain, text/html, */*"
    OBJHTTP.setRequestHeader "Content-Type", "image/gif"
    OBJHTTP.setRequestHeader "Content-Length", cstr(len(xmlReq))
 
    objHttp.Send xmlReq
 
    outstr = objHttp.responseText
 
    dim objDoc, i, j, status
    Set objDoc = CreateObject("Microsoft.XMLDOM")
    objDoc.async = False
    objDoc.LoadXml(outstr)
 
'    Response.Write objDoc.getElementsByTagName("v9:HighestSeverity")(0).text
 
    Set NodeList = objDoc.getElementsByTagName("v9:RateReplyDetails")
 
    for i=0 to (NodeList.length-1)
'        for j=0 to (NodeList(0).childNodes.length-1)
'            if Trim(NodeList(i).childNodes(j).nodename) = "v9:ServiceType" then
'                Response.write "ServiceType: " & Trim(NodeList(i).childNodes(j).text) & "<br/>"
'            elseif Trim(NodeList(i).childNodes(j).nodename) = "v9:DeliveryDayOfWeek" then
'                Response.write "Day: " & Trim(NodeList(i).childNodes(j).text) & "<br/>"
'            elseif Trim(NodeList(i).childNodes(j).nodename) = "v9:DeliveryTimestamp" then
'                Response.write "Date: " & left(Trim(NodeList(i).childNodes(j).text), 10) & "<br/>"
'                Response.write "Time: " & right(Trim(NodeList(i).childNodes(j).text), 8) & "<br/>"
'            elseif Trim(NodeList(i).childNodes(j).nodename) = "v9:DestinationServiceArea" then
'                Response.write "Area: " & Trim(NodeList(i).childNodes(j).text) & "<br/><br/>"
'            end if
'        next
    next
 
    if Len(outstr) = 0 then
        Response.write "<br/> Error: Unable to communicate with Fedex Server. Please check your Internet connection.<br/>"
    else
        Response.Write outstr
    end if
%>