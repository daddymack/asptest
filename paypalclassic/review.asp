<!-- #include file ="paypalfunctions.asp" -->
<% 
On Error Resume Next

'----	Get the token value from the Request string
token = Request("token")

'-- This indicates that the directed flow to this page is originating from PayPal, if this value is not present then user is not
'-- coming from PayPal.	
if token <> "" then

	'------------------------------------
	' Calls the GetExpressCheckoutDetails API call
	'
	' The GetShippingDetails function is defined in PayPalFunctions.asp
	' included at the top of this file.
	'-------------------------------------------------
	
	set resArray = GetShippingDetails( token )
	
	ack = UCase(resArray("ACK"))

	If ack <> "SUCCESS" Then
		'Display a user friendly Error on the page using any of the following error information returned by PayPal
		ErrorCode = URLDecode( resArray("L_ERRORCODE0"))
		ErrorShortMsg = URLDecode( resArray("L_SHORTMESSAGE0"))
		ErrorLongMsg = URLDecode( resArray("L_LONGMESSAGE0"))
		ErrorSeverityCode = URLDecode( resArray("L_SEVERITYCODE0"))
		
	Else
		
		email 			= resArray("EMAIL") ' Email address of payer.
		payerId 		= resArray("PAYERID") ' Unique PayPal customer account identification number.
		payerStatus		= resArray("PAYERSTATUS") ' Status of payer. Character length and limitations: 10 single-byte alphabetic characters.
		salutation		= resArray("SALUTATION") ' Payer's salutation.
		firstName		= resArray("FIRSTNAME") ' Payer's first name.
		middleName		= resArray("MIDDLENAME") ' Payer's middle name.
		lastName		= resArray("LASTNAME") ' Payer's last name.
		suffix			= resArray("SUFFIX") ' Payer's suffix.
		cntryCode		= resArray("COUNTRYCODE") ' Payer's country of residence in the form of ISO standard 3166 two-character country codes.
		business		= resArray("BUSINESS") ' Payer's business name.
		shipToName		= resArray("PAYMENTREQUEST_0_SHIPTONAME") ' Person's name associated with this address.
		shipToStreet	= resArray("PAYMENTREQUEST_0_SHIPTOSTREET") ' First street address.
		shipToStreet2	= resArray("PAYMENTREQUEST_0_SHIPTOSTREET2") ' Second street address.
		shipToCity		= resArray("PAYMENTREQUEST_0_SHIPTOCITY") ' Name of city.
		shipToState		= resArray("PAYMENTREQUEST_0_SHIPTOSTATE") ' State or province
		shipToCntryCode	= resArray("PAYMENTREQUEST_0_SHIPTOCOUNTRYCODE") ' Country code. 
		shipToZip		= resArray("PAYMENTREQUEST_0_SHIPTOZIP") ' U.S. Zip code or other country-specific postal code.
		addressStatus 	= resArray("ADDRESSSTATUS") ' Status of street address on file with PayPal   
		invoiceNumber	= resArray("INVNUM") ' Your own invoice or tracking number, as set by you in the element of the same name in SetExpressCheckout request .
		phonNumber		= resArray("PHONENUM") ' Payer's contact telephone number. Note:  PayPal returns a contact telephone number only if your Merchant account profile settings require that the buyer enter one. 

		' This is just an example of what can be displayed on the Order Review page. You can use this code as sample to integrate into your Order Review page the information returned
		' by PayPal		
		
		strFormattedOutput = "<table><tr>"
		strFormattedOutput = strFormattedOutput + "<td> First Name </td><td>" + firstName + "</td></tr>"
		strFormattedOutput = strFormattedOutput + "<td> Last Name </td><td>" + lastName + "</td></tr>"
		strFormattedOutput = strFormattedOutput + "<td colspan='2'> Shipping Address</td></tr>"
		strFormattedOutput = strFormattedOutput + "<td> Name </td><td>" + shipToName + "</td></tr>"
		strFormattedOutput = strFormattedOutput + "<td> Street1 </td><td>" + shipToStreet + "</td></tr>"
		strFormattedOutput = strFormattedOutput + "<td> Street2 </td><td>" + shipToStreet2 + "</td></tr>"
		strFormattedOutput = strFormattedOutput + "<td> City </td><td>" + shipToCity + "</td></tr>"
		strFormattedOutput = strFormattedOutput + "<td> State </td><td>" + shipToState + "</td></tr>"
		strFormattedOutput = strFormattedOutput + "<td> Zip </td><td>" + shipToZip + "</td></tr>"
		
		Response.Write strFormatted
	
	End If
End If
%>