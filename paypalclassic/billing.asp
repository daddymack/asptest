﻿<!-- #include file ="paypalfunctions.asp" -->
<% 
if PaymentOption = "PayPal" then

	' ==================================
	' PayPal Express Checkout Module
	' ==================================
	On Error Resume Next

	'------------------------------------
	' The currencyCodeType and paymentType 
	' are set to the selections made on the Integration Assistant 
	'------------------------------------
	currencyCodeType = "USD"
	paymentType = "Sale"

	'------------------------------------
	' The returnURL is the location where buyers return to when a
	' payment has been succesfully authorized.
	'
	' This is set to the value entered on the Integration Assistant 
	'------------------------------------
	returnURL = "http://localhost/order/paypalclassic/confirm.asp"

	'------------------------------------
	' The cancelURL is the location buyers are sent to when they click the
	' return to XXXX site where XXX is the merhcant store name
	' during payment review on PayPal
	'
	' This is set to the value entered on the Integration Assistant 
	'------------------------------------
	cancelURL = "http://localhost/order/paypalclassic/default.asp"

	'------------------------------------
	' The paymentAmount is the total value of 
	' the shopping cart, that was set 
	' earlier in a session variable 
	' by the shopping cart page
	'------------------------------------
	paymentAmount = Session("Payment_Amount")

	'------------------------------------
	' When you integrate this code 
	' set the variables below with 
	' shipping address details 
	' entered by the user on the 
	' Shipping page.
	'------------------------------------
	shipToName = "<<ShiptoName>>"
	shipToStreet = "<<ShipToStreet>>"
	shipToStreet2 = "<<ShipToStreet2>>" 'Leave it blank if there is no value
	shipToCity = "<<ShipToCity>>"
	shipToState = "<<ShipToState>>"
	shipToCountryCode = "<<ShipToCountryCode>>" ' Please refer to the PayPal country codes in the API documentation
	shipToZip = "<<ShipToZip>>"
	phoneNum = "<<PhoneNumber>>"

	'------------------------------------
	' Calls the SetExpressCheckout API call
	'
	' The CallMarkExpressCheckout function is defined in PayPalFunctions.asp
	' included at the top of this file.
	'-------------------------------------------------
	Set resArray = CallMarkExpressCheckout (paymentAmount, currencyCodeType, paymentType, returnURL, cancelURL, shipToName,
						shipToStreet, shipToCity, shipToState, shipToCountryCode, shipToZip, shipToStreet2, phoneNum )

	ack = UCase(resArray("ACK"))

	If ack="SUCCESS" Then
		' Redirect to paypal.com
		SESSION("token") = resArray("TOKEN")
		ReDirectURL( resArray("TOKEN") )
	Else  
		'Display a user friendly Error on the page using any of the following error information returned by PayPal
		ErrorCode = URLDecode( resArray("L_ERRORCODE0"))
		ErrorShortMsg = URLDecode( resArray("L_SHORTMESSAGE0"))
		ErrorLongMsg = URLDecode( resArray("L_LONGMESSAGE0"))
		ErrorSeverityCode = URLDecode( resArray("L_SEVERITYCODE0"))
	End If

Else If (((PaymentOption = "Visa") Or (PaymentOption = "MasterCard") Or (PaymentOption = "Amex") or (PaymentOption = "Discover"))
			and ( PaymentProcessorSelected = "PayPal Direct Payment"))

	'------------------------------------
	' The paymentAmount is the total value of 
	' the shopping cart, that was set 
	' earlier in a session variable 
	' by the shopping cart page
	'------------------------------------
	paymentAmount = Session("Payment_Amount")

	'------------------------------------
	' The paymentType that was selected earlier 
	'------------------------------------
	paymentType = "Sale"
	
	' Set these values based on what was selected by the user on the Billing page Html form
	
	creditCardType 			= "<<Visa/MasterCard/Amex/Discover>>" ' Set this to one of the acceptable values (Visa/MasterCard/Amex/Discover) match it to what was selected on your Billing page
	creditCardNumber 		= "<<CC number>>" ' Set this to the string entered as the credit card number on the Billing page
	expDate 				= "<<Expiry Date>>" ' Set this to the credit card expiry date entered on the Billing page
	cvv2 					= "<<cvv2>>" ' Set this to the CVV2 string entered on the Billing page 
	firstName 				= "<<firstName>>" ' Set this to the customer's first name that was entered on the Billing page 
	lastName 				= "<<lastName>>" ' Set this to the customer's last name that was entered on the Billing page 
	street 					= "<<street>>" ' Set this to the customer's street address that was entered on the Billing page 
	city 					= "<<city>>" ' Set this to the customer's city that was entered on the Billing page 
	state 					= "<<state>>" ' Set this to the customer's state that was entered on the Billing page 
	zip 					= "<<zip>>" ' Set this to the zip code of the customer's address that was entered on the Billing page 
	countryCode 			= "<<PayPal Country Code>>" ' Set this to the PayPal code for the Country of the customer's address that was entered on the Billing page 
	currencyCode 			= "<<PayPal Currency Code>>" ' Set this to the PayPal code for the Currency used by the customer 
				
	'------------------------------------------------
	' Calls the DoDirectPayment API call
	'
	' The DirectPayment function is defined in PayPalFunctions.asp included at the top of this file.
	'-------------------------------------------------
	Set resArray = DirectPayment (paymentType, paymentAmount, creditCardType, creditCardNumber,
							expDate, cvv2, firstName, lastName, street, city, state, zip, 
							countryCode, currencyCode ) 

	ack = UCase(resArray("ACK"))

	If ack <> "SUCCESS" Then
		'Display a user friendly Error on the page using any of the following error information returned by PayPal
		ErrorCode = URLDecode( resArray("L_ERRORCODE0"))
		ErrorShortMsg = URLDecode( resArray("L_SHORTMESSAGE0"))
		ErrorLongMsg = URLDecode( resArray("L_LONGMESSAGE0"))
		ErrorSeverityCode = URLDecode( resArray("L_SEVERITYCODE0"))

		' DirectPayment related error information	
		ErrorAVSCode = URLDecode( resArray("AVSCODE"))
		ErrorCVV2Match = URLDecode( resArray("CVV2MATCH"))
		ErrorTransactionId = URLDecode( resArray("TRANSACTIONID"))
		
	End If

End If
%>