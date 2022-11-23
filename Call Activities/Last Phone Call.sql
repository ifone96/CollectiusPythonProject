--Last Call
SELECT

	a.alternis_portfolioidname as 'Portfolio',
	a.alternis_number as 'Account Number',
	a.alternis_invoicenumber as 'Invoice',
	a.alternis_contactidname as 'Name',
	phone.alternis_phonetypename as 'Phone Type',
	phone_call.phonenumber as 'Phone Number',
	phone_call.alternis_calloutcomename as 'Call Outcome',
	phone_call.alternis_contactdispositionname as 'Contact Disposition',
	phone_call.description as 'Description',
	phone_call.createdon as 'Last Phonecall Createdon',
	phone_call.actualdurationminutes as 'Duration',
	phone_call.subject as 'Subject',
	phone_call.modifiedbyname as 'Agent Call'

FROM Stage.alternis_account a
FULL JOIN Stage.alternis_phone phone on phone.alternis_contactid = a.alternis_contactid
FULL JOIN Stage.phonecall phone_call on phone_call.phonenumber = phone.alternis_number and phone_call.regardingobjectid = a.alternis_accountid
WHERE a.alternis_portfolioidname IN ('SEAMONEY SPL SVC TH','SEAMONEY BCL SVC TH') 
--Change The Date
AND phone_call.createdon >= '2022-10-31 00:00:00.000' 
AND phone_call.createdon <= '2022-11-01 00:00:00.000'
ORDER BY phone_call.createdon DESC