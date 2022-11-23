--PTP
SELECT

	a.alternis_portfolioidname as 'Portfolio'
	,a.alternis_contactidname as 'Name'
	,a.alternis_number as 'Account Number'
	,a.alternis_invoicenumber as 'Invoice Number'
	,p.alternis_firstpaymentdate as '1st Payment Date'
	,p.alternis_installmentamount as 'Installment Amount'
    ,p.alternis_amountoninstallments as 'Total Amount on Installment'			
    ,p.alternis_totaldiscountvalue	as 'Total Discount Value'
	,p.alternis_amountpaid 'Paid'
	,p.statuscode as 'Status Reason'
	,p.createdon as 'Created On'
	,p.alternis_paymentplanid 'PTP ID'

FROM Stage.alternis_account a 
INNER JOIN Stage.alternis_paymentplan p ON p.alternis_accountid = a.alternis_accountid
WHERE a.alternis_portfolioidname IN ('SEAMONEY SPL SVC TH','SEAMONEY BCL SVC TH')
--Change The Date
AND p.alternis_firstpaymentdate >= '2022-10-27 00:00:00.000'
ORDER BY a.alternis_portfolioidname, p.alternis_firstpaymentdate DESC