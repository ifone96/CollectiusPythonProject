--AEON on Local dwh_th_2022
select a.[alternis_portfolioidname] as Portfolio,
    a.[alternis_batchidname] as Batch,
    a.alternis_number as "Account Number",
    cast(a.alternis_invoicenumber as text) as "Invoice Number",
    a.[alternis_accountid] as uuid,
    phone.alternis_number as "Phone Number",
    phone.alternis_phonetypename as "Phone Type",
    a.[alternis_contactidname] as "Debtor Name",
    a.alternis_idnumber as "ID Card",
    --phone.alternis_sourcename as "Source",
    phone.alternis_verificationstatusname as "Verification Status",
    a.[alternis_processstagename] as "Process Stage",
    --a.owneridname as "Mediator",
    a.alternis_outstandingprincipal as "Outstanding Principal",
    a.alternis_lastpaymentdate as "Last Payment Date",
    a.alternis_outstandingbalance as "Outstanding Balance",
    --c.contactid as "contactid",
    --task.subject as "Task Subject",
    phone_call.subject as "PhoneCall Subject",
    phone_call.alternis_contactdispositionname as "Contact Disposition",
    phone_call.alternis_calloutcomename as "Calloutcome",
    phone_call.createdon as "Last Phonecall Createdon",
    datediff(day,floor(cast(phone_call.createdon as float)),CAST( GETDATE() AS Date )) as "Last Touch Day",
    case when datediff(day,floor(cast(phone_call.createdon as float)),CAST( GETDATE() AS Date )) >40000 then 'No Activity'
when datediff(day,floor(cast(phone_call.createdon as float)),CAST( GETDATE() AS Date )) >120 then '07. More than 4 Months'
when datediff(day,floor(cast(phone_call.createdon as float)),CAST( GETDATE() AS Date )) >90 then '06. 3 Monhts to 4 Months'
when datediff(day,floor(cast(phone_call.createdon as float)),CAST( GETDATE() AS Date )) >60 then '05. 2 Months to 3 Months'
when datediff(day,floor(cast(phone_call.createdon as float)),CAST( GETDATE() AS Date )) >30 then '04. 1 Month to 2 Months'
when datediff(day,floor(cast(phone_call.createdon as float)),CAST( GETDATE() AS Date )) >21 then '03. 3 Weeks to 1 Months'
when datediff(day,floor(cast(phone_call.createdon as float)),CAST( GETDATE() AS Date )) >14 then '02. 2 Weeks to 3 Weeks'
when datediff(day,floor(cast(phone_call.createdon as float)),CAST( GETDATE() AS Date )) >7 then '01. 1 Week to 2 Weeks'
when datediff(day,floor(cast(phone_call.createdon as float)),CAST( GETDATE() AS Date )) >=0 then '00. Less Than 1 Week'
else 'No Activity'
end as Last_Touch
from stage.alternis_account a
    join stage.contact c on c.contactid = a.alternis_contactid
    left join stage.alternis_phone phone on phone.alternis_contactid = a.alternis_contactid
    left join stage.phonecall phone_call on phone_call.phonenumber = phone.alternis_number and phone_call.regardingobjectid = a.alternis_accountid and phone_call.activityid = (SELECT TOP(1)
            activityid
        FROM [stage].[phonecall] phoneCall
        where phoneCall.phonenumber = phone.alternis_number and phoneCall.regardingobjectid = a.alternis_accountid
        ORDER BY phoneCall.createdon DESC)
    left join stage.task on task.regardingobjectid = a.alternis_accountid and task.activityid = (select top(1)
            activityid
        from stage.task tas
        where tas.regardingobjectid = a.alternis_accountid
        ORDER BY tas.createdon DESC)
where a.alternis_portfolioidname IN ('AEON1 TH','AEON2 TH','AEON3 TH')
order by a.alternis_portfolioidname, phone_call.createdon desc