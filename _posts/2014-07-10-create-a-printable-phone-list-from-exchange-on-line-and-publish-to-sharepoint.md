---
layout: post
title: "Create a Printable Phone List from Exchange On line and Publish to Sharepoint"
description: "Yes there are a million ways phone list from active directory and publish to sharepoint, but what happens when you move it all the the cloude."
category: ["scripts"]
tags: ["powershell","office365","sharepoint"]
---
{% include JB/setup %}

So what happens when your organization makes the move to on-premise stuff to cloud stuff. Well in some cases there is a lot of "Where is xxxx?", "I don't like yyyy.", "Would you make yyyy look like xxxx?".

Yes this is what we have to deal with sometimes. Here is one of those cases a simple phone list that can be printed out. 

```powershell
$LiveCred = Get-Credential

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ 
-Credential $LiveCred -Authentication Basic -AllowRedirection

Import-PSSession $Session

$a = "<style>"
$a = $a + "BODY{background-color:white; font-size:10pt;}"
$a = $a + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
$a = $a + "TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:#cccccc}"
$a = $a + "TD{border-width: 1px;padding: 2px;border-style: solid;border-color: black;}"
$a = $a + "tr:nth-child(even) {background-color: #cccccc}" ;
$a = $a + "</style>"

$date = Get-Date -format D

Get-Recipient -ResultSize Unlimited  | Sort-Object LastName | where {$_. Phone -ne ""} |
select LastName ,FirstName, Phone,Department |
ConvertTo-HTML -head $a -body "<H2 style='color: #08205C'>Phone List</H2><h4>Updated:$date </h4>" |
Out-File \\<server>\<share>\PhoneList.htm 

``` 