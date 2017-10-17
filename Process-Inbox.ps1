
[CmdletBinding()]
Param()

# This script is full of terrible examples. Please do not use this as an example
# of good powershell coding practices. Off the top of my head, it includes
# these poor practices:
# * Global variables everywhere
# * Hard-coded options which should be parameters to the script
# * No error checking. Good luck!
# * You have to make a reg entry or two to allow this to work. These reg entries
#     DISABLE some Outlook security protections. I think these protections are
#     no big deal given the current threat landscape, but pronouncements of this
#     sort tend to look very silly in hindsight.

# Required reg entries to make Outlook VULNERABLE to this script. If you don't
# make these entries, Outlook will (depending on security configuration) either
# silently deny access to this script, or repeatedly prompt you to allow access.

#Example for Outlook 2010 32bit on 64 bit Windows:

#Windows Registry Editor Version 5.00
#
#[HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Office\14.0\Outlook\Security]
#"ObjectModelGuard"=dword:00000002

#If you run a different version of outlook, change the 14.0 as appropriate.
#If you run 64 bit Office on 64 bit windows (or 32 on 32...upgrade!) then ditch
# the Wow6432Node folder.

$homecalendar = "" #If you really want to leak meeting info to your home cal.
$smtp = "" #hostname, required for leaking meeting info to your home calendar
$homemeetings_include_subject = $false # leak or generic?


Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null 

#$outlook = new-object -comobject outlook.application 
$outlook = New-Object -TypeName Microsoft.Office.Interop.Outlook.ApplicationClass
$namespace = $outlook.Session
$olFolders = "Microsoft.Office.Interop.Outlook.olDefaultFolders" -as [type]  
$olCloseOpts = "Microsoft.Office.Interop.Outlook.OlInspectorClose" -as [type] 
$olTableContents = "Microsoft.Office.Interop.Outlook.OlTableContents" -as [type]
$olItemTypes = "Microsoft.Office.Interop.Outlook.OlItemType" -as [type]
$olMeetingStatus = "Microsoft.Office.Interop.Outlook.OlMeetingStatus" -as [type]
$olMeetingResponse = "Microsoft.Office.Interop.Outlook.olMeetingResponse" -as [type] 

#Here's how to get a default folder.
$inbox = $namespace.getDefaultFolder($olFolders::olFolderInBox) 


#Here's how to get a subfolder by name:
#$dlpfolder = $inbox.Folders['DLP Folder']
#$alerts = $inbox.Folders['Alerts']


#Used a couple places later to send you email or send email from you.
$myuser = $outlook.Application.Session.CurrentUser.AddressEntry.GetExchangeUser()
$myaddy = $myuser.PrimarySmtpAddress
$myfullname = $myuser.Name

#note that all of these cut-offs use generous rounding-up. its my fault.
$cutoff = get-date
$cutoff1d = $cutoff.AddDays(-1)
$cutoff7d = $cutoff.AddDays(-7)
# a month has 31 days, despite what our round numbers think
$cutoff30d = $cutoff.AddDays(-31)
# yeah I know this is 92. 90 days is commonly-used for 3 months ago. but 3
# months can be a max of 92 days (go look at a calendar)
$cutoff90d = $cutoff.AddDays(-30-31-31) 
$cutoff1y = $cutoff.Add(-366)
# max of 2 leap years in a 7 year span
$cutoff7y = $cutoff.AddDays(-(365 * 5 + 366 * 2)) 

$expiry = $cutoff.AddDays(7)

#this global variable (Hi, anyone involved in teaching me Com Sci. I'm sorry)
#will contain the body of the Process-Inbox Results email that will be emailed
#to yourself at the end of the script. It is updated via the Add-Mail function.
$email_txt = ""


#used internally to add text to the message body to be sent to yourself with results
#when you modify this script, use Add-Mail instead of anything that might write
#to the host, and then set $VerbosePreference = "Continue". This will ensure that
#the text is sent to you at the end of the script via email, and while you are
#testing the script, $VerbosePreference = "Continue" will ensure the text is
#also printed to the terminal.
Function Add-Mail {
    [CmdletBinding()]
    param(
        [Parameter(Position=1,mandatory=$True)][string]$txt
        )
    $script:email_txt+="$txt`n"
    Write-Verbose $txt
}

#Used internally to build filter strings
Function Safe-Append {
    [CmdletBinding()]
    param(
        [Parameter(Position=1)]$dest,
        [Parameter(Position=2)][string[]]$toappend,
        [Parameter(Position=3)][string]$jointxt = " And "
        )
    return ( ( ($dest -split $jointxt) -ne "") + $toappend | Sort-Object | Get-Unique ) -join $jointxt
}

#Used internally to send an email to yourself with the results of the script
Function Send-OlMail {
    [CmdletBinding()]
    param(
        [Parameter()]
        [String]$Subj,
        [Parameter()]
        [String]$Body
    )
    
    $m = $outlook.CreateItem($olItemTypes::olMailItem)
    $m.Subject = $Subj
    $m.Body = $Body
    $m.Recipients.Add($myaddy) | out-null
    $m.Recipients.ResolveAll()
    $m.Send()
}

#Delete stuff matching criteria. See examples below.
Function Delete-Stuff {
    [CmdletBinding()]
    param(
        [Object]$mbox,
        [System.DateTime]$recd,
        [System.DateTime]$expy,
        [string[]]$yesCats,
        [string[]]$noCats,
        [string]$otherParam="",
        [string]$otherVal
    )

    $filt=$null

    If($expy -ne $null) {
        $filt = Safe-Append $filt "[ExpiryTime] < '$($expy.ToShortDateString())'"
    }
    If($recd -ne $null) {
        $filt = Safe-Append $filt "[ReceivedTime] < '$($recd.ToShortDateString())'"
    }
    $yesCats | ForEach-Object {If($PSItem -ne $null -and $PSItem -ne '') { $filt = Safe-Append $filt "[Categories] = '$PSItem'"} }
    $noCats | ForEach-Object {If($PSItem -ne $null -and $PSItem -ne '') { $filt = Safe-Append $filt "Not [Categories] = '$PSItem'"} }
    If($otherParam -ne "") {
        $filt = Safe-Append $filt ("[$otherParam] = '$otherVal'")
    }

    $items = $mbox.Items.Restrict($filt)
    $agetxt = If($recd -ne $null) {$recd.ToShortDateString()} elseif($expy -ne $null) {$expy.ToShortDateString()} else {"n/a"}
    Add-Mail "$($mbox.Name) Age:$agetxt Including:$yesCats Excluding:$noCats Count:$($items.Count)"
    $itm = $items.GetLast()
    While($itm -ne $null) {
        Add-Mail "    delete $($mbox.Name) $($itm.Subject) From($($itm.SenderName)) Cats($($itm.Categories)) received $($itm.ReceivedTime)"
        $itm.ExpiryTime = $expiry
        $itm.Save()
        $itm.Delete()

        $itm = $items.GetLast()
    }
}

# Add categories and/or move to destination folder and/or mark read all emails
#  matching filters.
Function Modify-MailItem {
    [CmdletBinding()]
    param(
        [Object]$mbox,
        [string[]]$filters,
        [string[]]$cats,
        [Object]$destinationFolder,
        [boolean]$markread = $false
    )

    $filt=$null

    $filters | ForEach-Object {If($PSItem -ne $null -and $PSItem -ne '') { $filt = Safe-Append -dest $filt -toappend $PSItem -jointxt " And "} }
    $cats | ForEach-Object {If($PSItem -ne $null -and $PSItem -ne '') { $filt = Safe-Append -dest $filt -toappend "Not [Categories] = '$PSItem'" -jointxt " And "} }
    $items = $mbox.Items.Restrict($filt)
    Add-Mail "$($mbox.Name) Including:$filters Adding Cats:$cats Count:$($items.Count)"
    $itm = $items.GetFirst()
    While($itm -ne $null) {
        Add-Mail "    Addcat $($mbox.Name) $($itm.Subject) From($($itm.SenderName)) Cats($($itm.Categories)) received $($itm.ReceivedTime)"
        $itm.Categories = Safe-Append -dest $itm.Categories -toappend $cats -jointxt ", "
        if($markread) {
            $itm.UnRead = $false
        }
        $itm.Save()
        if($destinationFolder) {
            $itm.Move($destinationFolder)
        }
        $itm = $items.GetNext()
    }
}

# Add "Deletable Alerts" category to all emails from spammy senders, mark as read.
# Next time you run the script, they will be deleted by one of the calls to
# Delete-Stuff (assuming you have uncommented the right stuff)
Function Mark-Spammers {
    [CmdletBinding()]
    param(
        [Object]$mbox,
        [string[]]$emails
    )

    Modify-MailItem -mbox $mbox -filters ( ( $emails | ForEach-Object { "[SenderEmailAddress] = '$PSItem'" } ) -join " Or ") -cats "Deletable Alerts" -markread $true
}

# We handle PTO by creating a meeting request and sending it to the whole
#   department. This will add all of these items to your calendar by accepting
#   the requests, but will not reply to the sender.
# The meeting requests must have "PTO" or "WFH" in the subject and you haven't
#   responded yet.
function Reply-PTO {
    $newitems = $inbox.Items.Restrict("[ReceivedTime] > '$($cutoff7d.ToShortDateString())'")
    $msg = $newitems.GetLast()
    While($msg -ne $null) {
        if ($msg.MessageClass -eq 'IPM.Schedule.Meeting.Request') {
            Add-Mail "Meeting Request $($msg.Subject) From($($msg.SenderName)) received $($msg.ReceivedTime)"
            if ($msg.Subject -like '*PTO*' -or $msg.Subject -like '*WFH*') {
                $x = $msg.GetAssociatedAppointment($true)
                if ($x.ResponseStatus -eq 5) {
                    Add-Mail "    responding to PTO $($msg.Subject) From($($msg.SenderName)) received $($msg.ReceivedTime)"
                    $respitem = $x.Respond($olMeetingResponse::olMeetingAccepted, $true, $false)
                    #$respitem.Send()
                    $msg.UnRead = $false
                    $msg.Save()
                    $msg.Delete()
                } else {
                    Add-Mail "    Already responded"
                }
            } else {
                Add-Mail "    Subject did not match"
            }
        }
        $msg = $newitems.GetPrevious()
    }
}

#Forwards your work meetings to your home email. Requires setup at the top
# of the script. This is probably a violation of your workplace policies. Do not
# do this.
function Forward-Meetings {

    #Only today's stuff
    $startdate = $cutoff.ToString("MM/dd/yyyy 12:00 A\M")
    $enddate = $cutoff7d.ToString("MM/dd/yyyy 11:59 P\M")
    $calir = $cali.Restrict("[Start] <= '" + $enddate + "' AND [End] >= '" + $startdate + "'")

    $calitm = $calir.GetFirst()
    Add-Mail "Calendar Items: $($calir.Count)"
    while ($calitm -ne $null) {
        if ($homemeetings_include_subject) {
            $subject = $calitm.Subject
            $location = $calitm.Location
        } else {
            $subject = "Work meeting"
            $location = "Work"
        }
        $startdate = $calitm.Start.ToUniversalTime() | Get-Date -UFormat "%Y%m%dT%H%M%SZ"
        $enddate = $calitm.End.ToUniversalTime() | Get-Date -UFormat "%Y%m%dT%H%M%SZ"
        $tstamp = (Get-Date).ToUniversalTime() | Get-Date -UFormat "%Y%m%dT%H%M%SZ"
        $globalid = $calitm.GlobalAppointmentID
        $created = $calitm.CreationTime.ToUniversalTime() | Get-Date -UFormat "%Y%m%dT%H%M%SZ"
        $modified = $calitm.LastModificationTime.ToUniversalTime() | Get-Date -UFormat "%Y%m%dT%H%M%SZ"

        $emailbody = @"

Type       : Single Meeting
Organizer  : $($myaddy)
Start Time : $($calitm.Start)
End Time   : $($calitm.End)
Time Zone  : Local time
Location   : $location

"@

        $attachmentbody= @"
BEGIN:VCALENDAR
METHOD:REQUEST
PRODID:Microsoft CDO for Microsoft Exchange
VERSION:2.0
BEGIN:VEVENT
DTSTAMP:$tstamp
DTSTART:$startdate
SUMMARY:$subject
UID:$globalid
ATTENDEE;ROLE=REQ-PARTICIPANT;PARTSTAT=NEEDS-ACTION;RSVP=TRUE;CN='$homecalendar':MAILTO:$homecalendar
ORGANIZER;CN=$($myfullname):MAILTO:$($myaddy)
LOCATION:$location
DTEND:$enddate
SEQUENCE:0
PRIORITY:5
CLASS:PUBLIC
CREATED:$created
LAST-MODIFIED:$modified
STATUS:CONFIRMED
TRANSP:OPAQUE
END:VEVENT
END:VCALENDAR
"@

#$attachmentbody | Out-File -FilePath $attach_filename

        $_smtpClient = New-Object Net.Mail.SmtpClient
        $_smtpClient.Host = $smtp
        $_smtpClient.Port = 25
        $_smtpClient.UseDefaultCredentials = $true

        $_message = New-Object Net.Mail.MailMessage
        $_message.From = $myaddy
        $_message.To.Add("$homecalendar")
        if ($homemeetings_include_subject) {
            $_message.Subject = "Work meeting: $subject"
        } else {
            $_message.Subject = "Work meeting"
        }
        $_message.AlternateViews.Add([Net.Mail.AlternateView]::CreateAlternateViewFromString($emailbody, 'text/plain'))
        $_message.AlternateViews.Add([Net.Mail.AlternateView]::CreateAlternateViewFromString($attachmentbody, 'text/calendar; method=REQUEST; charset="utf-8"'))
        Add-Mail "Sending meeting invite/update for $subject"
        $_smtpClient.Send($_message)

        #$calitm = $null        
        $calitm = $calir.GetNext()  
    }
}


#Here are examples of how to use this stuff. Uncomment/edit as needed. I created
#a scheduled task to periodically run this script for me.

#auto-accept PTO / WFH fyi meetings
Reply-PTO

#delete deleted items older than a day.
#Delete-Stuff -mbox $deleted_items -expy $cutoff1d
#
#delete really old alerts, honoring retention categores I've made.
# PREREQS: create a bunch of categories in Outlook with these names:
#   Keep 7 Years
#   Keep 1 Year
#   Keep 90 Days
#   Keep 30 Days
#   Deletable Alerts

#Delete-Stuff -mbox $alerts -recd $cutoff7y
#Delete-Stuff -mbox $alerts -recd $cutoff1y -noCats  @('Keep 7 Years')
#Delete-Stuff -mbox $alerts -recd $cutoff90d -noCats  @('Keep 7 Years','Keep 1 Year')
#Delete-Stuff -mbox $alerts -recd $cutoff   -noCats  @('Keep 7 Years','Keep 1 Year','Keep 90 Days')
#Delete-Stuff -mbox $alerts -recd $cutoff7d -noCats  @('Keep 7 Years','Keep 1 Year','Keep 90 Days','Keep 30 Days')
#Delete-Stuff -mbox $alerts -recd $cutoff1d -yesCats @('Deletable Alerts') -noCats @('Keep 7 Years','Keep 1 Year','Keep 90 Days','Keep 30 Days','Keep 7 Days')
#
##Inbox items must be categorized to be deleted. Otherwise they are kept.
#Delete-Stuff -mbox $inbox  -recd $cutoff7y
#Delete-Stuff -mbox $inbox  -recd $cutoff1y -yesCats @('Keep 1 Year')      -noCats @('Keep 7 Years')
#Delete-Stuff -mbox $inbox  -recd $cutoff90d -yesCats @('Keep 90 Days')     -noCats @('Keep 7 Years','Keep 1 Year')
#Delete-Stuff -mbox $inbox  -recd $cutoff   -yesCats @('Keep 30 Days')     -noCats @('Keep 7 Years','Keep 1 Year','Keep 90 Days')
#Delete-Stuff -mbox $inbox  -recd $cutoff7d -yesCats @('Keep 7 Days')      -noCats @('Keep 7 Years','Keep 1 Year','Keep 90 Days','Keep 30 Days')
#Delete-Stuff -mbox $inbox  -recd $cutoff1d -yesCats @('Deletable Alerts') -noCats @('Keep 7 Years','Keep 1 Year','Keep 90 Days','Keep 30 Days','Keep 7 Days')
#
#
#Modify-MailItem -mbox $inbox -filters @("[Subject] = 'Programming Tasks'", "[SenderName] = 'Joe Programboss'") -cats @("Deletable Alerts") -markread $true
#Modify-MailItem -mbox $inbox -filters @("[Subject] = 'Splunk Alert: Authentication Failure - Account Locked Out'", "[SenderName] = 'splunk'") -cats @("Keep 90 Days") -markread $true
#Modify-MailItem -mbox $inbox -filters @("[Subject] = 'Re: Authentication Failure - Account Locked Out'") -cats @("Keep 90 Days") -markread $true
#Modify-MailItem -mbox $inbox -filters @("[Subject] = 'Fw: Authentication Failure - Account Locked Out'") -cats @("Keep 90 Days") -markread $true
#Modify-MailItem -mbox $inbox -filters @("( [Subject] = 'Today' Or [Subject] = 'Late' Or [Subject] = 'Out' )", "( [SenderName] = 'Boss Lady' Or [SenderName] = 'Coworker Dudette' Or [SenderName] = 'Fellow Itnerd' Or [SenderName] = 'James Exampleman' )") -cats @("Deletable Alerts") -markread $true
#
#Modify-MailItem -mbox $inbox -filters @("[SenderEmailAddress] = 'US-CERT@ncas.us-cert.gov'") -cats @("Keep 7 Days")
#Modify-MailItem -mbox $inbox -filters @("[Subject] = 'Websense Alert: Approaching Subscription Limit'", "[SenderEmailAddress] = 'administrator@wherever.com'") -cats @("Deletable Alerts")
#Modify-MailItem -mbox $inbox -filters @("[Subject] = 'Mail Queue Alert'", "[SenderEmailAddress] = 'QueueAlerts@example.com'") -cats @("Deletable Alerts")
#Modify-MailItem -mbox $inbox -filters @("[To] = 'Production Stuff'") -cats @("Keep 90 Days")
#Modify-MailItem -mbox $inbox -filters @("[Subject] = 'Message Notification'", "[SenderEmailAddress] = 'MAILER-DAEMON@ironport.example.com'") -cats @("Deletable Alerts") -destinationFolder $dlpfolder
#Modify-MailItem -mbox $inbox -filters @("[Subject] = 'Splunk Alert: AD Group Change Summary by User'", "[SenderName] = 'splunk'") -cats @("Keep 1 Year")
#Modify-MailItem -mbox $inbox -filters @("[Subject] = 'Splunk Alert: SMTP Traffic'", "[SenderName] = 'splunk'") -cats @("Keep 90 Days") -destinationFolder $alerts
#Modify-MailItem -mbox $inbox -filters @("[Subject] = 'Splunk Alert: Denied Ports - DNS'", "[SenderName] = 'splunk'") -cats @("Deletable Alerts") -destinationFolder $alerts
#Modify-MailItem -mbox $inbox -filters @("[Subject] = 'Splunk Alert: Denied Ports - DNS - Serious'", "[SenderName] = 'splunk'") -cats @("Keep 1 Year")
#Modify-MailItem -mbox $inbox -filters @("[Subject] = 'Splunk Alert: monitored groups alert'", "[SenderName] = 'splunk'") -cats @("Keep 1 Year")
#Modify-MailItem -mbox $inbox -filters @("[Subject] = 'High CPU utilization(95%)'", "[SenderEmailAddress] = 'solarwinds@example.com'") -cats @("Deletable Alerts") -destinationFolder $alerts
#
#
#Mark-Spammers -mbox $inbox -emails ("sender@example1.com","sender2@example2.net","someoneelse@foobar.com")
#

#Here's an example of something a little more complicated. This looks through 
#the inbox for very short emails, like "OK", or "I'll be late" sent internally
#and marks them as deletable after 30 days. I made this when I worked at a place
#with draconian tiny mailbox quotas.

#$inbox_items = $inbox.Items.Restrict("Not ( [Categories] = 'Deletable Alerts' Or [Categories] = 'Keep 7 Years' Or [Categories] = 'Keep 1 Year'" + 
#    " Or [Categories] = 'Keep 90 Days' Or [Categories] = 'Keep 30 Days' ) And [SenderEmailType] = 'EX'")
#Add-Mail "Inbox Search For Tiny Emails $($inbox_items.Count)"
#$itm = $inbox_items.GetFirst()
#$i=0
#While($itm -ne $null) {
##    If(($i % 100) -eq 0) { Add-Mail $i }
##    Add-Mail "Inbox Tiny Emails Search $($itm.Subject) Cats($($itm.Categories)) received $($itm.ReceivedTime)"
#        If($itm.Body.Length -le 2000) {
#            $b = $itm.Body
#
#            #Cut off the signature, if it exists.
#            $sig_pos = $b.IndexOf($itm.SenderName)
#            #Paranoid: The signature starts a new line, so previous character is \n.
#            If($sig_pos -gt 0 -and $b[$sig_pos - 1] -eq [System.Char](10) ) {
#                $b = $b.Substring(0,$sig_pos-2)
#            }
#
#            $body_lines = $b | Measure-Object -Line | Select -Expand Lines
#            If( $body_lines -le 2) {
#                Add-Mail "Inbox Tiny Email Added D30D $($itm.Subject) Cats($($itm.Categories)) received $($itm.ReceivedTime) Lines($body_lines)"
#                $itm.Categories += ",Keep 30 Days"
#                $itm.Save()
#            }
#        }
#    $i+=1
#    $itm = $inbox_items.GetNext()
#}

#Finally, send an email to yourself with all the text from Add-Mail above.
Send-OlMail -Body $email_txt -Subj "Process-Inbox Results" 


