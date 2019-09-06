# Outlook Rule Monitor
Monitor the mailbox rules of all users and aggregate the information into an emailed report/notification.

This script is meant to run as a _Scheduled Task_ on an _Exchange Server_.
On its initial run, it will collect **all** client-side rules and **all** forward/delete rules that are stored server-side.

### Why Client-Side Rules?
Notably, Exchange does not know the _action_ being taken on client-side rules, only the _conditions_ to match before flagging it for the client's action later (when they log on).

Because of this, it's helpful for an administrator to know when a new client-side rule is added so that they can see if, for example, a _permanently delete_ rule is added for emails coming from the organization's CEO.

Again, you will not know (and cannot know from the Exchange Server) **which action** the client is taking on their rule, only **which conditions** engage the rule.


## Parameters
### CheckNow
**Optional** switch. Scan through all inboxes and instead of reporting changes including delete/forward rules, report on _any_ existing rule that is either forwarding or deleting inbound emails (as well as client-side rules).


## Tweaks
The _"Tweaks_" subsection near the top of the script is used to define static variables that are later expanded in the script, to change its functionality.
Each tweak will include its own description (if its variable name isn't descriptive enough).

#### Current Tweaks Variables
+ **$NotificationsAddress** -- The target email address to which notifications are sent.
+ **$NotificationsSource** -- The "From" address of emails sent from the script.
+ **$NotificationsSubject** -- The subject line used in email notifications.
+ **$NotificationsServer** -- A target server used to relay emails to their destination. This is _required_ to send notifications.
+ **$NotificationsServerPort** -- The relay server's target port.
+ **$NotificationsStyling** -- CSS styling for notifications, included in the HTML \<head\> tag for each notification.
+ **$NotificationsDaily** -- Whether or not to notify the target email address in the given time-range, regardless of whether or not any changes are detected. This is mostly useful as a "_heartbeat_" to let an administrator know the task is still running daily.
  + **$NotificationsDailyTimeRange**
  + **$NotificationsDailyTimeStart**
  
> ```
> <# A time range within which notifications will be sent to the target as per the above setting.
>  #    The "range" is an amount of minutes as a window for sending the notification.
>  #    The "start" is the beginning of the range variable.
>  # 
>  # Keep in mind, if the script is scheduled to run once every 10 minutes and the Range is set to 60 minutes,
>  #    you will receive notifications every time the script runs through those 60 minutes (not recommended).
>  #
>  # Example: Range: 10, Start: "18:05"
>  # If the script runs every 10 minutes (like so */10 -- 00, 10, 20, etc) this will dispatch a daily notification at
>  #    around 6:10 PM when the script runs.
>  #>
> ```

+ **$ReportName** -- The filename of the _report_ that the script maintains at the given location (this can be a full path), to index changes across mailboxes.
+ **$DesiredProperties** -- When a rule matches, which properties/fields are selected by the script to enter into the HTML table.


## Notifications
Notifications are designed to be generated to a target SMTP relay server (using the Tweaks section), and require HTML formatting.

Below is a sample notification. Since this is from a _production environment_, I've blacked out a few identifying pieces of information.
Most important in the redacted pieces is the thick black line over the table: that is the email address (i.e. _inbox_) that has added the particular rules in the table underneath it.

![Sample Notification from the Monitoring Script](https://raw.githubusercontent.com/NotsoanoNimus/outlook-rule-monitor/master/docs/Notification_Sample.png)


## TODOs
- [X] Hash rule condition-descriptions to detect when an already-discovered rule is _modified_.
- [ ] Make the scanning of client-side rules _optional_, so administrators can exclude these results if desired.
- [ ] Create an option for plaintext-only email notifications.
- [ ] Add more detailed filtering mechanisms (perhaps regex-based) to exclude results from a notification.