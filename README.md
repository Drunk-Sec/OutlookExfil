# OutlookExfil
POC - Work in progress

This is a little project i thought about recently and wanted to see if it'd work. Plenty still to be done but the core concept is there

Exfiltrate data silently from Outlook

Takes action as each email is selected and writes subject, received time, sender, recipients, body and folder path to a text file in docs folder (could easily be modified to use something like a discord webhook or POST elsewhere)
