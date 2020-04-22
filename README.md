# Check email domains

This VBA macro for Microsoft Outlook warns the user before sending email
to recipients from multiple domains.

## Requirements

Microsoft Outlook 2016 or later desktop app running on Microsoft Windows.

## Installation

1. Open the Microsoft Outlook desktop app.
2. If you can't see the Developer tab in the ribbon bar, go to File/Options, 
   choose Customise Ribbon on the left, and tick Developer on the right.
3. From the Developer tab choose Visual Basic.
4. At the top left of the Visual Basic window, expand `Project1`, 
   `Microsoft Outlook Objects`, and double-click `ThisOutlookSession` to open
   the module pane.
5. Copy the [VBA code from here](https://raw.githubusercontent.com/blokeley/check_email_domains/master/CheckEmailDomains.vba) and paste it into the `ThisOutlookSession` module.
6. Close the VBA editor and save changes to the module.
7. On the Developer tab click Macro Security, and change the level to 
   `Notifications for all macros or lower`.
8. Restart Outlook.

Each time you start Outlook you will be asked to enable macros in `ThisOutlookSession`, which you should do.  

On the other hand, do not enable macros for any other modules because
they could be malicious.
