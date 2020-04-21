# Check email domains

This VBA macro for Microsoft Outlook warns the user before sending email
to recipients from multiple domains.

## Installation

1. If you can't see the Developer tab in the ribbon bar, go to File/Options, 
   choose Customise Ribbon on the left, and tick Developer on the right.
2. From the Developer tab choose Visual Basic.
3. Expand Project1, Microsoft Outlook Objects, and double-click 
   ThisOutlookSession (top left).
4. Paste the VBA code in `CheckEmailDomains.vba` into the module.
5. Close the VBA editor and save changes to the module.
6. On the Developer tab click Macro Security, and change the level to 
   `Notifications for all macros or lower`.
7. Restart Outlook.
