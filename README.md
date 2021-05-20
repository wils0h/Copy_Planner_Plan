# Copy_Planner_Plan
A better way to copy M365 Group's Planner Plan(s) than MSFT's built-in Copy Plan function.

Created by: wils0h

Version 1.02

Last Update Date: May 19, 2021

Language: POWERSHELL

Required PowerShell Modules: MicrosoftTeams & ExchangeOnlineManagement
****************************************************************************************
This script is based on the PlannerMigration.ps1 script created by Github user smbm.

I made significant changes to smbm's script, mainly because that script was meant for copying Plans to Groups in different M365 tenants. In my use case, Plans needeh to be copied to Groups within the same M365 tenant.
Many of the original script's functions relied on array item comparisons in order for the script to proceed, so those had to be modified or removed.
This script requires the user to input the source Group and destination Group for copying all Plans. Beyond that, the script runs alone without any other input.
The original script was also missing a few types of Plan data that Graph API can copy to the new destination Plan, so I added as many copier functions as possible.

This script copies:
1. Task start and due dates
2. Task progress: not started, in progress, completed
3. Task notes
4. Task comments, who completed the task, and when the task was completed. The aforementioned info is added to the task notes section
5. If the task is completed, then the task preview will be set as "description" so the task completion information is visible on the task card. Task completeion data is always added to the top of the notes section, even if other data is already present in the notes section
6. Task's assigned users
7. Task checklist items and completion status of checklist items
8. Task labels

Script limitations:
1. Checklist items are out of order when there is 1 or more completed checklist item.
2. Can't copy task attachments.
****************************************************************************************
My inspiration for creating this script was my general dissappointment in Microsoft's built-in GUI Plan copying function that does not return task comments, task completion status, checklist item completion status, task assignments, and who completed the task. I think that it is unreasonable that those features were not added when the features are clearly accessible through the Microsoft Graph API.

Last but not least, I believe in freely sharing this kind of time-saving information because you should not have to pay a vendor thousands of dollars for this. Or pay them a lot and not even get this Plan copying solution.
****************************************************************************************
