# Task Sequence Powershell Module
Awhile ago I submitted a request to Microsoft for a built-in Powershell Module during the Task Sequence. Mostly for basic things like getting and setting variables, but once I noticed what all the COM Objects exposed, I wanted to take this a step further.

Primarily you’ll have 2 COM objects you can create during a Task Sequence: Microsoft.SMS.TSEnvironment (getting/setting variables) and Microsoft.SMS.TSProgressUI (manipulating the Progress and messages). I’ve uploaded the module for you all to use at my github. I’ll run though some of the details below including visuals for the TSProgressUI items. Some, if not most, of these are-for lack of a more appropriate label-useless; but can still be fun to mess around with!


## Microsoft.SMS.TSEnvironment
### Confirm-TSEnvironment
This function is not exported, as it is used within the module to verify that the Microsoft.SMS.TSEnvironment COM Object is loaded before attempting any operations.

### Get-TSVariables
Calls `GetVariables()` on the COM Object. This will return an array of all Task Sequence Variable NAMES only.

### Get-TSValue
Calls `Value()` on the COM Object to return the value of a specific variable. the “Name” Parameter should be the Task Sequence Variable you want the value for.

### Set-TSVariables
Assigns a value to a specific variable. A Boolean return indicates success. Note this is wrapped in a try/catch in the event you attempt to set a read-only variable.

### Get-TSAllValues
Returns an object of all Task Sequence Variables and their assigned Values.

## Microsoft.SMS.TSProgressUI
### Confirm-TSProgressUISetup
This function is not exported, as it is used within the module to verify that the Microsoft.SMS.TSProgressUI COM Object is loaded before attempting any operations.

### Show-TSActionProgress
Calls `ShowActionProgress()` on the COM object. This is the function I use 99.9% of the time for UI related actions. This will add the “second” progress bar and update the Text for that progress bar as well as fill the progress meter as you see fit. The neat thing here, is that the “Step” and “MaxStep” will automatically show a “100%” representation. For example, if i have 8 of 400 complete, I don’t have to do the extra scripting to pass in “2” and “100”. I can pass in 8, and 400, and the COM Object does the math for us.

There are other parameters here with “ShowActionProgress()”. With this function, those use the built-in variables “designed” to be used here. If you want to extend this function to allow those additional parameters, feel free. Personally, I don’t find them useful enough to include.

### Close-TSProgress
Calls `CloseProgressDialog()` on the COM object. Closes or “hides” the progress bar window. This will automatically reappear once the current step has completed, or if you call one of the “show” functions for the UI.

### Show-TSProgress
Calls `ShowTSProgress()` on the COM object. This function manipulates the “Top part” of the progress bar. If, for whatever reason, you wanted to manipulate the progress bar without adding the “second/bottom” bar, this is what you should use.

### Show-TSErrorDialog
Calls `ShowErrorDialog()` on the COM object. This wills how the typical “Error in the Task Sequence” Window. There are parameters for the error title, error message/details, exit code, a timeout, and a force reboot. The force reboot will trigger a reboot after the timeout (in seconds) has expired.

### Show-TSMessage
Calls `ShowMessage()` on the COM Object. This essentially shows a MessageBox allowing you to specify the button configuration. The interesting (and somewhat useless) part of this, is that you do not receive a return value. If you want a simple “OK” message box, this should work, otherwise I’d recommend using a [Normal MessageBox](https://msdn.microsoft.com/en-us/library/system.windows.forms.messagebox.show(v=vs.110).aspx).

*Button Options (Type Parameter):*
* 0 = OK
* 1 = OK, Cancel
* 2 = Abort, Retry, Ignore
* 3 = Yes, No, Cancel
* 4 = Yes, No
* 5 = Retry, Cancel
* 6 = Cancel, Try Again, Continue

### Show-TSRebootDialog
Calls `ShowRebootDialog()` on the COM Object. This will allow you to present a custom reboot window. I would recommend against this, as untimely reboots during a task sequence can cause some issues. This is essentially what causes the issue where [Updates cause Double Reboots during OSD](https://support.microsoft.com/en-us/kb/2894518). If you do need a reboot after a step, either inset a restart computer step in your task sequence or just have your script exit `3010`.

Typically OrganzationName will use the TS Variable `_SMSTSOrgName`.

### Show-TSSwapMediaDialog
Calls `ShowSwapMediaDialog()` on the COM Object. If you’ve ever used standalone media that was broken up into multiple CD/DVD sets (or played a Final Fantasy on Playstation), you’ve seen this one. This is a message indicating for you to swap media, letting you know what disk number to insert.
