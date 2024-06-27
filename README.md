# PowerReview
PowerReview is a PowerShell implementation of baseline log review on extracted Windows Event Logs.
```
Usage:
  PowerReview.ps1 -sd <StartDate> -ed <EndDate> -ep <Base Directory>

Description:
  PowerReview is a PowerShell implementation of baseline log review on extracted Windows Event Logs.

Options:
  -sd    The start date in format YYYY-MM-DD.
  -ed    The end date in format YYYY-MM-DD.
  -ep    The base directory containing the evtx files

Note:
  The output of this script is based on the timezone of the computer it ran on.
```
