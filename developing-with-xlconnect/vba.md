---
description: Control scripts from VBA
---

# VBA

The XLConnect add-in can be controlled from VBA for a better user experience when users click on a button, doubleclick a cell, change a value etc.

Below is an example that runs the script named "Pull TrialBalance"

```
Sub PullConnections()
    Set xlc = Application.COMAddIns("XLConnect").Object
    xlc.fetchData "Pull TrialBalance"
End Sub
```

## To use this in your workbook:&#x20;

* [Import script ](import-scripts.md)"Connections"
* On the `Insert` ribbon of Excel, insert an Icon
* Right-click the icon and select `Assign Macro...`
* In the Assign Macro screen, select `New`

<figure><img src="../.gitbook/assets/image (6).png" alt=""><figcaption></figcaption></figure>

* This should bring you to the VBA editor. If not, press ALT-F11.&#x20;
* Paste in the two lines so your code looks like this:

```
Sub Graphic2_Click()
    Set xlc = Application.COMAddIns("XLConnect").Object
    xlc.fetchData "Connections"
End Sub
```

Make sure the name of the VBA method corresponds to the name in the `Assign Macro` editor in Excel (in this case `Graphic2_Click`).

* Go back to Excel and click the icon, that should now run the Graphic2\_Click VBA method, that will execute the Connections script.&#x20;

This is how simple it is to execute a script by clicking an icon.&#x20;



## Available methods



