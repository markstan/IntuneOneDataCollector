# Intune One Data Collector

Intune One Data Collector (ODC) is a support script to enable the collection of logs, Registry data, and command output from Intune Windows 10 and Windows 11 clients.  IntuneODCStandAlone.ps1 ingests the Intune.xml file to automate data collection.  Once the data is gathered, the script will open an Explorer window in the folder where the ODC contents are located. The default file name is COMPUTERNAME_CollectedData.zip (where COMPUTERNAME is the name of the Windows device).
 
 
To run this tool, open an elevated PowerShell window (right-click, "Run as administrator..."), create a temporary folder, then run these three commands:

         wget https://aka.ms/intunexml -outfile Intune.xml
         wget https://aka.ms/intuneps1 -outfile IntuneODCStandAlone.ps1
         PowerShell -ExecutionPolicy Bypass -File .\IntuneODCStandAlone.ps1

(**Hint**: You can copy and paste the commands from this page and paste them directly in to the PowerShell window).

The first two commands download the XML driver file and the PowerShell script.  The last line runs the script (both files must be in the same folder).

![Example PowerShell commands and Explorer window](https://github.com/markstan/IntuneOneDataCollector/blob/master/Resources/PS_Example.png)


If you have any problems downloading the files using the commands above (usually due to network or firewall restrictions), you can also click on the green 'Code' button above and choose **Download ZIP** to save the contents of this project.

![Download zip](https://github.com/markstan/IntuneOneDataCollector/blob/master/Resources/Download_Zip.png)

For more information, please refer to [the FAQ](https://github.com/markstan/IntuneOneDataCollector/wiki/FAQ) for this project.

Please contact [markstan@microsoft.com](mailto:markstan@microsoft.com) for any bug reports or feature requests.
