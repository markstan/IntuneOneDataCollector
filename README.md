# Intune One Data Collector

 This package is used in conjunction with the Microsoft Support One Data Collector SDP tool to gather data from Windows Intune client machines.

The XML file contains data locations and scripts to collect a variety of files, registry keys, and command line output to assist support engineers in troubleshooting Intune issues.

Download the Intune.xml file [here](https://raw.githubusercontent.com/markstan/IntuneOneDataCollector/master/Intune.xml).

To run this tool without a Microsoft support incident, open an elevated ("Run as administrator...") PowerShell window and run these three commands:

         wget https://aka.ms/intunexml -outfile Intune.xml
         wget https://aka.ms/intuneps1 -outfile IntuneODCStandAlone.ps1
         PowerShell -ExecutionPolicy Bypass -File .\IntuneODCStandAlone.ps1

The first two commands download the XML driver file and the PowerShell script.  The last line will actually run the script (assuming both files are in the same folder).  When the script finishes its execution, it will automatically launch an Explorer window showing the output of the data collection.

Please contact [markstan@microsoft.com](mailto:markstan@microsoft.com) for any bug reports or feature requests.