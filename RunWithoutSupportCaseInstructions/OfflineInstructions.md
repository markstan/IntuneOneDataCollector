# Offline Instructions

1. Download http://download.microsoft.com/download/F/2/2/F22D5FDB-59CD-4275-8C95-1BE17BF70B21/1dc-c.diagcab (you may need to add download.microsoft.com to your trusted sites in your browser settings).
1. Double-click on 1dc-c.diagcab to launch the troubleshooter
1. Choose Send diagnostic information to Microsoft when prompted (there is not an associated case, so this will have no effect).
1. Choose Custom Data Collection from the Data Collector List and click Next.
1. Browse to Intune.xml and click next when prompted.
1. Allow the One Data Collector to complete and click Close when prompted. It is normal to see  Problems found: Intune with a status of Fixed.
1. Open Explorer and navigate to %localappdata%\elevateddiagnostics. Sort by date and open the most recently modified folder. Latest.cab will contain the results of the last run. Uncompressed results are located in a numeric folder in the same location.