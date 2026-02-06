Stupid easy way to track toner/paper usage among printers on the local network

# PS_Printer_Ink_Level_SNMP
Create a HTML report of printer toner level using SNMP and PowerShell

This script will generate an html page containing all printers described in the printerlist.txt file
For each printer, it will get the level for all cartridge + any existing status

It will retrieve information using SNMP and the default printer MIB
  Input:
       printerlist.txt : csv file containing printers where information will be retrieved.
                         each line contains value,name,description
                         with 
                             value: DNS name of your printer
                             name: Name of the printer as displayed in the report
                             description: Description of the printer that will appear in the report

Once the program has been filled out with the local networks printer information you can use the PS1 script included in the repo to easily host the printer report page via any computer or server.

Create a task to run the printer report weekly on the hosted computer and every week you will see the delta variable of how many pages were printed.

Included in the latest changes was a card view to see printers a different way other than just the table view, percent change per week. This will show toner usage on a week by week basis and lay it out using a percentage. ]
This is tracked using a history CSV that is outputed into from the powershell script and referenced to create the html file.

More upgrades coming soon as my organization is looking to upgrade how they track toner usage among printers and workers.

Truly amazing how this ultra simple powershell script is able to change the game for tracking toner information.
