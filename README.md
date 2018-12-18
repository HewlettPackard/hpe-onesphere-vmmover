# Script goal

The goal of this script is to dispatch VMs created by HPE OneSphere in VCenter Datacenter root folder, to their appropriate project folders. We use an Excel spreadsheet to handle mappings of OneSphere Projects to VCenter folders. Several options are available to handle nonexisting target folders (-CreateTarget=never,always,only) and default behaviour will not create the folder and not move the VM (never).

The -reset option (mutually exclusive with -createTarget) will move ALL Onesphere managed VM back to root of Datacenter folder, allowing to revert changes and/or rework the folder structure.

The script takes also a -ExcelFilename parameter to specify the name of the Excel spreadsheet for the mappings. It defaults to ProjectMapping.xlsx in the same location as the script.
 
# Dependencies

This script requires 2 PowerShell modules, one for manipulating Excel spreadsheets and the second for interfacing with VMware VCenter. These tow modules can be installed from the Microsoft PowerShell Gallery with the following instructions:

```` PowerShell
install-module importexcel
install-module VMware.VimAutomation.Core
````
# Examples of calls
```` PowerShell
.\OneSphereVMMover.ps1 -VCServerName vcenter-os1.etss.lab -DatacenterName NHITS-DC -Username XXX -Password XXX -CreateTarget always
.\OneSphereVMMover.ps1 -VCServerName vcenter-os1.etss.lab -DatacenterName NHITS-DC -Username XXX -Password XXX -CreateTarget always -verbose
.\OneSphereVMMover.ps1 -VCServerName vcenter-os1.etss.lab -DatacenterName NHITS-DC -Username XXX -Password XXX -CreateTarget never
.\OneSphereVMMover.ps1 -VCServerName vcenter-os1.etss.lab -DatacenterName NHITS-DC -Username XXX -Password XXX -CreateTarget only
.\OneSphereVMMover.ps1 -VCServerName vcenter-os1.etss.lab -DatacenterName NHITS-DC -Username XXX -Password XXX -reset
.\OneSphereVMMover.ps1 -VCServerName vcenter-os1.etss.lab -DatacenterName NHITS-DC -Username XXX -Password XXX -ExcelFilename myfile.xls
````
