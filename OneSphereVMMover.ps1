##############################################################################
## (C) Copyright 2017-2018 Hewlett Packard Enterprise Development LP 
##############################################################################
<#
Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:
The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
#>

[CmdletBinding(DefaultParameterSetName="Create")]
Param
(
    # VCenter FQDN or IP address
    [Parameter(Mandatory, ParameterSetName = "Reset")]
    [Parameter(Mandatory, ParameterSetName = "Create")]
    [string]$VCServerName,

    # VCenter Datacenter Name
    [Parameter(Mandatory, ParameterSetName = "Reset")]
    [Parameter(Mandatory, ParameterSetName = "Create")]
    [string]$DatacenterName,

    # VCenter admin username 
    [Parameter(Mandatory, ParameterSetName = "Reset")]
    [Parameter(Mandatory, ParameterSetName = "Create")]
    [string]$Username,

    # VCenter admin password 
    [Parameter(Mandatory, ParameterSetName = "Reset")]
    [Parameter(Mandatory, ParameterSetName = "Create")]
    [string]$Password,
    
    # Optional: Excel file name for project-to-folder mappings
    [Parameter(Mandatory=$False, ParameterSetName = "Reset")]
    [Parameter(Mandatory=$False, ParameterSetName = "Create")]
    [string]$ExcelFilename = "ProjectMapping.xlsx",

    # Optional: reset switch, will move all OneSpehere discovered VM in Datacenter to root folder (no question asked)
    [Parameter(Mandatory=$False, ParameterSetName = "Reset")]
    [switch]$Reset = $False,

    # Optional: createfolder option to never (default), always or only create the folder required to move the VM
    [Parameter(Mandatory=$False, ParameterSetName = "Create")]
    [ValidateSet('never','always','only')]
    [string]$CreateTarget = "never"

    # Examples of calls
    # -----------------
    # .\OneSphereVMMover.ps1 -VCServerName vcenter-os1.etss.lab -DatacenterName NHITS-DC -Username XXX -Password XXX -CreateTarget always
    # .\OneSphereVMMover.ps1 -VCServerName vcenter-os1.etss.lab -DatacenterName NHITS-DC -Username XXX -Password XXX -CreateTarget never
    # .\OneSphereVMMover.ps1 -VCServerName vcenter-os1.etss.lab -DatacenterName NHITS-DC -Username XXX -Password XXX -CreateTarget only
    # .\OneSphereVMMover.ps1 -VCServerName vcenter-os1.etss.lab -DatacenterName NHITS-DC -Username XXX -Password XXX -reset
    # .\OneSphereVMMover.ps1 -VCServerName vcenter-os1.etss.lab -DatacenterName NHITS-DC -Username XXX -Password XXX -ExcelFilename myfile.xls
    #
)

# Constants
$FOLDERSEPARATOR = "/"

# Import dependencies
import-module importexcel
import-module VMware.VimAutomation.Core

# Support Functions
function folder_exists { param ([string] $foldername, [System.Object] $location) 
# Returns if folder exist in VCenter at given location

    try {
        get-folder -name $foldername -location $location -erroraction stop
        return $true
    }
    catch 
    {
        write-verbose "Folder $foldername not found at location $($location.name)"
        return $false
    }
}

function process_multi_level_folder { param ([string] $foldername, [string] $option)
# Process the target folder if it's a multi level folder
# use option passed in command to take action:
# Never: flag missing folder
# Always and Only: create each level if not there and return deepest level VM folder

    $location=$RootFolder
    write-verbose "Processing path $foldername with option $option"
    # split folder using separator 
    $folders = $foldername.split($FOLDERSEPARATOR)
    write-verbose "Path $foldername has $($folders.count) level(s) to process..."
    # now verify existence of each level 
    for ($i=0; ($i -lt $folders.length); $i++) {
        if (folder_exists $folders[$i] $location) {
            # level exists
            write-verbose "... folder $($folders[$i]) exists under $($location.name),  getting next..."
            $location = get-folder $folders[$i] -location $location
        }
        else {
            # level doesn't exist
            switch ($option) {
                "never" {
                    # cannot do anymore as option says never
                    write-verbose "... folder $($folders[$i]) does not exists under $($location.name), cannot proceed."
                    $location = $null
                    return $null
                }
                {"always" , "only" } {
                    # Create this folder in previous folder's parent
                    write-verbose "... folder $($folders[$i]) does not exists under $($location.name), creating it."
                    try {
                        # Just create folder
                        $location = new-folder -name $folders[$i] -location $location
                    }
                    catch {
                        write-verbose  "Error creating new folder $folders[$i] under $($location.name)" 
                        $PSCmdlet.ThrowTerminatingError($_)
                    } 
                }
            }
        }
    }       
    return $location
}

# Load XLS for mapping 
$map = import-excel $EXCELFILENAME

# Connect to vCenter
Connect-VIServer -Server $VCServerName -username $Username -password $Password

# Set Source folder for VMs to move (Default is VM)
$RootFolder = get-folder -name VM  | where Parent -match $DatacenterName

# Retrieve list of VMs to move 
if ($Reset) {
    # If it's a reset we retrieve list of VMs to move from complete VM list 
    $VmList = get-vm 
    write-verbose "Found $($VmList.count) VM candidate for move to root folder"
}
else { 
    # We need the list of VM from the root of the datacenter
    $VmList = get-vm -location $RootFolder -norecursion
    write-verbose "Found $($VmList.count) VM to in root folder for datacenter $DataCenterName"
}

foreach ($vm in $VMlist) 
{    
    $ProjectName = ""
    $ProjectName = $vm.Notes.split() | foreach { $p = $_.split(":") ; if ($p[0] -match "projectname") { $p[1] } }
    
    # Leverage the Notes field set by OneSphere to retrieve the project this VM belongs to 
    if ($ProjectName) 
    {
        Write-verbose "VM $vm from project $ProjectName managed by HPEOneSphere, candidate for move..."
    }
    else    
    {
        # Leave this VM alone it doesn't belong to OneSphere
        Write-verbose "VM $vm does not seem to be managed by HPEOneSphere, skipping..."
        continue
    }

    # Lookup project in EXCEL Mapping table 
    if ($Reset) {
        # For Reset we don't need anything, target is the root folder of the Datacenter
        $ProcessedFolder = $RootFolder
    }
    else {
        $TargetFolder = $map | where {$_.'OneSphere Project' -contains $ProjectName} | select 'vCenter Path'
        if ($TargetFolder) 
        {
            Write-verbose "Mapping for project $ProjectName found in mapping table: target is $($TargetFolder.'vCenter Path'), processing path"
            # Need to process target folder of it has multiple levels
            $ProcessedFolder = process_multi_level_folder -foldername $TargetFolder.'vCenter Path' -option $CreateTarget
            # When we come our of this, we have checked that all folder in the VCenter path exists
        }
        else    
        {
            # We have no mapping for this project, so skipping
            Write-verbose "Project $ProjectName was not found in mapping table, skipping..."
            continue
        }
    }
    
    if ($CreateTarget -eq "only") {
        # don't even try to move VM in $CreateFolder Only mode
        Continue
    }

    try {
        # Attempting to move VM in target folder
        move-vm -VM $vm -location $ProcessedFolder -erroraction stop 
    }
    catch [Exception] {
        write-verbose "Exception trapped while moving VM with option: $CreateTarget"
        if ($CreateTarget -eq "never") { 
                # Normal, we have not created the target folder because option was never, so move failed
                Write-verbose "Target folder for project $ProjectName does not exist, VM $vm was not moved (use -CreateTarget always)"
        }
        else {  
                # It should have moved the Vm, but something went wrong, and it's beyond our control, so raise exception
                write-verbose  "Error moving VM $vm to folder $($ProcessedFolder.name)" 
                $PSCmdlet.ThrowTerminatingError($_)
        }                    
    }
}

# Retrieve list of VMs now in Root folder of Datacenter
$VmList = get-vm -location $RootFolder -norecursion
write-verbose "$($VmList.count) VM left in root folder for datacenter $DataCenterName"

# Disconnect from VCenter
Disconnect-VIServer -Confirm:$False