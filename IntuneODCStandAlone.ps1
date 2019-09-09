<#
    Stand-alone implementation of One Data Collector

#>

# Copyright © 2016, Microsoft Corporation. All rights reserved.
# :: ======================================================= ::

<#
-DESCRIPTION
 Utils_OneDataCollector.ps1 Contains common functionalities used across the One Data Collector diagnostics.

-FUNCTIONS
 Test-PostBack
 Get-UserReadableSize
 Get-ItemSize
 Pop-Message
 Create-ZipFromDirectory
 Update-Progress
 Export-Report
 Get-OSVersion
 Get-ValidPath
 Get-ValidFileName
 Get-NewFileNameIfExists
 Initialize
 Collect-Files
 Collect-RegistryKeys
 Collect-EventLogs
 Collect-Commands
 Get-Packages
 Get-UserSelectedPackages
 Process-Package
 Compress-CollectedDataAndReport
 Get-PackageFromUserAndProcess
#>

#====================================================================================
# Initialize
#====================================================================================
#Import-LocalizedData -BindingVariable Utils_OneDataCollector_Strings -FileName Utils_OneDataCollector

#region Fields

$ResultRootDirectory = [System.IO.Path]::Combine(($env:TEMP), 'CollectedData')
$CompressedResultFileName = "CollectedData.ZIP"
[System.Nullable[bool]] $newZipperAvailable = $null # Stores flag whether [System.IO.Compression.ZipFile] can be used.

#endregion

#==================================================================================
# Functions
#==================================================================================

#region Functions

Function Test-PostBack
{
    [CmdletBinding()]
    PARAM
    (
        [Alias('S')]
        [Parameter(Position = 1, Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [string] $CurrentScriptName
    )
    PROCESS 
    {
        # Writing the trace to current directory
        $CurrentScriptName = ("{0}.temp" -f [System.IO.Path]::GetFileNameWithoutExtension($CurrentScriptName))

        if(Test-Path($CurrentScriptName))
        {
            return $true
        }

        'Executed' >> $CurrentScriptName
        return $false
    }
}

function Get-UserReadableSize
{
    [CmdletBinding()]
    PARAM
    (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [long] $Length
    )
    PROCESS 
    {
        if($Length -gt 1GB)
        {
            $unitValue = 1GB
            $unitText = 'GB'
        }
        elseif($Length -gt 1MB)
        {
            $unitValue = 1MB
            $unitText = 'MB'
        }
        else
        {
            $unitValue = 1KB;
            $unitText = 'KB'
        }        

        return ("{0:N2} $unitText" -f ($Length / $unitValue))
    }
}

function Get-ItemSize
{
    [CmdletBinding()]
    PARAM
    (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [string] $Path
    )
    PROCESS 
    {

        [long] $Length = 0
        if(Test-Path($Path))
        {
            Get-ChildItem $Path -Recurse | ?{ $_.Attributes -ine 'Directory' } | %{ $Length += $_.Length }
        }
        return (Get-UserReadableSize -Length ($Length))
    }
}

function Pop-Message
{
    <#

    .DESCRIPTION
    Displays a message box with given details

    .PARAMETER Message
    The message you want to display.
    
    .PARAMETER Title
    The title of the message box.
    
    .PARAMETER Type
    The type of the message box as defined below. Default is 64.
    - 16: (0x10) Show "Stop Mark" icon.
    - 32: (0x20) Show "Question Mark" icon.
    - 48: (0x30) Show "Exclamation Mark" icon.
    - 64: (0x40) Show "Information Mark" icon.
    
    .EXAMPLE
    Pop-Message -Message 'Your Message' -Caption 'Your Caption' -Type 48

    .EXAMPLE
    Pop-Message -M 'Your Message' -C 'Your Caption' -T 32

    #>
   
    [CmdletBinding()]
    PARAM
    (
        [Alias('M')]
        [Parameter(Position = 1, Mandatory = $true)]
        [string] $Message ="Message",

        [Alias('C')]
        [Parameter(Position = 2, Mandatory = $false)]
        [string] $Caption = "Title",

        [Alias('T')]
        [Parameter(Position = 3, Mandatory = $false)]
        [int] $Type = 64
    )
    PROCESS 
    {
        $popWindow = New-Object -ComObject wscript.shell
        $popWindow.Popup($Message, 0, $Caption, $Type) | Out-Null
        Remove-Variable popWindow
    }
}

function Create-ZipFromDirectory
{
    <#

    .DESCRIPTION
    Creates a ZIP file from a given the directory.

    .PARAMETER SourceDirectory
    The folder with the files you intend to zip.
    
    .PARAMETER ZipFileName
    The zip file that you intend to create
    
    .PARAMETER IncludeParentDirectory
    Setting this option will include the parent directory.
    
    .PARAMETER Overwrite
    Setting this option will overwrite the zip file if already exits.
    
    .EXAMPLE
    Create-ZipFromDirectory -Source $ResultRootDirectory -ZipFileName $CompressedResultFileName -IncludeParentDirectory -Overwrite

    .EXAMPLE
    Create-ZipFromDirectory -S $ResultRootDirectory -O $CompressedResultFileName -Rooted -Force

    #>

    PARAM
    (
        [Alias('S')]
        [Parameter(Position = 1, Mandatory = $true)]
        [ValidateScript({Test-Path -Path $_})]
        [string]$SourceDirectory,
     
        [Alias('O')]
        [parameter(Position = 2, Mandatory = $false)]
        [string]$ZipFileName,

        [Alias('Rooted')]
        [Parameter(Mandatory = $false)]
        [switch]$IncludeParentDirectory,

        [Alias('Force')]
        [Parameter(Mandatory = $false)]
        [switch]$Overwrite
    )
    PROCESS
    {
        $ZipFileName = (("{0}.zip" -f $ZipFileName), $ZipFileName)[$ZipFileName.EndsWith('.zip', [System.StringComparison]::OrdinalIgnoreCase)]

        if(![System.IO.Path]::IsPathRooted($ZipFileName))
        {
            $ZipFileName = ("{0}\{1}" -f (Get-Location), $ZipFileName)
        }
                    
        if($Overwrite)
        {
           if(Test-Path($ZipFileName)){ Remove-Item $ZipFileName -Force -ErrorAction SilentlyContinue }
        }
        
        $source = Get-Item $SourceDirectory

        if ($source.PSIsContainer)
        {
            if($newZipperAvailable -eq $null)
            {
                try
                {
                    $ErrorActionPreference = 'Stop'
                    Add-Type -AssemblyName System.IO.Compression.FileSystem
                    $newZipperAvailable = $true
                }
                catch
                {
                    $newZipperAvailable = $false
                }
            }

            if($newZipperAvailable -eq $true) # More efficent and works silently.
            {
                [System.IO.Compression.ZipFile]::CreateFromDirectory($source.FullName, $ZipFileName, [System.IO.Compression.CompressionLevel]::Optimal, $IncludeParentDirectory)
            }
            else # Will show progress dialog.
            {

                # Preparing zip if not available.
                if(-not(Test-Path($ZipFileName)))
                {
                    Set-Content $ZipFileName (“PK” + [char]5 + [char]6 + (“$([char]0)” * 18))
                    (dir $ZipFileName).IsReadOnly = $false
                }

                if(-not $IncludeParentDirectory)
                {
                    $source = Get-ChildItem $SourceDirectory
                }
            
                $zipPackage = (New-Object -ComObject Shell.Application).NameSpace($ZipFileName)
        
                [System.Int32]$NoProgressDialog = 16 #Tried but not effective.
                foreach($file in $source)
                { 
                    $zipPackage.CopyHere($file.FullName, $NoProgressDialog)
                    do
                    {
                        Start-Sleep -Milliseconds 256
                    }
                    while ($zipPackage.Items().count -eq 0) # Waiting for an operation to complete.
                }
            }
        }
        else
        {
            Write-Error 'The directory name is invalid.'
        }
    }
}

function Update-Progress
{
    <#

    .DESCRIPTION
    Updates the diagnostics progress for the items being collected.

    .PARAMETER Activity
    Activity which is being in progress. Possible Values are Files, RegistryKeys, EventLogs, and Commands.
  
    .PARAMETER PackageID
    The ID of a package which is being processed.

    .PARAMETER Filename
    The name of a file which is being collected.

    .EXAMPLE
    Update-Progress -Activity 'FILES' -PackageID 'Apps' -Filename 'SomeFile.log'
   
    #>

    [CmdletBinding()]
    PARAM
    (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateScript({-not([string]::IsNullOrEmpty($_)) -or ($_.Trim().Length -le 0)})]
        [string] $Activity,
        
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateScript({-not([string]::IsNullOrEmpty($_)) -or ($_.Trim().Length -le 0)})]
        [string] $PackageID,

        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateScript({-not([string]::IsNullOrEmpty($_)) -or ($_.Trim().Length -le 0)})]
        [string] $Filename
    )
    PROCESS
    {
        Write-DiagProgress -Activity ("{0}: {1}" -f $PackageID, ($Utils_OneDataCollector_Strings.ProgressActivity_DataCollectionGenericMessage)) -Status ("{0}: {1}" -f $Activity, ([System.IO.Path]::GetFileName($Filename)))
        
        <#
        switch($Activity.ToUpperInvariant())
        {
            'FILES'
            {
                Write-DiagProgress -Activity ($Utils_OneDataCollector_Strings.ProgressActivity_FilesCollection) -Status ("{0} ({1}): {2}" -f $PackageID, $Activity, ([System.IO.Path]::GetFileName($Filename)))
                break
            }
                
            'REGISTRYKEYS'
            {
                Write-DiagProgress -Activity ($Utils_OneDataCollector_Strings.ProgressActivity_RegistryKeysCollection) -Status ("{0} ({1}): {2}" -f $PackageID, $Activity, ([System.IO.Path]::GetFileName($Filename)))
                break
            }

            'EVENTLOGS'
            {
                Write-DiagProgress -Activity ($Utils_OneDataCollector_Strings.ProgressActivity_EventLogsCollection) -Status ("{0} ({1}): {2}" -f $PackageID, $Activity, ([System.IO.Path]::GetFileName($Filename)))
                break
            }

            'COMMANDS'
            {
                Write-DiagProgress -Activity ($Utils_OneDataCollector_Strings.ProgressActivity_CommandsCollection) -Status ("{0} ({1}): {2}" -f $PackageID, $Activity, ([System.IO.Path]::GetFileName($Filename)))
                break
            }
        }
        #>
    }
}

function Export-Report
{
    <#

    .DESCRIPTION
    Exports the given object as XML and (if specified) CSV.

    .PARAMETER InputObject
    The object which needs to be exported.
    
    .PARAMETER Name
    The name of the report
    
    .PARAMETER ExportDirectory
    The directory where the report has to be exported.
    
    .PARAMETER ExportAsCSV
    Setting this option will export the input object as CSV along with XML.
    
    .PARAMETER Overwrite
    Setting this option will overwrite the report files if already exits.
    
    .EXAMPLE
    Export-Report -InputObject $files -Name 'Files Collected' -ExportDirectory $packageDirectory -PackageID 'Apps' -ExportAsCSV -Overwrite

    .EXAMPLE
    Export-Report -I $files -N 'Files Collected' -D $packageDirectory -P 'Apps' -CSV -Force

    #>

    [CmdletBinding()]
    PARAM
    (
        [Alias('I')]
        [ValidateScript({$_ -ne $null})]
        [Parameter(Position = 1, Mandatory = $true)]
        [object] $InputObject,

        [Alias('N')]
        [parameter(Position = 2, Mandatory = $true)]
        [ValidateScript({-not([string]::IsNullOrEmpty($_)) -or ($_.Trim().Length -le 0)})]
        [string] $Name,

        [Alias('D')]
        [parameter(Position = 3, Mandatory = $true)]
        [ValidateScript({-not([string]::IsNullOrEmpty($_)) -or ($_.Trim().Length -le 0)})]
        [string] $ExportDirectory,

        [Alias('P')]
        [Parameter(Position = 4, Mandatory = $true)]
        [ValidateScript({-not([string]::IsNullOrEmpty($_)) -or ($_.Trim().Length -le 0)})]
        [string] $PackageID,

        [Alias('CSV')]
        [Parameter(Mandatory = $false)]
        [switch]$ExportAsCSV,
        
        [Alias('Force')]
        [Parameter(Mandatory = $false)]
        [switch]$Overwrite

    )
    PROCESS
    {
        if($InputObject)
        {
            $Name = [System.IO.Path]::GetFileNameWithoutExtension($Name)

            $ReportFilename = ([System.IO.Path]::Combine($ExportDirectory, ("{0}_{1}.XML" -f $PackageID, $Name)))

            if($Overwrite)
            {
                if(Test-Path($ReportFilename)){ Remove-Item $ReportFilename -Force -ErrorAction SilentlyContinue }
            }
            else
            {
                $ReportFilename = $ReportFilename | Get-NewFileNameIfExists
            }
            ($InputObject | ConvertTo-XML).Save($ReportFilename)

            if($ExportAsCSV)
            {
                $ReportFilename = ([System.IO.Path]::Combine($ExportDirectory, ("{0}_{1}.CSV" -f $PackageID, $Name)))
                
                if($Overwrite)
                {
                    if(Test-Path($ReportFilename)){ Remove-Item $ReportFilename -Force -ErrorAction SilentlyContinue }
                }
                else
                {
                    $ReportFilename = $ReportFilename | Get-NewFileNameIfExists
                }

                $InputObject | Export-Csv -Path $ReportFilename -Force -NoTypeInformation
            }
        }
    }
}

function Get-OSVersion
{
    $osVersion = [System.Environment]::OSVersion.Version;
    return [float]("{0}.{1}" -f $osVersion.Major, $osVersion.Minor)
}

function Get-ValidPath
{
    [CmdletBinding()]
    PARAM
    (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateScript({-not([string]::IsNullOrEmpty($_)) -or ($_.Trim().Length -le 0)})]
        [string] $Path
    )
    PROCESS 
    {
        foreach( $invalidChar in ([System.IO.Path]::GetInvalidPathChars()))
        {
            $Path = $Path.Replace($invalidChar, '#');
        }
        return $Path
    }
}

function Get-ValidFileName
{
    [CmdletBinding()]
    PARAM
    (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateScript({-not([string]::IsNullOrEmpty($_)) -or ($_.Trim().Length -le 0)})]
        [string] $FileName
    )
    PROCESS 
    {
        foreach( $invalidChar in ([System.IO.Path]::GetInvalidFileNameChars()))
        {
            $FileName = $FileName.Replace($invalidChar, '_');
        }
        return $FileName
    }
}

function Get-NewFileNameIfExists
{
    [CmdletBinding()]
    PARAM
    (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateScript({-not([string]::IsNullOrEmpty($_)) -or ($_.Trim().Length -le 0)})]
        [string] $FileName
    )
    PROCESS 
    {
        if(Test-Path($FileName))
        {
            $FileNameBackup = $FileName
            
            [int] $duplicateFilenameCounter = 0;
            while(Test-Path($FileName))
            {
                $duplicateFilenameCounter += 1
                $FileName = $FileNameBackup
                            
                if($FileName.LastIndexOf('.') -ne -1) #Files
                {
                    $FileName = $FileName.Insert($FileName.LastIndexOf('.'), (" ({0})" -f $duplicateFilenameCounter))
                }
                else #Directory/File without extension
                {
                    $FileName = "$FileName ({0})" -f $duplicateFilenameCounter
                }
            }
        }
        else
        {
            New-Item -ItemType Directory -Force -Path (Split-Path $FileName) | Out-Null
        }

        return $FileName
    }
}

#region Initializing Types 

function Initialize
{

    [string]$sourceCode = @"

namespace Microsoft.One.DataCollector
{
    using System;
    using System.IO;
    using System.Text;
    using System.Xml;
    using System.Xml.Serialization;
        
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.6.1055.0")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true, Namespace = "urn:Microsoft.One.DataCollector")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "urn:Microsoft.One.DataCollector", IsNullable = false)]
    public partial class DataPoints
    {

        private DataPointsPackage[] packageField;

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("Package")]
        public DataPointsPackage[] Package
        {
            get
            {
                return this.packageField;
            }
            set
            {
                this.packageField = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.6.1055.0")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true, Namespace = "urn:Microsoft.One.DataCollector")]
    public partial class DataPointsPackage
    {

        private DataPointsPackageCommands commandsField;

        private DataPointsPackageEventLogs eventLogsField;

        private DataPointsPackageFiles filesField;

        private DataPointsPackageRegistries registriesField;

        private string idField;

        /// <remarks/>
        public DataPointsPackageCommands Commands
        {
            get
            {
                return this.commandsField;
            }
            set
            {
                this.commandsField = value;
            }
        }

        /// <remarks/>
        public DataPointsPackageEventLogs EventLogs
        {
            get
            {
                return this.eventLogsField;
            }
            set
            {
                this.eventLogsField = value;
            }
        }

        /// <remarks/>
        public DataPointsPackageFiles Files
        {
            get
            {
                return this.filesField;
            }
            set
            {
                this.filesField = value;
            }
        }

        /// <remarks/>
        public DataPointsPackageRegistries Registries
        {
            get
            {
                return this.registriesField;
            }
            set
            {
                this.registriesField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string ID
        {
            get
            {
                return this.idField;
            }
            set
            {
                this.idField = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.6.1055.0")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true, Namespace = "urn:Microsoft.One.DataCollector")]
    public partial class DataPointsPackageCommands
    {

        private DataPointsPackageCommandsCommand[] commandField;

        private string[] textField;

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("Command")]
        public DataPointsPackageCommandsCommand[] Command
        {
            get
            {
                return this.commandField;
            }
            set
            {
                this.commandField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlTextAttribute()]
        public string[] Text
        {
            get
            {
                return this.textField;
            }
            set
            {
                this.textField = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.6.1055.0")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true, Namespace = "urn:Microsoft.One.DataCollector")]
    public partial class DataPointsPackageCommandsCommand
    {

        private string typeField;

        private string teamField;

        private string outputFileNameField;

        private string valueField;

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string Type
        {
            get
            {
                return this.typeField;
            }
            set
            {
                this.typeField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string Team
        {
            get
            {
                return this.teamField;
            }
            set
            {
                this.teamField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string OutputFileName
        {
            get
            {
                return this.outputFileNameField;
            }
            set
            {
                this.outputFileNameField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlTextAttribute()]
        public string Value
        {
            get
            {
                return this.valueField;
            }
            set
            {
                this.valueField = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.6.1055.0")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true, Namespace = "urn:Microsoft.One.DataCollector")]
    public partial class DataPointsPackageEventLogs
    {

        private DataPointsPackageEventLogsEventLog[] eventLogField;

        private string[] textField;

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("EventLog")]
        public DataPointsPackageEventLogsEventLog[] EventLog
        {
            get
            {
                return this.eventLogField;
            }
            set
            {
                this.eventLogField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlTextAttribute()]
        public string[] Text
        {
            get
            {
                return this.textField;
            }
            set
            {
                this.textField = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.6.1055.0")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true, Namespace = "urn:Microsoft.One.DataCollector")]
    public partial class DataPointsPackageEventLogsEventLog
    {

        private string teamField;

        private string valueField;

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string Team
        {
            get
            {
                return this.teamField;
            }
            set
            {
                this.teamField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlTextAttribute()]
        public string Value
        {
            get
            {
                return this.valueField;
            }
            set
            {
                this.valueField = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.6.1055.0")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true, Namespace = "urn:Microsoft.One.DataCollector")]
    public partial class DataPointsPackageFiles
    {

        private DataPointsPackageFilesFile[] fileField;

        private string[] textField;

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("File")]
        public DataPointsPackageFilesFile[] File
        {
            get
            {
                return this.fileField;
            }
            set
            {
                this.fileField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlTextAttribute()]
        public string[] Text
        {
            get
            {
                return this.textField;
            }
            set
            {
                this.textField = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.6.1055.0")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true, Namespace = "urn:Microsoft.One.DataCollector")]
    public partial class DataPointsPackageFilesFile
    {

        private string teamField;

        private string valueField;

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string Team
        {
            get
            {
                return this.teamField;
            }
            set
            {
                this.teamField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlTextAttribute()]
        public string Value
        {
            get
            {
                return this.valueField;
            }
            set
            {
                this.valueField = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.6.1055.0")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true, Namespace = "urn:Microsoft.One.DataCollector")]
    public partial class DataPointsPackageRegistries
    {

        private DataPointsPackageRegistriesRegistry[] registryField;

        private string[] textField;

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("Registry")]
        public DataPointsPackageRegistriesRegistry[] Registry
        {
            get
            {
                return this.registryField;
            }
            set
            {
                this.registryField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlTextAttribute()]
        public string[] Text
        {
            get
            {
                return this.textField;
            }
            set
            {
                this.textField = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.6.1055.0")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true, Namespace = "urn:Microsoft.One.DataCollector")]
    public partial class DataPointsPackageRegistriesRegistry
    {

        private string teamField;

        private string outputFileNameField;

        private string valueField;

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string Team
        {
            get
            {
                return this.teamField;
            }
            set
            {
                this.teamField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string OutputFileName
        {
            get
            {
                return this.outputFileNameField;
            }
            set
            {
                this.outputFileNameField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlTextAttribute()]
        public string Value
        {
            get
            {
                return this.valueField;
            }
            set
            {
                this.valueField = value;
            }
        }
    }

    /// <summary>
    /// Contains common functionalities used across the troubleshooter.
    /// </summary>
    public static class Utilities
    {
        /// <summary>
        /// Gets the data points from file.
        /// </summary>
        /// <param name="filename">The filename.</param>
        /// <param name="showError">if set to <c>true</c> [show error].</param>
        /// <returns></returns>
        public static DataPoints GetDataPointsFromFile(string filename, out bool isValid, out string errorMessage)
        {
            isValid = false;
            errorMessage = String.Empty;            

            try
            {
                if (!File.Exists(filename))
                {
                    return null;
                }

                using (TextReader reader = new StreamReader(filename))
                {
                    XmlSerializer serializer = new XmlSerializer(typeof(DataPoints));
                    var dataPoints = (DataPoints)serializer.Deserialize(reader);
                    
                    isValid = true;
                    return dataPoints;
                }
            }
            catch (InvalidOperationException ex)
            {
                isValid = false;
                errorMessage = String.Format("{0}{1}", ex.Message, (ex.InnerException != null ? String.Format("{0}{1}", Environment.NewLine, ex.InnerException.Message) : String.Empty));
                return null;
            }
            catch (Exception ex)
            {
                isValid = false;
                errorMessage = String.Format("{0}{1}", ex.Message, (ex.InnerException != null ? String.Format("{0}{1}", Environment.NewLine, ex.InnerException.Message) : String.Empty));
                return null;
            }
        }
		/// <summary>
        /// Gets the data points from XML String.
        /// </summary>
        /// <param name="filename">XML string source.</param>
        /// <param name="showError">if set to <c>true</c> [show error].</param>
        /// <returns></returns>

		public static DataPoints GetDataPointsFromString(string xmlSource, out bool isValid, out string errorMessage)
        {
            isValid = false;
            errorMessage = String.Empty;            

            try
            {
                StringReader strReader = new StringReader(xmlSource);
                XmlTextReader xmlread;
                xmlread = new XmlTextReader(strReader);
                
                XmlSerializer serializer = new XmlSerializer(typeof(DataPoints));
                var dataPoints = (DataPoints)serializer.Deserialize(xmlread);
                    
                isValid = true;
                return dataPoints;
            }
            catch (InvalidOperationException ex)
            {
                isValid = false;
                errorMessage = String.Format("{0}{1}", ex.Message, (ex.InnerException != null ? String.Format("{0}{1}", Environment.NewLine, ex.InnerException.Message) : String.Empty));
                return null;
            }
            catch (Exception ex)
            {
                isValid = false;
                errorMessage = String.Format("{0}{1}", ex.Message, (ex.InnerException != null ? String.Format("{0}{1}", Environment.NewLine, ex.InnerException.Message) : String.Empty));
                return null;
            }
        }

    }
}

"@

    if([float](Get-OSVersion) -le [float](6.1))
    {
        try
        { 
            $type = Add-Type -TypeDefinition $sourceCode -PassThru -Language CSharpVersion3 -ReferencedAssemblies System.Xml -ErrorAction Continue
        } 
        catch 
        {
            $_.Exception.Message | ConvertTo-Xml | Update-DiagReport -ID 'TS_Main' -Name 'Type Initialization Error' -Verbosity Informational			
        }
    }
    else
    {
        try
        { 
            $type = Add-Type -TypeDefinition $sourceCode -PassThru -ReferencedAssemblies System.Xml 
        }
        catch 
        {
            $_.Exception.Message | ConvertTo-Xml | Update-DiagReport -ID 'TS_Main' -Name 'Type Initialization Error' -Verbosity Informational			
        }
    }
}

Initialize # Initializing the types so below functionalities can be used.

#endregion

function Collect-Files
{
    [CmdletBinding()]
    PARAM
    (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Microsoft.One.DataCollector.DataPointsPackageFilesFile[]] $Files, 
        
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [string] $PackageID
    )
    PROCESS 
    {
        $filesCollected = @();
        $filesNotCollected = @();
        
        $currentActivity = 'Files'

        $packageDirectory = [System.IO.Path]::Combine($ResultRootDirectory, $PackageID)
        New-Item -ItemType Directory -Path $packageDirectory -Force | Out-Null
        foreach($file in $Files)
        {
			$filePath = $file.Value
            $sourceFilename = [System.Environment]::ExpandEnvironmentVariables($filePath)
			$sourceFilename = $sourceFilename -replace '"', ""
			if(Test-Path($sourceFilename))
			{
				try
				{
					$ErrorActionPreference = 'Stop'
                
					$resolvedFiles = Get-ChildItem -Path $sourceFilename
					foreach($resolvedFile in $resolvedFiles)
					{
						$teamName = ($file.Team, 'General')[([string]::IsNullOrEmpty($file.Team)) -or ($file.Team.Trim().Length -le 0)];
							
						$dstFileName = $env:COMPUTERNAME + "_" + ($resolvedFile.Name)

						$destinationFilename = (@($packageDirectory, $currentActivity, $teamName, $dstFileName) -join '\')| Get-ValidPath | Get-NewFileNameIfExists

						#Update-Progress -Activity $currentActivity -PackageID $PackageID -Filename $destinationFilename

						Copy-Item -Path ($resolvedFile.FullName) -Destination $destinationFilename -Force #-ErrorAction SilentlyContinue | Out-Null
					}
				}
				catch
				{
					Add-Member -InputObject $file -MemberType NoteProperty -Name Status -Value ($error[0].ToString()) -Force
					$filesNotCollected += $file;
				}

				Add-Member -InputObject $file -MemberType NoteProperty -Name Status -Value 'Collected' -Force
				$filesCollected += $file;
			}
			else
			{
				Add-Member -InputObject $file -MemberType NoteProperty -Name Status -Value 'The system cannot find the file specified.' -Force
				$filesNotCollected += $file;
			}
			
            
        }

        # Exporting list of files collected
        if($filesCollected)
        {
            Export-Report -InputObject $filesCollected -Name 'FilesCollected' -ExportDirectory $packageDirectory -PackageID $PackageID -CSV -Force
        }
        
        # Exporting list of files NOT collected
        if($filesNotCollected)
        {
            Export-Report -InputObject $filesNotCollected -Name 'FilesNotCollected' -ExportDirectory $packageDirectory -PackageID $PackageID -CSV -Force
        }
    }
}

function Collect-RegistryKeys
{
    [CmdletBinding()]
    PARAM
    (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Microsoft.One.DataCollector.DataPointsPackageRegistriesRegistry[]] $RegistryKeys, 
        
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [string] $PackageID
    )
    PROCESS 
    {
        $registryKeysCollected = @();
        $registryKeysNotCollected = @();
        
        $currentActivity = 'RegistryKeys'

        $packageDirectory = [System.IO.Path]::Combine($ResultRootDirectory, $PackageID)
        New-Item -ItemType Directory -Path $packageDirectory -Force | Out-Null

        foreach($registryKey in $RegistryKeys)
        {
            $registryKeyToExport = $registryKey.Value.Replace('\*', '')
            
            $teamName = ($registryKey.Team, 'General')[([string]::IsNullOrEmpty($registryKey.Team)) -or ($registryKey.Team.Trim().Length -le 0)];

            $outputFilename = ($registryKey.OutputFileName, ($registryKeyToExport | Get-ValidFileName))[([string]::IsNullOrEmpty($registryKey.OutputFileName)) -or ($registryKey.OutputFileName.Trim().Length -le 0)];
			$outputFilename = $env:COMPUTERNAME + "_" + $outputFilename 
            $outputFilename = (@($ResultRootDirectory, $PackageID, $currentActivity, $teamName, ("{0}.reg" -f [System.IO.Path]::GetFileNameWithoutExtension($outputFilename))) -join '\') | Get-ValidPath | Get-NewFileNameIfExists

            $registryKey.OutputFileName = [System.IO.Path]::GetFileName($outputFilename)

            try
            {	
                Update-Progress -Activity $currentActivity -PackageID $PackageID -Filename $outputFilename

                $ErrorActionPreference = 'Stop'

                [string] $result = REG.exe EXPORT ($registryKeyToExport) ($outputFilename) /y /reg:64 2>&1
                
                Add-Member -InputObject $registryKey -MemberType NoteProperty -Name Status -Value $result -Force
                $registryKeysCollected += $registryKey
            }
            catch [Exception]        
            {
                Add-Member -InputObject $registryKey -MemberType NoteProperty -Name Status -Value ($error[0].ToString()) -Force
                $registryKeysNotCollected += $registryKey
            }
        }

        # Exporting list of registry keys collected
        if($registryKeysCollected)
        {
            Export-Report -InputObject $registryKeysCollected -Name 'RegistryKeysCollected' -ExportDirectory $packageDirectory -PackageID $PackageID -CSV -Force
        }
        
        # Exporting list of registry keys NOT collected
        if($registryKeysNotCollected)
        {
            Export-Report -InputObject $registryKeysNotCollected -Name 'RegistryKeysNotCollected' -ExportDirectory $packageDirectory -PackageID $PackageID -CSV -Force
        }
    }
}

function Collect-EventLogs
{
    [CmdletBinding()]
    PARAM
    (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Microsoft.One.DataCollector.DataPointsPackageEventLogsEventLog[]] $EventLogs, 
        
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [string] $PackageID
    )
    PROCESS 
    {
        $eventLogsCollected = @();
        $eventLogsNotCollected = @();

        $currentActivity = 'EventLogs'

        $packageDirectory = [System.IO.Path]::Combine($ResultRootDirectory, $PackageID)
        New-Item -ItemType Directory -Path $packageDirectory -Force | Out-Null

        foreach($eventLog in $EventLogs)
        {
            $sourceEventLog = [System.Environment]::ExpandEnvironmentVariables($eventLog.Value);
            if(Test-Path($sourceEventLog))
            {
                try
                {
                    $ErrorActionPreference = 'Stop'

                    $resolvedEventLogs = Get-ChildItem -Path $sourceEventLog
                    foreach($resolvedEventLog in $resolvedEventLogs)
                    {
                        $teamName = ($eventLog.Team, 'General')[([string]::IsNullOrEmpty($eventLog.Team)) -or ($eventLog.Team.Trim().Length -le 0)];
                        $dstEventLog = $env:COMPUTERNAME + "_" + ($resolvedEventLog.Name)
                        $destinationEventLog = (@($packageDirectory, $currentActivity, $teamName, $dstEventLog ) -join '\') | Get-ValidPath | Get-NewFileNameIfExists

                        Update-Progress -Activity $currentActivity -PackageID $PackageID -Filename ($destinationEventLog)

                        Copy-Item -Path ($resolvedEventLog.FullName) -Destination $destinationEventLog -Force -ErrorAction SilentlyContinue | Out-Null
                    }
                }
                catch
                {
                    Add-Member -InputObject $eventLog -MemberType NoteProperty -Name Status -Value ($error[0].ToString()) -Force
                    $eventLogsNotCollected += $eventLog;
                }
                
                Add-Member -InputObject $eventLog -MemberType NoteProperty -Name Status -Value 'Collected' -Force
                $eventLogsCollected += $eventLog;
            }
            else
            {
                Add-Member -InputObject $eventLog -MemberType NoteProperty -Name Status -Value 'The system cannot find the specified event log.' -Force
                $eventLogsNotCollected += $eventLog;
            }
        }

        # Exporting list of event logs collected
        if($eventLogsCollected)
        {
            Export-Report -InputObject $eventLogsCollected -Name 'EventLogsCollected' -ExportDirectory $packageDirectory -PackageID $PackageID -CSV -Force
        }
        
        # Exporting list of event logs NOT collected
        if($eventLogsNotCollected)
        {
            Export-Report -InputObject $eventLogsNotCollected -Name 'EventLogsNotCollected' -ExportDirectory $packageDirectory -PackageID $PackageID -CSV -Force
        }
    }
}

function Collect-Commands
{
    [CmdletBinding()]
    PARAM
    (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Microsoft.One.DataCollector.DataPointsPackageCommandsCommand[]] $Commands, 
        
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [string] $PackageID
    )
    PROCESS 
    {
        $commandsCollected = @();
        $commandsNotCollected = @();

        $currentActivity = 'Commands'

        $packageDirectory = [System.IO.Path]::Combine($ResultRootDirectory, $PackageID)
        New-Item -ItemType Directory -Path $packageDirectory -Force | Out-Null

        foreach($command in $Commands)
        {
            $teamName = ($command.Team, 'General')[([string]::IsNullOrEmpty($command.Team)) -or ($command.Team.Trim().Length -le 0)];
			$fileName = $command.OutputFileName
			if(!($fileName -eq "NA"))
			{
				$fileName = "$env:COMPUTERNAME" + "_" + $filename
				$outputFilename = ($fileName, ([System.IO.Path]::GetRandomFileName()))[([string]::IsNullOrEmpty($fileName)) -or ($fileName.Trim().Length -le 0)];
				$outputFilename = (@($ResultRootDirectory, $PackageID, $currentActivity, $teamName, ("{0}.txt" -f [System.IO.Path]::GetFileNameWithoutExtension($outputFilename))) -join '\') | Get-ValidPath | Get-NewFileNameIfExists
            
				$command.OutputFileName = [System.IO.Path]::GetFileName($outputFilename)
			
				Update-Progress -Activity $currentActivity -PackageID $PackageID -Filename $outputFilename
			}
			else
			{
				$outputFilename = $null
			}
            if($command.Type)
            {
                switch($command.Type.ToUpperInvariant())
                {	
                    "PS"
                    {
                        try
                        {
							$ErrorActionPreference = 'Stop'	
							if($outputFilename)
							{
								(Invoke-Expression ($command.Value)) > $outputFilename
							}
							else
							{
								(Invoke-Expression ($command.Value))
							}
                            Add-Member -InputObject $command -MemberType NoteProperty -Name Status -Value 'Collected' -Force
                            $commandsCollected += $command
                        }
                        catch [Exception]
                        {
                            Add-Member -InputObject $command -MemberType NoteProperty -Name Status -Value ($error[0].ToString()) -Force
                            $commandsNotCollected += $command
                        }

                        break
                    }
                    "CMD"
                    {	
                        try
                        { 
                            $ErrorActionPreference = 'Stop'
                            (CMD.exe /c ($command.Value) 2>&1) > $outputFilename

                            Add-Member -InputObject $command -MemberType NoteProperty -Name Status -Value 'Collected' -Force
                            $commandsCollected += $command
                        }
                        catch [Exception]        
                        {             
                            Add-Member -InputObject $command -MemberType NoteProperty -Name Status -Value ($error[0].ToString()) -Force
                            $commandsNotCollected += $command
                        }

                        break
                    }				
                }
            }
            else
            {
                Add-Member -InputObject $command -MemberType NoteProperty -Name Status -Value 'Invalid Command Type'
                $commandsNotCollected += $command
            }
        }

        # Exporting list of commands collected
        if($commandsCollected)
        {
            Export-Report -InputObject $commandsCollected -Name 'CommandsCollected' -ExportDirectory $packageDirectory -PackageID $PackageID -CSV -Force
        }
        
        # Exporting list of commands NOT collected
        if($commandsNotCollected)
        {
            Export-Report -InputObject $commandsNotCollected -Name 'CommandsNotCollected' -ExportDirectory $packageDirectory -PackageID $PackageID -CSV -Force
        }
    }
}

function Get-Packages
{
    [CmdletBinding()]
    PARAM
    (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [string] $Filename
    )
    PROCESS 
    {
        $packages = "" | Select ValidPackages, InvalidPackages
        $packages.ValidPackages = @();
        $packages.InvalidPackages = @();
        if (([string]::IsNullOrEmpty($Filename)) -or ($Filename.Trim().Length -le 0))
        {
            Pop-Message -Message ($Utils_OneDataCollector_Strings.MessageBoxValue_BlankFilename) -Caption ($Utils_OneDataCollector_Strings.MessageBoxCaption_ImportError) -Type 64
            return $packages
        }
    
        if (-not (Test-Path($Filename)))
        {
            Pop-Message -Message ($Utils_OneDataCollector_Strings.MessageBoxValue_FileMissing) -Caption ($Utils_OneDataCollector_Strings.MessageBoxCaption_ImportError) -Type 48
            return $packages
        }

        # Getting XML files from Package
        $packageFilenames = @()
        if ($Filename.EndsWith('.CAB', [System.StringComparison]::OrdinalIgnoreCase))
        {
            # Extracting all files in the CAB
            $unpackedPackagePath = 'UnpackedPackage'
            
            if(Test-Path($unpackedPackagePath))
            {
                Remove-Item $unpackedPackagePath -Force -Recurse -ErrorAction SilentlyContinue
            }
            New-Item $unpackedPackagePath -Type Directory -Force | Out-Null
            
            $expandFiles = Expand.exe -r $filename -F:* $unpackedPackagePath

            # Moving attachments to the root so commands can use it.
            $attachments = Get-ChildItem $unpackedPackagePath | ? {$_.Extension -inotlike '.xml'}
			if($attachments)
			{
				$attachments | % { Move-Item -Path ($_.FullName) -Destination . -Force}
			}
            # Processing packages
            $packageFiles = Get-ChildItem $unpackedPackagePath | ? {$_.Extension -ilike '.xml'}
			
            foreach($packageFile in $packageFiles) 
            {
                $packageFilenames += $packageFile.FullName
            }
        }
        elseif ($Filename.EndsWith('.XML', [System.StringComparison]::OrdinalIgnoreCase))
        {
            $packageFilenames += $Filename
        }
        else
        {
            Pop-Message -Message ($Utils_OneDataCollector_Strings.MessageBoxValue_InvalidPackageType) -Caption ($Utils_OneDataCollector_Strings.MessageBoxCaption_ImportError) -Type 48
            return $packages
        }

        foreach($file in $packageFilenames)
        {
            # Getting data points from XML (package) files
            [bool] $isValid = $false;
            [string] $errorMessage = '';
            $dataPoints = [Microsoft.One.DataCollector.Utilities]::GetDataPointsFromFile($file, [ref] $isValid, [ref] $errorMessage)

            if($isValid)
            {
                foreach($package in ($dataPoints.Package))
                {
                    if($package -ne $null)
                    {
						$packageID = $package.ID
						$strLength = $packageID.length
						if($strLength -gt 0)
						{
							$packages.ValidPackages += $package
						}
						else
						{
							Pop-Message -Message ($Utils_OneDataCollector_Strings.MessageBoxValue_PackageID) -Caption ($Utils_OneDataCollector_Strings.MessageBoxCaption_ImportError) -Type 64
						}
                    }
                }
            }
            else
            {
                $invalidPackage = "" | Select FileName, ErrorMessage
                $invalidPackage.FileName = [System.IO.Path]::GetFileName($file)
                $invalidPackage.ErrorMessage = $errorMessage
                
                $packages.InvalidPackages += $invalidPackage
            }
        }
        return $packages
    }
}
Function Get-ResourceFromZipFile
{
	[CmdletBinding()]
    PARAM
    (
		[Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [string] $PackageID
    )

	$ResourceFileName = $PackageID + ".zip"
	if(Test-Path $ResourceFileName)
	{
		$UnzipPath = (Get-Location).Path
		$Shell = New-Object -com Shell.Application
		$Location = $Shell.NameSpace($UnzipPath)
		$ZipFiles = Get-Childitem $ResourceFileName -Recurse -Include *.ZIP
		$ZipFolder = $Shell.NameSpace($ZipFiles.fullname)
		$Location.Copyhere($ZipFolder.items(), 4) 
	}

} 

function Get-UserSelectedPackages
{
    [CmdletBinding()]
    PARAM
    (
		[Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [string] $PackageName
    )
    PROCESS 
    {
        $packages = "" | Select ValidPackages, InvalidPackages
        $packages.ValidPackages = @();
        $packages.InvalidPackages = @();
		if(![String]::IsNullOrEmpty($PackageName))
        {
			$packages = Get-Packages -Filename ($PackageName) 
			if($packages.InvalidPackages)
            {
                $errorMessages = @()
    
                ($packages.InvalidPackages) | %{$errorMessages += ("`r`n{0}: {1} `r`n{2}: {3}" -f ($Utils_OneDataCollector_Strings.MessageBoxValue_FilenameLabel), $_.FileName, ($Utils_OneDataCollector_Strings.MessageBoxValue_ErrorLabel), $_.ErrorMessage)}

                $errorMessage = "{0} `r`n{1}" -f ($Utils_OneDataCollector_Strings.MessageBoxValue_InvalidPackage), ($errorMessages -join "`r`n")
        
                Pop-Message -Message $errorMessage -Caption ($Utils_OneDataCollector_Strings.MessageBoxCaption_ImportError) -Type 48
			}
            foreach($package in ($packages.ValidPackages))
            {
                if($package)
                {
                    $notImported = $false
                    Process-Package -Package $package
                    .\RC_OneDataCollector.ps1 ($package.ID) ($package.GetType())
                }
            }
		}
		else
        {
            Pop-Message -Message ($Utils_OneDataCollector_Strings.MessageBoxValue_FileMissing) -Caption ($Utils_OneDataCollector_Strings.MessageBoxCaption_ImportError) -Type 64
        }
        
    }
}

function Get-PackageFromUserAndProcess
{
    [int]$count =  0
    [bool]$notImported = $true

    while($notImported -and ($count -lt 3))
    {
        $count++
        
        $input = "intune.xml" #Get-DiagInput -ID 'INT_FileLocation' -Parameter @{'descriptionValue' = ($Utils_OneDataCollector_Strings.INT_Desc_SourceLocation)}
        $packagePath = $input[0]
        if(![String]::IsNullOrEmpty($packagePath))
        {
            $packages = Get-Packages -Filename ($packagePath)

            if($packages.InvalidPackages)
            {
                $errorMessages = @()
    
                ($packages.InvalidPackages) | %{$errorMessages += ("`r`n{0}: {1} `r`n{2}: {3}" -f ($Utils_OneDataCollector_Strings.MessageBoxValue_FilenameLabel), $_.FileName, ($Utils_OneDataCollector_Strings.MessageBoxValue_ErrorLabel), $_.ErrorMessage)}

                $errorMessage = "{0} `r`n{1}" -f ($Utils_OneDataCollector_Strings.MessageBoxValue_InvalidPackage), ($errorMessages -join "`r`n")
        
                Pop-Message -Message $errorMessage -Caption ($Utils_OneDataCollector_Strings.MessageBoxCaption_ImportError) -Type 48
            }

            foreach($package in ($packages.ValidPackages))
            {
                if($package)
                {
                    $notImported = $false
                    Process-Package -Package $package
                    .\RC_OneDataCollector.ps1 ($package.ID) ($package.GetType())
                }
            }
        }
        else
        {
            Pop-Message -Message ($Utils_OneDataCollector_Strings.MessageBoxValue_FileMissing) -Caption ($Utils_OneDataCollector_Strings.MessageBoxCaption_ImportError) -Type 64
        }
    }

    if(-not $notImported)
    {
        $count =  0
    }
} 

function Process-Package
{
    [CmdletBinding()]
    PARAM
    (
        [Alias('P')]
        [ValidateScript({$_ -ne $null})]
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Microsoft.One.DataCollector.DataPointsPackage] $Package
    )
    PROCESS 
    {
		if(($Package.Commands) -ne $null -and ($Package.Commands.Command) -ne $null)
        {
            Collect-Commands -Commands ($Package.Commands.Command) -PackageID ($Package.ID)
        }
        if(($Package.Files) -ne $null -and ($Package.Files.File) -ne $null)
        {
            Collect-Files -Files ($Package.Files.File) -PackageID ($Package.ID)
        }
        if(($Package.Registries) -ne $null -and ($Package.Registries.Registry) -ne $null)
        {
            Collect-RegistryKeys -RegistryKeys ($Package.Registries.Registry) -PackageID ($Package.ID)
        }
        if(($Package.EventLogs) -ne $null -and ($Package.EventLogs.EventLog) -ne $null)
        {
            Collect-EventLogs -EventLogs ($Package.EventLogs.EventLog) -PackageID ($Package.ID)
        }
    }
}

function Compress-CollectedDataAndReport
{
    if(Test-Path($ResultRootDirectory))
    {
        Write-DiagProgress -Activity ($Utils_OneDataCollector_Strings.ProgressActivity_CompressingData)
        Create-ZipFromDirectory -Source $ResultRootDirectory -ZipFileName $CompressedResultFileName -Force #-Rooted
                
		Copy-Item -Path $ResultRootDirectory -Destination (get-location) -Force -Recurse -ErrorAction SilentlyContinue

		Remove-Item -Path $ResultRootDirectory -Force -Recurse -ErrorAction SilentlyContinue
    }
}

#endregion

function Write-DiagProgress {
    Param (
        $activity,
        $status

    )

    Write-Output "$status"
    $status | Out-File $env:systemroot\temp\stdout.log -Append -Force

}


Initialize # Initializing the types so below functionalities can be used.


Function Test-IsAdmin
{
    ([Security.Principal.WindowsPrincipal] `
      [Security.Principal.WindowsIdentity]::GetCurrent() `
    ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

#Region stand-alone

Initialize # Initializing the types so below functionalities can be used.
if (-not (Test-IsAdmin) ) {
    Return "Please run PowerShell elevated (run as administrator) and run the script again."
    Break
    }
$ResultRootDirectory = [System.IO.Path]::Combine(($env:TEMP), 'CollectedData')
$CompressedResultFileName = "CollectedData.ZIP"
$xmlPath =  join-path $pwd "Intune.XML"
# we assume that the XML and .ps1 are in the same folder
$scriptDir = Split-Path $script:MyInvocation.MyCommand.Path
cd $scriptDir



if (-not (Test-Path $xmlPath) ) {
    try {
       $downloadLocation = "https://raw.githubusercontent.com/markstan/IntuneOneDataCollector/master/Intune.xml"
      
       if ( [System.Net.WebProxy]::GetDefaultProxy().Address -eq $null ) {
                Invoke-WebRequest -Uri $downLoadLocation -OutFile .\Intune.XML
                }
            else {
    
                $myproxy = ([System.Net.WebProxy]::GetDefaultProxy().Address.AbsoluteURI)    
                Invoke-WebRequest -Uri $downLoadLocation -OutFile .\Intune.XML -Proxy $myproxy
            }
            
    }

    catch {
        $message = "Unable to download Intune.XML.  Exiting"
        Write-Output $message
        $message | Out-File $env:systemroot\temp\stdout.log -Append -Force
        Pop-Message -Caption "Unable to download Intune.xml" -message "Unable to download Intune.XML.  Please download from http://aka.ms/IntuneXML and run the script again" -Type 48
    }
}

$package = get-packages -Filename $xmlPath
Process-Package -Package $package.ValidPackages[0]
Compress-CollectedDataAndReport
if (Test-Path .\collecteddata.zip) {
    rd -Path .\CollectedData -Recurse -Force -ErrorAction SilentlyContinue
}
start .

