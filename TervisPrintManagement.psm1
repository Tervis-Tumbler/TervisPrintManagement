function Add-PrinterMetadataMember {
    param(
        [Switch]$PassThrough,
        [Parameter(ValueFromPipeline, Mandatory)]$Printer
    )
    process {
        $PrinterMetadata = try {$Printer.comment | convertfrom-json} catch {}
        foreach ($Property in $PrinterMetadata.psobject.Properties) {
            $Printer | Add-Member -MemberType NoteProperty -Name $($Property.Name) -Value $($Property.Value)
        }

        if($PassThrough) { $Printer }
    }
}

function Add-TervisPrinterCustomProperites {
    param(
        [Switch]$PassThrough,
        [Parameter(ValueFromPipeline, Mandatory)]$Printer
    )
    process {
        $Printer | Add-Member -MemberType ScriptProperty -Name PageCount -Value {
            Get-PrinterPageCount -Name $this.Name
        }

        $Printer | Add-PrinterMetadataMember

        if($PassThrough) { $Printer }
    }
}


Function Get-TervisPrinter {
    Get-Printer -ComputerName disney |
    Add-TervisPrinterCustomProperites -PassThrough |
    Where-Object DeviceType -eq Print
}

Function Set-TervisPrinterMetadataMember {
    param (
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]
        $Name,
        
        [ValidateSet("HP","Zebra","Kyocera","Brother","Konica","Evolis","Epson","Highlight Technologies Inc")]
        $Vendor,

        $Model,
        
        [ValidateSet("Gulf Business Systems","Tervis")]
        $ServicedBy
    )
    $Paramaters = $PSBoundParameters

    $MetaDataProperties = $Paramaters | 
    ConvertFrom-PSBoundParameters | 
    select -Property * -ExcludeProperty Name

    Set-Printer -ComputerName Disney -Name $Name -Comment $($MetaDataProperties | ConvertTo-JSON)
}

Function Get-TervisPrinterLifeTimePageCount {

    $Printers = Get-TervisPrinter
    
    $GBSPrinters = $Printers |
    Where-Object ServicedBy -EQ "Gulf Business Systems"

    $SNMP = new-object -ComObject olePrn.OleSNMP

    foreach ($Printer in $GBSPrinters| where Vendor -eq HP) {
        $Printer.Name
        $snmp.open($Printer.Name,"public",2,3000)
        $snmp.gettree("43.11.1.1.8.1")
        $printertype = $snmp.Get(".1.3.6.1.2.1.25.3.2.1.3.1")
        $printertype
        $snmp.get(".1.3.6.1.4.1.11.2.3.9.4.2.1.4.1.2.6")
        #$snmp.get(".1.3.6.1.2.1.43.10.2.1.4.1.1")  

    }

    $snmp.open("10.172.28.13","public",2,3000)
    $snmp.gettree("43.11.1.1.8.1")
    $printertype = $snmp.Get(".1.3.6.1.2.1.25.3.2.1.3.1")

    $snmp.open("10.172.24.116","public",2,3000)
    $snmp.gettree("43.11.1.1.8.1")
    $printertype = $snmp.Get(".1.3.6.1.2.1.25.3.2.1.3.1")
    $snmp.get(".1.3.6.1.2.1.43.10.2.1.4.1.1")  
}

function Get-PrinterSNMPProperties {
    param (
        [Parameter(Mandatory)]$Name
    )
    $SNMP = new-object -ComObject olePrn.OleSNMP
    $SNMP.open($Name,"public",2,3000)
    $Result = $SNMP.gettree("printmib")
    ConvertFrom-TwoDimensionalArray -Array $Result
    $SNMP.Close()
}

function Get-PrinterPageCount {
    param (
        [Parameter(Mandatory)]$Name
    )
    if (Test-Connection -ComputerName $Name -Count 1 -BufferSize 16 -Delay 1 -quiet -ErrorAction SilentlyContinue) {
        $PrinterProperties = Get-PrinterSNMPProperties -Name $Name
        $PrinterProperties."printmib.prtMarker.prtMarkerTable.prtMarkerEntry.prtMarkerLifeCount.1.1"
    }
}

function Get-GBSPrintCounts {
    Get-TervisPrinter | 
    where ServicedBy -eq "Gulf Business Systems" | 
    select Name, PageCount
}

function Update-PrinterLocation {
    param(
        [parameter(Mandatory = $true)]$PrinterToUpdate,
        [parameter(Mandatory = $true)]$NewPrinterLocation,
        $PrintServer = "Disney.tervis.prv"
    )

    Set-Printer -Name $PrinterToUpdate -ComputerName $PrintServer -Location $NewPrinterLocation
}

function Set-ZebraDriversToPackaged {
    Set-Location -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Print\Environments\Windows x64\Drivers\Version-3'    
    Get-ChildItem -Name Z* | 
        ForEach-Object {    
            $Driver = Get-ItemProperty -Path $_ -Name PrinterDriverAttributes
            if ($Driver.PrinterDriverAttributes -eq 0) {        
                Write-Host "$($Driver.PSChildName) is not packaged. Setting status to `"packaged.`""
                Set-ItemProperty -Path $Driver.PSChildName -Name PrinterDriverAttributes -Value 1
            }    
        }
}

function Set-AllPrinterDriversToPackaged {
    # PrinterDriverAttributes is defined by DRIVER_INFO_8 struct: https://msdn.microsoft.com/en-us/library/windows/desktop/dd162507(v=vs.85).aspx
    param (
        [Parameter(Mandatory=$true)]$ComputerName
    )
    
    Invoke-Command -ComputerName $ComputerName -ScriptBlock {
        Set-Location -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Environments\Windows x64\Drivers\Version-3"

        $PrinterDriversNotSetAsPackage = Get-ChildItem | 
            Get-ItemProperty -Name PrinterDriverAttributes,DriverVersion | 
            where {($_.PrinterDriverAttributes -band 1) -eq 0} | 
            select PSChildName,DriverVersion,PrinterDriverAttributes

        foreach ($PrinterDriver in $PrinterDriversNotSetAsPackage) {
            $NewPrinterDriverAttributeValue = $PrinterDriver.PrinterDriverAttributes -bor 1
            Set-ItemProperty -Path $PrinterDriver.PSChildName -Name PrinterDriverAttributes -Value $NewPrinterDriverAttributeValue
            New-Object -TypeName PSObject -Property @{
                Name = $PrinterDriver.PSChildName
                DriverVersion = $PrinterDriver.DriverVersion
                OldPrinterDriverAttributes = $PrinterDriver.PrinterDriverAttributes
                NewPrinterDriverAttributes = $NewPrinterDriverAttributeValue
            }
        }
    } | select -Property Name,DriverVersion,OldPrinterDriverAttributes,NewPrinterDriverAttributes 
}
