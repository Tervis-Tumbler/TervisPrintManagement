﻿function Add-PrinterMetadataMember {
    param(
        [Switch]$PassThrough,
        [Parameter(ValueFromPipeline, Mandatory)]$Printer
    )
    process {
        $PrinterMetadata = try {$Printer.comment | convertfrom-json} catch {}
        foreach ($Property in $PrinterMetadata.psobject.Properties) {
            $Printer | Add-Member -MemberType NoteProperty -Name $($Property.Name) -Value $($Property.Value) -Force
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
        $Printer | Add-Member -MemberType ScriptProperty -Force -Name PageCount -Value {
            if ($this.PrinterStatus -notmatch "Offline") {
                if ($this.Model -eq "TASKalfa 500ci") {
                    $Result = Invoke-WebRequest -Uri "http://$($This.Name)/start/StatCntFunc.htm"
                    
                    ($Result.content -split "`r`n") | 
                    Where-Object { $_ -match "var monochrome_total = " } |
                    ForEach-Object {$_ -replace "var ", "$"} | 
                    Invoke-Expression
    
                    $monochrome_total
                } else {
                    Get-PrinterPageCount -Name $this.Name
                }    
            }
        }

        $Printer | Add-PrinterMetadataMember

        if($PassThrough) { $Printer }
    }
}


Function Get-TervisPrinter {
    param (
        $Name
    )
    $PrintServers = "disney","INF-PrintSrv01","INF-PrintSrv02"

    $PrintServers |
    ForEach-Object {Get-Printer -ComputerName $_} |
    Where-Object { -not $Name -or $_.Name -eq $Name} |
    Where-Object DeviceType -eq Print | 
    Add-TervisPrinterCustomProperites -PassThrough
}

Function Set-TervisPrinterMetadataMember {
    param (
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]
        $Name,
        
        [ValidateSet("HP","Zebra","Kyocera","Brother","Konica","Evolis","Epson","Highlight Technologies Inc")]
        $Vendor,

        $Model,
        
        [ValidateSet("Gulf Business Systems","Tervis")]
        $ServicedBy,
        
        $WiredMacAddress,
        $WirelessMacAddress,
        [ValidateSet("Direct-Thermal","Thermal-Transfer")]$MediaType,
        $LabelWidth,
        $LabelHeight,
        [ValidateSet(203,300,600)]$DPI
    )
    process {
        Write-Progress -Activity "Set-TervisPrinterMetadataMember" -CurrentOperation "Adding metadata to $Name"
        $MetaDataProperties = $PSBoundParameters | 
        ConvertFrom-PSBoundParameters -ExcludeProperty Name
    
        $MetaDataPropertyNames = $MetaDataProperties | Get-PropertyName
    
        $ExistingMetatDataPropertiesToInclude = Get-TervisPrinter -Name $Name | 
        Select-Object -ExpandProperty Comment |
        ConvertFrom-Json |
        Select-Object -Property * -ExcludeProperty $MetaDataPropertyNames
        
        $CombinedMetaDataProperties = $MetaDataProperties, $ExistingMetatDataPropertiesToInclude | Merge-Object
        
        Set-Printer -ComputerName Disney -Name $Name -Comment $($CombinedMetaDataProperties | ConvertTo-JSON)
    }
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
    param (
        [Switch]$DisableWakeupPrints
    )

    $GBSPrinters = Get-TervisPrinter | 
    where ServicedBy -eq "Gulf Business Systems"

    if (-Not $DisableWakeupPrints) {
        Start-ParallelWork -Parameters (
            $GBSPrinters |
            Where-Object PrinterStatus -NotMatch Offline |
            Select-Object -ExpandProperty Name
        ) -ScriptBlock {
            param (
                $PrinterName
            )
            try {
                Send-TCPClientData -ComputerName $PrinterName -Port 9100 -Data "Wake up printer text" -NoReply
            } catch {
                Write-Warning "No response from $PrinterName"
            }
        }
    }

    $GBSPrinters |
    Select-Object -Property Name, PageCount, Model, PrinterStatus
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
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$ComputerName
    )
    process {
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
}

function Invoke-FixBravesNotPrinting {
    $Service = get-service -Name lpdsvc -ComputerName disney
    if ($Service.status -ne "Running") {
        $Service | Restart-Service
    }
    
    $Service = get-service -Name lpdsvc -ComputerName disney
    if ($Service.status -ne "Running") {
        Throw "Tried to start lpdsvc but the service is still not running"
    }
}

function Add-TervisPrinter {
    param(
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]$Name,
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]$DriverName,
        [Parameter(Mandatory)]$ComputerName,
        [Switch]$Shared,
        [Switch]$Force
    )
    process {
        Add-PrinterDriver -Name $DriverName -ComputerName $ComputerName
        if ($Force) {
            Remove-Printer -Name $Name -ComputerName $ComputerName -ErrorAction SilentlyContinue
            Remove-PrinterPort -Name $Name -ComputerName $ComputerName -ErrorAction SilentlyContinue
        }        
        Add-PrinterPort -Name $Name -PrinterHostAddress $Name -ComputerName $ComputerName -ErrorAction SilentlyContinue
        if ($Shared) {
            Add-Printer -PortName $Name -Name $Name -DriverName $DriverName -ComputerName $ComputerName -ErrorAction SilentlyContinue -Shared -ShareName $Name
        } else {
            Add-Printer -PortName $Name -Name $Name -DriverName $DriverName -ComputerName $ComputerName -ErrorAction SilentlyContinue
        }
    }
}

function Remove-TervisPrinter {
    param(
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]$Name,
        [Parameter(Mandatory)]$ComputerName
    )
    process {
        if ($ComputerName) {
            Remove-Printer -Name $Name -ComputerName $ComputerName 
            Remove-PrinterPort -Name $Name -ComputerName $ComputerName
        } else {
            Remove-Printer -Name $Name
            Remove-PrinterPort -Name $Name
        }
    }
}

function Get-TervisZebraPrinterPropertiesFromConfigPage {
    param (
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$Name
    )
    process {
        $URL = "http://$Name/config.html"
        try {
            $ConfigPage = Invoke-WebRequest $URL        
        } catch {
            throw "Could not get to config page for $Name"
        }

        $PreTag = $ConfigPage.ParsedHtml.body.all | Where-Object tagname -eq "PRE"
        
        $Properties = $PreTag.innerText -split "`n" | 
            ForEach-Object { $_.Trim() } |
            Select-Object -Skip 1
        
        $PrintMethod = $Properties |
            Where-Object {$_ -like "*PRINT METHOD"} |
            Split-String " " |
            Select-Object -First 1

        switch ($PrintMethod) {
            "DIRECT-THERMAL" { $MediaType = "Direct-Thermal"; break }
            "THERMAL-TRANS." { $MediaType = "Thermal-Transfer"; break }
            Default { throw "No media type for $Name" }
        }

        [int]$LabelWidth = $Properties |
            Where-Object {$_ -like "*PRINT WIDTH"} |
            Split-String " " |
            Select-Object -First 1
 
        [int]$LabelHeight = $Properties |
            Where-Object {$_ -like "*LABEL LENGTH"} |
            Split-String " " |
            Select-Object -First 1

        return [PSCustomObject]@{
            Name = $Name
            MediaType = $MediaType
            LabelWidth = $LabelWidth
            LabelHeight = $LabelHeight
        }
    }
}
