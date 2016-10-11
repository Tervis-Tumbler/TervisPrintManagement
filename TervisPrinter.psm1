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


Function Get-TervisPrinter {
    Get-Printer -ComputerName disney |
    Add-PrinterMetadataMember -PassThrough
}

Function Get-TervisPrinterLifeTimePageCount {
$SNMP = new-object -ComObject olePrn.OleSNMP

$snmp.open("10.172.28.13","public",2,3000)
$snmp.gettree("43.11.1.1.8.1")
$printertype = $snmp.Get(".1.3.6.1.2.1.25.3.2.1.3.1")

$snmp.open("10.172.24.116","public",2,3000)
$snmp.gettree("43.11.1.1.8.1")
$printertype = $snmp.Get(".1.3.6.1.2.1.25.3.2.1.3.1")
$snmp.get(".1.3.6.1.2.1.43.10.2.1.4.1.1")  
}