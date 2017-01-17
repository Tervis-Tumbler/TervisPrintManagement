Describe "Printer Driver Validation" {
    #$PrintServer = "disney"
    #$Session = New-PSSession -ComputerName $PrintServer
    $PrinterDriverRegKeyPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Environments\Windows x64\Drivers\Version-3"

    $PrinterDrivers = Get-ChildItem $PrinterDriverRegKeyPath | select -ExpandProperty PSChildName    

    foreach ($PrinterDriver in $PrinterDrivers) {
        $FullPathToPrinterDriver = Join-Path -Path $PrinterDriverRegKeyPath -ChildPath $PrinterDriver
        $PrinterDriverProperties = Get-ItemProperty -Path $FullPathToPrinterDriver -Name PrinterDriverAttributes
        $PackagedStatus = $PrinterDriverProperties.PrinterDriverAttributes -band 1      
        It "$PrinterDriver should be set to Packaged" {
            $PackagedStatus | Should be 1
        }
    }
}