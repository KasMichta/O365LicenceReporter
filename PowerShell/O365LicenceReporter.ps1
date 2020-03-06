$Path = "C:\O365Reports\"

$TableHeader = "<table class=`"table table-bordered table-hover`" style=`"width:80%`">"
$Whitespace = "<br/>"
$TableStyling = "<th>", "<th style=`"background-color:#0d81bb`">"
$prodlist = Import-Csv .\ProductReference.csv
$colourlist = @("DarkYellow", "DarkRed", "DarkGreen", "DarkCyan")

try {
    Import-Module MSOnline 
}
catch {
    Install-Module MSOnline
}
Connect-MsolService

$PartConts = Get-MsolPartnerContract 
$coloi = 0

Foreach ($Partner in $PartConts) {
    
    # Check if color is on last item ? Reset : Ignore
    If ($coloi -eq 4) {
        $coloi = 0
    }

    # Set Colour for Org
    $OrgCol = $colourlist[$coloi]

    $EndHTML = ""
    $tid = $Partner.tenantid
    $partname = $Partner.Name
    $filepath = $Path + $partname + "_License_Report.html"

    $LiceSubs = Get-MsolAccountSku -TenantId $tid
    $NiceTotal = @()
    Write-Host "Gathering licence information for" -NoNewline
    Write-Host "$partname" -ForegroundColor $OrgCol
    
    $coltoggle = $false

    Foreach ($licesub in $LiceSubs) {

        If ($coltoggle -eq $false) {
            $licecol = "DarkMagenta"
        }
        else {
            $licecol = "DarkGray"
        }

        $prodref = $licesub.SkuId
        $ProdNiceName = $prodlist | ? { $prodref -contains $_.GUID }

        If ($ProdNiceName) {
            Write-Host "$($partname): " -ForegroundColor $OrgCol -NoNewline
            Write-Host "$($ProdNiceName.'Product name'): " -ForegroundColor $licecol -NoNewline
            Write-Host "Found Licence" -ForegroundColor Green
            $LiceUsers = Get-Msoluser -tenantid $tid | Where-Object { ($_.licenses).AccountSkuID -match ($ProdNiceName).'String ID' } | Select-Object DisplayName, UserPrincipalName
            Write-Host "$($partname): " -ForegroundColor $OrgCol -NoNewline
            Write-Host "$($ProdNiceName.'Product name'): " -ForegroundColor $licecol -NoNewline
            Write-Host "Found $(($LiceUsers | Measure-Object).Count) Users" -ForegroundColor Yellow
            $AvaiLice = $licesub.ActiveUnits - $licesub.ConsumedUnits
            If ($AvaiLice -gt 0) {
                Write-Host "$($partname): " -ForegroundColor $OrgCol -NoNewline
                Write-Host "$($ProdNiceName.'Product name'): " -ForegroundColor $licecol -NoNewline
                Write-Host "Found $AvaiLice Available/Unused Licences" -ForegroundColor Yellow
            }
            If ($ProdNiceName.'Product Name') {
                $CleanSub = [pscustomobject][ordered]@{
                    Product       = $ProdNiceName.'Product name'
                    UnitsTotal    = $licesub.ActiveUnits
                    UnitsConsumed = $licesub.ConsumedUnits
                    Users         = $LiceUsers
                }
                $NiceTotal += $CleanSub
            }
            $coltoggle = -not $coltoggle
        }
    }

    $coloi += 1

   <#  $TotalHTMLRAW = $NiceTotal | Select-Object Product, UnitsTotal, UnitsConsumed | convertto-html -Fragment | Select-Object -Skip 1
    $EndHTML += $TableHeader + ($TotalHTMLRAW -replace $TableStyling) + $Whitespace

    Foreach ($product in $Nicetotal) {
        $EndHTML += $Product.Users | ConvertTo-Html -PreContent "<h2>$($product.product) - Users:</h2>" -Head $Header

    }

    $endHTML | Out-File $filepath #>

}









