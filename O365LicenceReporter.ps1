$Path = "C:\O365Reports\"


try{
    Import-Module MSOnline
} catch {
    Install-Module MSOnline
}
Connect-MsolService

$PartConts = Get-MsolPartnerContract 

Foreach ($Partner in $PartConts){
    $EndHTML = ""
    $tid = $Partner.tenantid
    $partname = $Partner.Name
    $filepath = $Path + $partname + "_License_Report.html"
    
    $LiceSubs = Get-MsolAccountSku -TenantId $tid
    $prodlist = Import-Csv .\ProductReference.csv
    $NiceTotal = @()
    
    Foreach ($licesub in $LiceSubs){
        $prodref = $licesub.SkuId
        $ProdNiceName = $prodlist | ? {$prodref -contains $_.GUID}
        $LiceUsers = Get-Msoluser -tenantid $tid | Where-Object {($_.licenses).AccountSkuID -match ($ProdNiceName).'String ID'} | Select-Object DisplayName, UserPrincipalName
        If ($ProdNiceName.'Product Name'){
        $CleanSub = [pscustomobject][ordered]@{
            Product = $ProdNiceName.'Product name'
            UnitsTotal = $licesub.ActiveUnits
            UnitsConsumed = $licesub.ConsumedUnits
            Users = $LiceUsers
            }
        $NiceTotal += $CleanSub
        }
    }

    $EndHTML += $NiceTotal| Select-Object Product, UnitsTotal, UnitsConsumed | convertto-html -PreContent "<h1> Total Count </h1>" -Fragment
    

    Foreach ($product in $Nicetotal){
        $EndHTML += $Product.Users | ConvertTo-Html -Fragment -PreContent "<h2>$($product.product) - Users:</h2>" 

    }

    $endHTML| Out-File $filepath

}









