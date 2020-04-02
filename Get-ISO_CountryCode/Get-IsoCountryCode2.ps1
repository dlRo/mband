Function FindISOCountryCode {

    param ([string]$countryName)

    # Example Bing Maps Key - paste your key here
    $bingMapsKey = 'EmgyDsPmANmh1oIN4VKtA4ajr2KsW6wesLCd5dbRsrtPCGDI2GkvjG4iMFL2nEm9'

    $restData = invoke-restmethod -uri "https://dev.virtualearth.net/REST/v1/Locations?q=$countryName&incl=ciso2&key=$bingMapsKey"
    
    If ($restData.resourcesets.resources.address.countryRegionIso2) 
    {
        Return $restData.resourcesets.resources.address.countryRegionIso2
    }
    Else
    {
        Return $false
    }
}

FindISOCountryCode 'Australia'
