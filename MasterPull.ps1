#in this file all the azure resources reservations are done on are called, and the corresponding properties needed in regards to reservations are collected, and parsed

function loadSQLObjects{
$SQLObjects = Search-AzGraph -Query '
resources
| where type =~ "microsoft.sql/servers/elasticpools"
| extend skuName = tostring(sku.name)
| extend skuCapacity = sku.capacity
| extend Kind = kind
| project name, Kind, location, subscriptionId, skuName, skuCapacity
| union (
    resources
    | where type =~ "microsoft.sql/servers/databases"
    | extend Kind = kind
    | extend skuName = tostring(sku.name)
    | extend skuCapacity = sku.capacity
    | project name,  subscriptionId, Kind, location, skuName, skuCapacity
    | where skuCapacity != 0
)
|union(
    resources
    | where type =~ "microsoft.sql/managedinstances"
    | extend skuName = strcat("MI_", tostring(sku.name))
    | extend skuCapacity = sku.capacity
    | extend Kind = "Managed Instance"
    | project name, subscriptionId, Kind, location, skuName, skuCapacity
)
|extend type = "SqlDatabases"
'

Write-Host "SQL Objects loaded"
return $SQLObjects
}


function loadAppServiceObjects{
$AppServiceObjects = Search-AzGraph -Query '
resources
| where type == "microsoft.web/serverfarms"
| extend apps = properties.numberOfSites
| extend appServiceEnvironmentId = properties.hostingEnvironmentId
| extend skuCapacity = sku.capacity
| extend Kind = kind
| extend skuName = strcat(iff(Kind == "app", "Windows_", "Linux_"), tostring(sku.name))
| where skuName contains "V3"
| extend type = "AppService"
| project name, subscriptionId, Kind, location, skuName, skuCapacity, type
'

Write-Host "AppService Objects loaded"
return $AppServiceObjects
}

function loadvmssObjects{
$vmssObjects = Search-AzGraph -Query '
resources
| where type == "microsoft.compute/virtualmachinescalesets"
| extend type == tostring(type)
| extend type == "ScaleSet"
| extend skuName = sku.name
| extend skuCapacity = sku.capacity
| project name, subscriptionId, kind, location, skuName, skuCapacity, type
'

Write-Host "ScaleSets Objects loaded"
return $vmssObjects
}


function createObjectCSV($ObjectsParam){
    $AzureItems = $ObjectsParam | ForEach-Object{
        foreach ($Object in $_){
            [PSCustomObject]@{
                name= $Object.name
                subscriptionId=$Object.subscriptionId
                kind= $Object.Kind
                location= $Object.location
                skuName= $Object.skuName
                capacity= $Object.skuCapacity
                Type=$Object.type
                DatePulled = Get-Date -Format "MM/dd/yyyy"
            };
        }
    }
        $AzureItems | Export-Csv -Path "$PSScriptRoot\Objects.csv"
        Write-Host "Objects CSV has been uploaded"
}


##======================================================== Reservation Script===================================================================================================

function createReservationsCSV{
    # Get all reservations
    $reservations = Get-AzReservation
    # Filter active and expiring reservations
    $filteredReservations = $reservations | Where-Object { ($_.DisplayProvisioningState -eq 'Succeeded' -or $_.DisplayProvisioningState -eq 'Expiring')}

    Write-Host "Reservations Records have been loaded"

    #Create a custom object with the desired columns
    $reportData = $filteredReservations | ForEach-Object{
        [PSCustomObject]@{
            Name = $_.DisplayName
            ReservationID = $_.Name
            ExperationDate = $_.expiryDate.ToString("yyyy-MM-dd")
            PurchaseDate = $_.PurchaseDate.ToString("yyyy-MM-dd") 
            Term = $_.Term
            Scope = $_.AppliedScopePropertyDisplayName
            Location = $_.Location
            Type = $_.ReservedResourceType
            Quantity = $_.Quantity
            Description = $_.SkuDescription
            ProductName = $_.SkuName
            UtilizationTrend = $_.UtilizationTrend
        };
    }
    # Export to Excel
    $reportData | Export-Csv -Path "$PSScriptRoot\ReservationsReport.csv" -NoTypeInformation

    Write-Host "Reservations CSV is uploaded"
}