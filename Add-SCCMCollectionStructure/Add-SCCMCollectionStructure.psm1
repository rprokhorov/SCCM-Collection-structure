
# main function
function Add-SCCMCollectionStructure(
    [Parameter(Mandatory = $true)]
    [string]$f_AppName,
    [Parameter(Mandatory = $true)]
    [array]$f_sites,
    [Parameter(Mandatory = $false)]
    [array]$sites_VIP,
    [Parameter(Mandatory = $true)]
    [ValidateSet("Standard","Not standard")]
    [string]$AppType,
    [Parameter(Mandatory = $false)]
    [ValidateSet("Available","Required")]
    [string]$VIPInstallType = "Available",
    [Parameter(Mandatory = $true)]
    [ValidateSet("Add new application","Add new domain to existing collections")] 
    [string]$Action = "Add new application",
    [Parameter(Mandatory = $true)]
    [string]$psite,
    [Parameter(Mandatory = $false)]
    $Schedule
)
{
    begin{
        Write-Host 'Function: Add-SCCMCollectionStructure' -ForegroundColor Magenta
        Write-Host "Import module for SCCM" -ForegroundColor Green
        Write-Host "Import-Module `"$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1`"" -ForegroundColor Yellow
        Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1"
        Set-Location $psite":"
        if (!$Schedule)
        {
            $Schedule = New-CMSchedule -Start "01/01/2014 9:00 PM" -DayOfWeek Saturday -RecurCount 1
        }
        # Create core folders
        if (!(Test-Path $psite":\DeviceCollection\Inventory"))
        {
            Write-Host "Create folder $psite`:\DeviceCollection\Inventory" -ForegroundColor Green
            Write-Host "New-Item  $psite`":\DeviceCollection\Inventory`"" -ForegroundColor Yellow
            New-Item  $psite":\DeviceCollection\Inventory"

        }
        if (!(Test-Path $psite":\DeviceCollection\Inventory\Applications"))
        {
            Write-Host "Create folder $psite`:\DeviceCollection\Inventory\Applications" -ForegroundColor Green
            Write-Host "New-Item  $psite`":\DeviceCollection\Inventory\Applications`"" -ForegroundColor Yellow
            New-Item  $psite":\DeviceCollection\Inventory\Applications"
        }
        if (!(Test-Path $psite":\DeviceCollection\_Exclude"))
        {
            Write-Host "Create folder $psite`:\DeviceCollection\_Exclude" -ForegroundColor Green
            Write-Host "New-Item  $psite`":\DeviceCollection\_Exclude`"" -ForegroundColor Yellow
            New-Item  $psite":\DeviceCollection\_Exclude"
        }
        if (!(Test-Path $psite":\DeviceCollection\Deploy Apps"))
        {
            Write-Host "Create folder $psite`:\DeviceCollection\Deploy Apps" -ForegroundColor Green
            Write-Host "New-Item  $psite`":\DeviceCollection\Deploy Apps`"" -ForegroundColor Yellow
            New-Item  $psite":\DeviceCollection\Deploy Apps"
        }
        # Create core collections
        if (!(Get-CMCollection -Name 'ALL WKS'))
        {
            Write-Host "Create collections 'ALL WKS' in folder '$psite`:\DeviceCollection'" -ForegroundColor Green
            Write-Host "New-CMDeviceCollection  -Name `"ALL WKS`" -LimitingCollectionName `"All Systems`" -Comment `"All workstations`" -RefreshSchedule `$Schedule | Out-Null" -ForegroundColor Yellow
            New-CMDeviceCollection  -Name "ALL WKS"`
                                -LimitingCollectionName "All Systems"`
                                -Comment "All workstations"`
                                -RefreshSchedule $Schedule #| Out-Null
        }

        if (!(Get-CMCollection -Name 'Exclude | DA | _Install Apps'))
        {
            Write-Host "Create collections 'Exclude | DA | _Install Apps' in folder '$psite`:\DeviceCollection\_Exclude'" -ForegroundColor Green
            Write-Host "New-CMDeviceCollection  -Name `"Exclude | DA | _Install Apps`" -LimitingCollectionName `"All WKS`" -Comment `"`" -RefreshSchedule `$Schedule | Move-CMObject -FolderPath $psite`":\DeviceCollection\_Exclude`"" -ForegroundColor Yellow
            New-CMDeviceCollection  -Name "Exclude | DA | _Install Apps"`
                                -LimitingCollectionName "All WKS"`
                                -Comment ""`
                                -RefreshSchedule $Schedule | Move-CMObject -FolderPath $psite":\DeviceCollection\_Exclude"
        }
        
        if (!(Get-CMCollection -Name 'DEP+VIP | ALL'))
        {
            Write-Host "Create collections 'DEP+VIP | ALL' in folder '$psite`:\DeviceCollection\_Exclude'" -ForegroundColor Green
            Write-Host "New-CMDeviceCollection  -Name `"DEP+VIP | ALL`" -LimitingCollectionName `"All WKS`" -Comment `"`" -RefreshSchedule `$Schedule | Move-CMObject -FolderPath $psite`":\DeviceCollection\_Exclude`"" -ForegroundColor Yellow
            New-CMDeviceCollection  -Name "DEP+VIP | ALL"`
                                -LimitingCollectionName "All WKS"`
                                -Comment ""`
                                -RefreshSchedule $Schedule | Move-CMObject -FolderPath $psite":\DeviceCollection\_Exclude"
        }

        foreach ($f_site in $f_sites)
        {
            if (!(Get-CMCollection -Name "All WKS | $f_site"))
            {
                Write-Host "Create collections 'All WKS | $f_site' in folder '$psite`:\DeviceCollection'" -ForegroundColor Green
                Write-Host "New-CMDeviceCollection  -Name `"All WKS | $f_site`" -LimitingCollectionName `"All WKS`" -Comment `"All workstations in domain $f_sites`" -RefreshSchedule `$Schedule | Out-Null" -ForegroundColor Yellow
                New-CMDeviceCollection  -Name "All WKS | $f_site"`
                                    -LimitingCollectionName "All WKS"`
                                    -Comment "All workstations in domain $f_sites"`
                                    -RefreshSchedule $Schedule | Out-Null
            }
            
            if (!(Get-CMCollection -Name "DEP+VIP | $f_site | ALL"))
            {
                Write-Host "Create collections 'DEP+VIP | $f_site | ALL' in folder '$psite`:\DeviceCollection\_Exclude'" -ForegroundColor Green
                Write-Host "New-CMDeviceCollection  -Name `"DEP+VIP | $f_site | ALL`" -LimitingCollectionName `"All WKS | $f_site`" -Comment `"All workstations in domain $f_sites`" -RefreshSchedule `$Schedule | Move-CMObject -FolderPath $psite`":\DeviceCollection\_Exclude`"" -ForegroundColor Yellow
                New-CMDeviceCollection  -Name "DEP+VIP | $f_site | ALL"`
                                    -LimitingCollectionName "All WKS | $f_site"`
                                    -Comment "All workstations in domain $f_sites"`
                                    -RefreshSchedule $Schedule | Move-CMObject -FolderPath $psite":\DeviceCollection\_Exclude"
            }
        }
    }

    process{
    
    if ($Action -eq "Add new application")
    {
        Write-Host "Create folder '$f_AppName' in '$psite`:\DeviceCollection\Deploy Apps'" -ForegroundColor Green
        Write-Host "New-Item -Name `"$f_AppName`" -Path $psite`:\DeviceCollection\Deploy Apps" -ForegroundColor Yellow
        New-Item -Name "$f_AppName" -Path $psite":\DeviceCollection\Deploy Apps"
        Write-Host "Create folder '$f_AppName' in '$psite`:\DeviceCollection\Deploy Apps'" -ForegroundColor Green
        Write-Host "New-Item -Name `"$f_AppName`" -Path $psite`:\DeviceCollection\Deploy Apps" -ForegroundColor Yellow
        New-Item -Name "$f_AppName" -Path $psite":\DeviceCollection\Inventory\Applications"
        Write-Host "Create-CollectionForInventoryApplication -f_AppName $f_AppName -f_sites $f_sites -Action $Action -Schedule $($Schedule.StartTime)" -ForegroundColor Yellow
        Create-CollectionForInventoryApplication -f_AppName $f_AppName -f_sites $f_sites -Action $Action -Schedule $Schedule
        Write-Host "Create-CollectionForExclude -f_AppName $f_AppName -Action $Action -Schedule `$Schedule" -ForegroundColor Yellow
        Create-CollectionForExclude -f_AppName $f_AppName -Action $Action -Schedule $Schedule
        Write-Host "Create-CollectionForMigration -f_AppName $f_AppName -AppType $AppType -Action $Action -Schedule `$Schedule" -ForegroundColor Yellow
        Create-CollectionForMigration -f_AppName $f_AppName -AppType $AppType -Action $Action -Schedule $Schedule
        Write-Host "Create-CollectionForDeployApplication -f_AppName $f_AppName -f_sites $f_sites -AppType $AppType -Action $Action -Schedule `$Schedule" -ForegroundColor Yellow
        Create-CollectionForDeployApplication -f_AppName $f_AppName -f_sites $f_sites -AppType $AppType -Action $Action -Schedule $Schedule
        Write-Host "Create-VIPCollectionForMigration -f_AppName $f_AppName -AppType $AppType -VIPInstallType $VIPInstallType -Action $Action -Schedule `$Schedule" -ForegroundColor Yellow
        Create-VIPCollectionForMigration -f_AppName $f_AppName -AppType $AppType -VIPInstallType $VIPInstallType -Action $Action -Schedule $Schedule
        if ($sites_VIP)
        {
            Write-Host "Create-VIPCollectionForDeployApplication -f_AppName $f_AppName -f_sites $f_sites -AppType $AppType -VIPInstallType $VIPInstallType -Schedule `$Schedule -sites_VIP $sites_VIP -Action $Action" -ForegroundColor Yellow
            Create-VIPCollectionForDeployApplication -f_AppName $f_AppName -f_sites $f_sites -AppType $AppType -VIPInstallType $VIPInstallType -Schedule $Schedule -sites_VIP $sites_VIP -Action $Action
        }
        else {
            Write-Host "Create-VIPCollectionForDeployApplication -f_AppName $f_AppName -f_sites $f_sites -AppType $AppType -VIPInstallType $VIPInstallType -Schedule `$Schedule -Action $Action" -ForegroundColor Yellow
            Create-VIPCollectionForDeployApplication -f_AppName $f_AppName -f_sites $f_sites -AppType $AppType -VIPInstallType $VIPInstallType -Schedule $Schedule -Action $Action
        }
    }
    elseif ($Action -eq "Add new domain to existing collections") {
        Write-host "Add new collection for domain to existing hierarchy"
        #New-Item -Name "$f_AppName" -Path $psite":\DeviceCollection\Deploy Apps"
        #New-Item -Name "$f_AppName" -Path $psite":\DeviceCollection\Inventory\Applications"
        Write-Host "Create-CollectionForInventoryApplication -f_AppName $f_AppName -f_sites $f_sites -Action $Action -Schedule `$Schedule" -ForegroundColor Yellow
        Create-CollectionForInventoryApplication -f_AppName $f_AppName -f_sites $f_sites -Action $Action -Schedule $Schedule
        Write-Host  "Create-CollectionForExclude -f_AppName $f_AppName -Action $Action -Schedule `$Schedule" -ForegroundColor Yellow
        Create-CollectionForExclude -f_AppName $f_AppName -Action $Action -Schedule $Schedule
        Write-Host "Create-CollectionForMigration -f_AppName $f_AppName -AppType $AppType -Action $Action -Schedule `$Schedule" -ForegroundColor Yellow
        Create-CollectionForMigration -f_AppName $f_AppName -AppType $AppType -Action $Action -Schedule $Schedule
        Write-Host "Create-CollectionForDeployApplication -f_AppName $f_AppName -f_sites $f_sites -AppType $AppType -Action $Action -Schedule `$Schedule" -ForegroundColor Yellow
        Create-CollectionForDeployApplication -f_AppName $f_AppName -f_sites $f_sites -AppType $AppType -Action $Action -Schedule $Schedule
        Write-Host "Create-VIPCollectionForMigration -f_AppName $f_AppName -AppType $AppType -VIPInstallType $VIPInstallType -Action $Action -Schedule `$Schedule" -ForegroundColor Yellow
        Create-VIPCollectionForMigration -f_AppName $f_AppName -AppType $AppType -VIPInstallType $VIPInstallType -Action $Action -Schedule $Schedule
        if ($sites_VIP)
        {
            Write-Host "Create-VIPCollectionForDeployApplication -f_AppName $f_AppName -f_sites $f_sites -AppType $AppType -VIPInstallType $VIPInstallType -Schedule `$Schedule -sites_VIP $sites_VIP -Action $Action"  -ForegroundColor Yellow
            Create-VIPCollectionForDeployApplication -f_AppName $f_AppName -f_sites $f_sites -AppType $AppType -VIPInstallType $VIPInstallType -Schedule $Schedule -sites_VIP $sites_VIP -Action $Action
        }
        else {
            Write-Host "Create-VIPCollectionForDeployApplication -f_AppName $f_AppName -f_sites $f_sites -AppType $AppType -VIPInstallType $VIPInstallType -Schedule `$Schedule -Action $Action" -ForegroundColor Yellow
            Create-VIPCollectionForDeployApplication -f_AppName $f_AppName -f_sites $f_sites -AppType $AppType -VIPInstallType $VIPInstallType -Schedule $Schedule -Action $Action
        }
    }
    }
} 

# Create Incentory
function Create-CollectionForInventoryApplication(
    [Parameter(Mandatory = $true)]
    [string]$f_AppName,
    [Parameter(Mandatory = $true)]
    [array]$f_sites,
    [Parameter(Mandatory = $true)]
    [string]$Action,
    [Parameter(Mandatory = $true)]
    $Schedule
)
{
    Write-Host 'Function: Create-CollectionForInventoryApplication' -ForegroundColor Magenta
    Set-Location $psite":"
    #Write-Host $(Get-Location)
    if ($Action -eq "Add new application")
    {
        # Создаём основную коллекцию для All WKS
        Write-Host "Create collection 'INV | All WKS | $f_AppName'" -ForegroundColor Green
        Write-Host "Add-CMDeviceCollectionQueryMembershipRule   -CollectionName ""INV | All WKS | $f_AppName"" -RuleName ""Query All workstations with $f_AppName"" -QueryExpression ""select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System inner join SMS_G_System_Installed_software on SMS_R_System.ResourceID = SMS_G_System_Installed_software.ResourceID where SMS_G_System_Installed_software.Productname like """"%$f_AppName%""""""" -ForegroundColor Yellow
        New-CMDeviceCollection  -Name "INV | All WKS | $f_AppName"`
                                -LimitingCollectionName "All WKS"`
                                -Comment "Компьютеры с ПО $f_AppName"`
                                -RefreshSchedule $Schedule | Move-CMObject -FolderPath $psite":\DeviceCollection\Inventory\Applications\$f_AppName"
        Write-Host "Add-CMDeviceCollectionQueryMembershipRule   -CollectionName `"INV | All WKS | $f_AppName`" -RuleName `"Query All workstations with $f_AppName`" -QueryExpression `"select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System inner join SMS_G_System_Installed_software on SMS_R_System.ResourceID = SMS_G_System_Installed_software.ResourceID where SMS_G_System_Installed_software.Productname like `"%$f_AppName%`"" -ForegroundColor Yellow
        Add-CMDeviceCollectionQueryMembershipRule   -CollectionName "INV | All WKS | $f_AppName"`
                                                    -RuleName "Query All workstations with $f_AppName"`
                                                    -QueryExpression "select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System inner join SMS_G_System_Installed_software on SMS_R_System.ResourceID = SMS_G_System_Installed_software.ResourceID where SMS_G_System_Installed_software.Productname like ""%$f_AppName%"""
    }
    # Создаём коллекции для каждой площадки
    ForEach ($f_site in $f_sites)
    {
        Write-Host "Create collection 'INV | All WKS | $f_site | $f_AppName'" -ForegroundColor Green
        Write-Host "New-CMDeviceCollection  -Name ""INV | All WKS | $f_site | $f_AppName"" -LimitingCollectionName ""INV | All WKS | $f_AppName"" -Comment ""Компьютеры с ПО $f_AppName в домене $f_site"" -RefreshSchedule `$Schedule | Move-CMObject -FolderPath $psite"":\DeviceCollection\Inventory\Applications\$f_AppName""" -ForegroundColor Yellow
        New-CMDeviceCollection  -Name "INV | All WKS | $f_site | $f_AppName"`
                                -LimitingCollectionName "INV | All WKS | $f_AppName"`
                                -Comment "Компьютеры с ПО $f_AppName в домене $f_site"`
                                -RefreshSchedule $Schedule | Move-CMObject -FolderPath $psite":\DeviceCollection\Inventory\Applications\$f_AppName"
        Write-Host "Add-CMDeviceCollectionIncludeMembershipRule -CollectionName ""INV | All WKS | $f_site | $f_AppName"" -IncludeCollectionName ""All WKS | $f_site"""  -ForegroundColor Yellow
        Add-CMDeviceCollectionIncludeMembershipRule -CollectionName "INV | All WKS | $f_site | $f_AppName"`
                            -IncludeCollectionName "All WKS | $f_site"
    }
}

function Create-CollectionForExclude(
    # Название приложения
    [Parameter(Mandatory = $true)]
    [string]$f_AppName,
    [Parameter(Mandatory = $true)]
    [string]$Action,
    [Parameter(Mandatory = $true)]
    $Schedule
    )  
{
    Write-Host 'Function: Create-CollectionForExclude' -ForegroundColor Magenta
    Set-Location $psite":"
    if ($Action -eq "Add new application")
    {
        Write-Host "Create collection 'Exclude | DA | $f_AppName'" -ForegroundColor Green
        Write-Host "New-CMDeviceCollection  -Name `"Exclude | DA | $f_AppName`" -LimitingCollectionName `"All WKS`" -Comment `"Компьютеры, на которые не ставим $f_AppName (prod версия)`" -RefreshSchedule `$Schedule | Move-CMObject -FolderPath $psite`":\DeviceCollection\_Exclude`"" -ForegroundColor Yellow
        New-CMDeviceCollection  -Name "Exclude | DA | $f_AppName"`
                                -LimitingCollectionName "All WKS"`
                                -Comment "Компьютеры, на которые не ставим $f_AppName (prod версия)"`
                                -RefreshSchedule $Schedule | Move-CMObject -FolderPath $psite":\DeviceCollection\_Exclude"
        Write-Host "Add-CMDeviceCollectionIncludeMembershipRule -CollectionName `"Exclude | DA | $f_AppName`" -IncludeCollectionName `"Exclude | DA | _Install Apps`"" -ForegroundColor Yellow
        Add-CMDeviceCollectionIncludeMembershipRule -CollectionName "Exclude | DA | $f_AppName"`
                                -IncludeCollectionName "Exclude | DA | _Install Apps"
    }
}

function Create-CollectionForMigration(
    # Название приложения
    [Parameter(Mandatory = $true)]
    [string]$f_AppName,
    [Parameter(Mandatory = $true)]
    [string]$AppType,
    [Parameter(Mandatory = $true)]
    [string]$Action,
    [Parameter(Mandatory = $true)]
    $Schedule
)
{
    Write-Host 'Function: Create-CollectionForMigration' -ForegroundColor Magenta
    Set-Location $psite":"
    if ($Action -eq "Add new application")
    {
        if ($AppType -eq "Standard")
        {
            Write-Host "Create collection 'DA | All WKS | $f_AppName | Migration | Required'" -ForegroundColor Green
            Write-Host "New-CMDeviceCollection  -Name `"DA | All WKS | $f_AppName | Migration | Required`" -LimitingCollectionName `"All WKS`" -Comment `"Коллекция для развёртывания тестовой версии ПО $f_AppName`" -RefreshSchedule `$Schedule | Move-CMObject -FolderPath $psite`":\DeviceCollection\Deploy Apps\$f_AppName`"" -ForegroundColor Yellow
            New-CMDeviceCollection  -Name "DA | All WKS | $f_AppName | Migration | Required"`
                                    -LimitingCollectionName "All WKS"`
                                    -Comment "Коллекция для развёртывания тестовой версии ПО $f_AppName"`
                                    -RefreshSchedule $Schedule | Move-CMObject -FolderPath $psite":\DeviceCollection\Deploy Apps\$f_AppName"
            Write-Host "Add-CMDeviceCollectionExcludeMembershipRule -CollectionName `"DA | All WKS | $f_AppName | Migration | Required`" -ExcludeCollectionName `"Exclude | DA | _Install Apps`"" -ForegroundColor Yellow
            Add-CMDeviceCollectionExcludeMembershipRule -CollectionName "DA | All WKS | $f_AppName | Migration | Required"`
                                    -ExcludeCollectionName "Exclude | DA | _Install Apps"
        }
        if ($AppType -eq "Not standard")
        {
            Write-Host "Create collection 'DA | All WKS | $f_AppName | Migration | Required'" -ForegroundColor Green
            Write-Host "New-CMDeviceCollection  -Name `"DA | All WKS | $f_AppName | Migration | Required`" -LimitingCollectionName `"INV | All WKS | $f_AppName`" -Comment `"Коллекция для развёртывания тестовой версии ПО $f_AppName`" -RefreshSchedule `$Schedule | Move-CMObject -FolderPath $psite`":\DeviceCollection\Deploy Apps\$f_AppName`"" -ForegroundColor Yellow
            New-CMDeviceCollection  -Name "DA | All WKS | $f_AppName | Migration | Required"`
                                    -LimitingCollectionName "INV | All WKS | $f_AppName"`
                                    -Comment "Коллекция для развёртывания тестовой версии ПО $f_AppName"`
                                    -RefreshSchedule $Schedule | Move-CMObject -FolderPath $psite":\DeviceCollection\Deploy Apps\$f_AppName"
            Write-Host "Add-CMDeviceCollectionExcludeMembershipRule -CollectionName `"DA | All WKS | $f_AppName | Migration | Required`" -ExcludeCollectionName `"Exclude | DA | _Install Apps`"" -ForegroundColor Yellow
            Add-CMDeviceCollectionExcludeMembershipRule -CollectionName "DA | All WKS | $f_AppName | Migration | Required"`
                                    -ExcludeCollectionName "Exclude | DA | _Install Apps"
        }
    }
}

function Create-VIPCollectionForMigration(
    # Название приложения
    [Parameter(Mandatory = $true)]
    [string]$f_AppName,
    [Parameter(Mandatory = $true)]
    [string]$AppType,
    [Parameter(Mandatory = $true)]
    [ValidateSet("Available","Required")]
    [string]$VIPInstallType,
    [Parameter(Mandatory = $true)]
    [string]$Action,
    [Parameter(Mandatory = $true)]
    $Schedule
)
{
    Write-Host 'Function: Create-CollectionForMigration' -ForegroundColor Magenta
    Set-Location $psite":"
    if ($action -eq "Add new application")
    {
        if ($AppType -eq "Standard")
        {
            Write-Host "Create collection 'DA | VIP | All WKS | $f_AppName | Migration | $VIPInstallType'" -ForegroundColor Green
            Write-Host "New-CMDeviceCollection  -Name `"DA | VIP | All WKS | $f_AppName | Migration | $VIPInstallType`" -LimitingCollectionName `"All WKS`" -Comment `"Коллекция для развёртывания тестовой версии ПО $f_AppName`" -RefreshSchedule `$Schedule | Move-CMObject -FolderPath $psite`":\DeviceCollection\Deploy Apps\$f_AppName`"" -ForegroundColor Yellow
            New-CMDeviceCollection  -Name "DA | VIP | All WKS | $f_AppName | Migration | $VIPInstallType"`
                                    -LimitingCollectionName "All WKS"`
                                    -Comment "Коллекция для развёртывания тестовой версии ПО $f_AppName"`
                                    -RefreshSchedule $Schedule | Move-CMObject -FolderPath $psite":\DeviceCollection\Deploy Apps\$f_AppName"
            Write-Host "Add-CMDeviceCollectionExcludeMembershipRule -CollectionName `"DA | VIP | All WKS | $f_AppName | Migration | $VIPInstallType`" -ExcludeCollectionName `"Exclude | DA | _Install Apps`"" -ForegroundColor Yellow
            Add-CMDeviceCollectionExcludeMembershipRule -CollectionName "DA | VIP | All WKS | $f_AppName | Migration | $VIPInstallType"`
                                    -ExcludeCollectionName "Exclude | DA | _Install Apps"
        }
        if ($AppType -eq "Not standard")
        {
            Write-Host "Create collection 'DA | VIP | All WKS | $f_AppName | Migration | $VIPInstallType'" -ForegroundColor Green
            Write-Host "New-CMDeviceCollection  -Name `"DA | VIP | All WKS | $f_AppName | Migration | $VIPInstallType`" -LimitingCollectionName `"INV | All WKS | $f_AppName`" -Comment `"Коллекция для развёртывания тестовой версии ПО $f_AppName`" -RefreshSchedule `$Schedule | Move-CMObject -FolderPath $psite`":\DeviceCollection\Deploy Apps\$f_AppName`"" -ForegroundColor Yellow
            New-CMDeviceCollection  -Name "DA | VIP | All WKS | $f_AppName | Migration | $VIPInstallType"`
                                    -LimitingCollectionName "INV | All WKS | $f_AppName"`
                                    -Comment "Коллекция для развёртывания тестовой версии ПО $f_AppName"`
                                    -RefreshSchedule $Schedule | Move-CMObject -FolderPath $psite":\DeviceCollection\Deploy Apps\$f_AppName"
            Write-Host "Add-CMDeviceCollectionExcludeMembershipRule -CollectionName `"DA | VIP | All WKS | $f_AppName | Migration | $VIPInstallType`" -ExcludeCollectionName `"Exclude | DA | _Install Apps`"" -ForegroundColor Yellow
            Add-CMDeviceCollectionExcludeMembershipRule -CollectionName "DA | VIP | All WKS | $f_AppName | Migration | $VIPInstallType"`
                                    -ExcludeCollectionName "Exclude | DA | _Install Apps"
        }
    }
}


# для всех типов приложений
# Создаём коллекции для распространения приложений
function Create-CollectionForDeployApplication (
    [Parameter(Mandatory = $true)]
    [string]$f_AppName,
    [Parameter(Mandatory = $true)]
    [array]$f_sites,
    [Parameter(Mandatory = $true)]
    [ValidateSet("Standard","Not standard")]
    [string]$AppType,
    [Parameter(Mandatory = $true)]
    [string]$Action,
    [Parameter(Mandatory = $true)]
    $Schedule
)
{
    Write-Host 'Function: Create-CollectionForDeployApplication' -ForegroundColor Magenta
    Set-Location $psite":"
    $collections_exclude = "Exclude | DA | _Install Apps", "Exclude | DA | $f_AppName", "DEP+VIP | ALL", "DA | All WKS | $f_AppName | Migration | Required"
    
    if ($AppType -eq "Not standard")
    {
        if ($Action -eq "Add new application")
        {
            # Создаём основную коллекцию для All WKS
            Write-Host "Create collection 'DA | All WKS | $f_AppName | Prod | Required'" -ForegroundColor Green
            Write-Host "New-CMDeviceCollection  -Name `"DA | All WKS | $f_AppName | Prod | Required`" -LimitingCollectionName `"INV | All WKS | $f_AppName`" -Comment `"Коллекция для установки `"продуктивной`" версии ПО $f_AppName на все рабочие станции.`" -RefreshSchedule `$Schedule | Move-CMObject -FolderPath $psite`":\DeviceCollection\Deploy Apps\$f_AppName`"" -ForegroundColor Yellow
            New-CMDeviceCollection  -Name "DA | All WKS | $f_AppName | Prod | Required"`
                                    -LimitingCollectionName "INV | All WKS | $f_AppName"`
                                    -Comment "Коллекция для установки ""продуктивной"" версии ПО $f_AppName на все рабочие станции."`
                                    -RefreshSchedule $Schedule | Move-CMObject -FolderPath $psite":\DeviceCollection\Deploy Apps\$f_AppName"
            Write-Host "Add-CMDeviceCollectionIncludeMembershipRule -CollectionName `"DA | All WKS | $f_AppName | Prod | Required`" -IncludeCollectionName `"All WKS`"" -ForegroundColor Yellow
            Add-CMDeviceCollectionIncludeMembershipRule -CollectionName "DA | All WKS | $f_AppName | Prod | Required"`
                                                        -IncludeCollectionName "All WKS"

            foreach ($collection_exclude in $collections_exclude)
            {
                Write-Host "Add-CMDeviceCollectionExcludeMembershipRule -CollectionName ""DA | All WKS | $f_AppName | Prod | Required"" -ExcludeCollectionName ""$collection_exclude""" -ForegroundColor Yellow
                Add-CMDeviceCollectionExcludeMembershipRule -CollectionName "DA | All WKS | $f_AppName | Prod | Required"`
                                                            -ExcludeCollectionName $collection_exclude
            }
        }
        # Создаём коллекции для каждой площадки
        ForEach ($f_site in $f_sites)
        {
            Write-Host "Create collection 'DA | All WKS | $f_site | $f_AppName | Prod | Required'" -ForegroundColor Green
            $collections_exclude = "Exclude | DA | _Install Apps", "Exclude | DA | $f_AppName", "DEP+VIP | $f_site | ALL", "DA | All WKS | $f_AppName | Migration | Required"
            Write-Host "New-CMDeviceCollection  -Name `"DA | All WKS | $f_site | $f_AppName | Prod | Required`" -LimitingCollectionName `"INV | All WKS | $f_site | $f_AppName`" -Comment `"Коллекция для установки `"продуктивной`" версии ПО $f_AppName на все рабочие станции в домене $f_site`" -RefreshSchedule `$Schedule | Move-CMObject -FolderPath $psite`":\DeviceCollection\Deploy Apps\$f_AppName`"" -ForegroundColor Yellow
            New-CMDeviceCollection  -Name "DA | All WKS | $f_site | $f_AppName | Prod | Required"`
                                    -LimitingCollectionName "INV | All WKS | $f_site | $f_AppName"`
                                    -Comment "Коллекция для установки ""продуктивной"" версии ПО $f_AppName на все рабочие станции в домене $f_site"`
                                    -RefreshSchedule $Schedule | Move-CMObject -FolderPath $psite":\DeviceCollection\Deploy Apps\$f_AppName"
            
            Write-Host "Add-CMDeviceCollectionIncludeMembershipRule -CollectionName ""DA | All WKS | $f_site | $f_AppName | Prod | Required"" -IncludeCollectionName `"All WKS | $f_site`"" -ForegroundColor Yellow
            Add-CMDeviceCollectionIncludeMembershipRule -CollectionName "DA | All WKS | $f_site | $f_AppName | Prod | Required"`
                                    -IncludeCollectionName "All WKS | $f_site"
            foreach ($collection_exclude in $collections_exclude)
            {
                Write-Host "Add-CMDeviceCollectionExcludeMembershipRule -CollectionName ""DA | All WKS | $f_site | $f_AppName | Prod | Required"" -ExcludeCollectionName `"$collection_exclude`"" -ForegroundColor Yellow
                Add-CMDeviceCollectionExcludeMembershipRule -CollectionName "DA | All WKS | $f_site | $f_AppName | Prod | Required"`
                                                            -ExcludeCollectionName $collection_exclude
            }
        }
    }
    if ($AppType -eq "Standard")
    {
        if ($Action -eq "Add new application")
        {
            # Создаём основную коллекцию для All WKS
            Write-Host "Create collection 'DA | All WKS | $f_AppName | Prod | Required'" -ForegroundColor Green
            Write-Host "New-CMDeviceCollection  -Name `"DA | All WKS | $f_AppName | Prod | Required`" -LimitingCollectionName `"All WKS`" -Comment `"Коллекция для установки `"продуктивной`" версии ПО $f_AppName на все рабочие станции.`" -RefreshSchedule `$Schedule | Move-CMObject -FolderPath $psite`":\DeviceCollection\Deploy Apps\$f_AppName`"" -ForegroundColor Yellow
            New-CMDeviceCollection  -Name "DA | All WKS | $f_AppName | Prod | Required"`
                                    -LimitingCollectionName "All WKS"`
                                    -Comment "Коллекция для установки ""продуктивной"" версии ПО $f_AppName на все рабочие станции."`
                                    -RefreshSchedule $Schedule | Move-CMObject -FolderPath $psite":\DeviceCollection\Deploy Apps\$f_AppName"
            Write-Host "Add-CMDeviceCollectionIncludeMembershipRule -CollectionName `"DA | All WKS | $f_AppName | Prod | Required`" -IncludeCollectionName `"All WKS`"" -ForegroundColor Yellow
            Add-CMDeviceCollectionIncludeMembershipRule -CollectionName "DA | All WKS | $f_AppName | Prod | Required"`
                                    -IncludeCollectionName "All WKS"
            
            foreach ($collection_exclude in $collections_exclude)
            {
                Write-Host "Add-CMDeviceCollectionExcludeMembershipRule -CollectionName `"DA | All WKS | $f_AppName | Prod | Required`" -ExcludeCollectionName `"$collection_exclude`"" -ForegroundColor Yellow
                Add-CMDeviceCollectionExcludeMembershipRule -CollectionName "DA | All WKS | $f_AppName | Prod | Required"`
                                                            -ExcludeCollectionName $collection_exclude
            }
        }
        # Создаём коллекции для каждой площадки
        ForEach ($f_site in $f_sites)
        {
            $collections_exclude_sites = "Exclude | DA | _Install Apps", "Exclude | DA | $f_AppName", "DEP+VIP | $f_site | ALL", "DA | All WKS | $f_AppName | Migration | Required"
            Write-Host "Create collection 'DA | All WKS | $f_site | $f_AppName | Prod | Required'" -ForegroundColor Green
            Write-Host "New-CMDeviceCollection  -Name `"DA | All WKS | $f_site | $f_AppName | Prod | Required`" -LimitingCollectionName `"All WKS | $f_site`" -Comment `"Коллекция для установки `"`"продуктивной`"`" версии ПО $f_AppName на все рабочие станции в домене $f_site`" -RefreshSchedule `$Schedule | Move-CMObject -FolderPath $psite`":\DeviceCollection\Deploy Apps\$f_AppName`""  -ForegroundColor Yellow
            New-CMDeviceCollection  -Name "DA | All WKS | $f_site | $f_AppName | Prod | Required"`
                                    -LimitingCollectionName "All WKS | $f_site"`
                                    -Comment "Коллекция для установки ""продуктивной"" версии ПО $f_AppName на все рабочие станции в домене $f_site"`
                                    -RefreshSchedule $Schedule | Move-CMObject -FolderPath $psite":\DeviceCollection\Deploy Apps\$f_AppName"
            Write-Host "Add-CMDeviceCollectionIncludeMembershipRule -CollectionName `"DA | All WKS | $f_site | $f_AppName | Prod | Required`" -IncludeCollectionName `"All WKS | $f_site`"" -ForegroundColor Yellow
            Add-CMDeviceCollectionIncludeMembershipRule -CollectionName "DA | All WKS | $f_site | $f_AppName | Prod | Required"`
                                    -IncludeCollectionName "All WKS | $f_site"
            foreach ($collection_exclude in $collections_exclude_sites)
            {
                Write-Host "Add-CMDeviceCollectionExcludeMembershipRule -CollectionName `"DA | All WKS | $f_site | $f_AppName | Prod | Required`" -ExcludeCollectionName `"$collection_exclude`"" -ForegroundColor Yellow
                Add-CMDeviceCollectionExcludeMembershipRule -CollectionName "DA | All WKS | $f_site | $f_AppName | Prod | Required"`
                                                            -ExcludeCollectionName $collection_exclude
            }
        }
    }
}

function Create-VIPCollectionForDeployApplication (
    [Parameter(Mandatory = $true)]
    [string]$f_AppName,
    [Parameter(Mandatory = $true)]
    [array]$f_sites,
    [Parameter(Mandatory = $true)]
    [ValidateSet("Standard","Not standard")]
    [string]$AppType,
    [Parameter(Mandatory = $true)]
    [ValidateSet("Available","Required")]
    [string]$VIPInstallType,
    [Parameter(Mandatory = $true)]
    $Schedule,
    [Parameter(Mandatory = $false)]
    [array]$sites_VIP,
    [Parameter(Mandatory = $true)]
    [string]$Action

)
{
    Write-Host 'Function: Create-VIPCollectionForDeployApplication' -ForegroundColor Magenta
    Set-Location $psite":"
    $collections_exclude = "Exclude | DA | _Install Apps", "Exclude | DA | $f_AppName", "DA | VIP | All WKS | $f_AppName | Migration | $VIPInstallType"
    
    if ($AppType -eq "Not standard")
    {
        if ($Action -eq "Add new application")
        {
            # Создаём основную коллекцию для All WKS
            Write-Host "Create collection 'DA | VIP | All WKS | $f_AppName | Prod | $VIPInstallType'" -ForegroundColor Green
            Write-Host "New-CMDeviceCollection  -Name `"DA | VIP | All WKS | $f_AppName | Prod | $VIPInstallType`" -LimitingCollectionName `"INV | All WKS | $f_AppName`" -Comment `"Коллекция для установки `"`"продуктивной`"`" версии ПО $f_AppName на все рабочие станции.`" -RefreshSchedule `$Schedule | Move-CMObject -FolderPath $psite`":\DeviceCollection\Deploy Apps\$f_AppName`"" -ForegroundColor Yellow
            New-CMDeviceCollection  -Name "DA | VIP | All WKS | $f_AppName | Prod | $VIPInstallType"`
                                    -LimitingCollectionName "INV | All WKS | $f_AppName"`
                                    -Comment "Коллекция для установки ""продуктивной"" версии ПО $f_AppName на все рабочие станции."`
                                    -RefreshSchedule $Schedule | Move-CMObject -FolderPath $psite":\DeviceCollection\Deploy Apps\$f_AppName"
            Write-Host "Add-CMDeviceCollectionIncludeMembershipRule -CollectionName `"DA | VIP | All WKS | $f_AppName | Prod | $VIPInstallType`" -IncludeCollectionName `"DEP+VIP | ALL`"" -ForegroundColor Yellow
            Add-CMDeviceCollectionIncludeMembershipRule -CollectionName "DA | VIP | All WKS | $f_AppName | Prod | $VIPInstallType"`
                                                        -IncludeCollectionName "DEP+VIP | ALL"
            
            foreach ($collection_exclude in $collections_exclude)
            {
                Write-Host "Add-CMDeviceCollectionExcludeMembershipRule -CollectionName ""DA | VIP | All WKS | $f_AppName | Prod | $VIPInstallType"" -ExcludeCollectionName `"$collection_exclude`"" -ForegroundColor Yellow
                Add-CMDeviceCollectionExcludeMembershipRule -CollectionName "DA | VIP | All WKS | $f_AppName | Prod | $VIPInstallType"`
                                                            -ExcludeCollectionName $collection_exclude
            }
        }

        # Создаём коллекции для каждой площадки
        ForEach ($f_site in $sites_VIP)
        {
            $collections_exclude_sites = "Exclude | DA | _Install Apps", "Exclude | DA | $f_AppName", "DA | VIP | All WKS | $f_AppName | Migration | $VIPInstallType"
            Write-Host "Create collection 'DA | VIP | All WKS | $f_site | $f_AppName | Prod | $VIPInstallType'" -ForegroundColor Green
            Write-Host "New-CMDeviceCollection  -Name `"DA | VIP | All WKS | $f_site | $f_AppName | Prod | $VIPInstallType`" -LimitingCollectionName `"INV | All WKS | $f_site | $f_AppName`" -Comment `"Коллекция для установки `"`"продуктивной`"`" версии ПО $f_AppName на все рабочие станции в домене $f_site`" -RefreshSchedule `$Schedule | Move-CMObject -FolderPath $psite`":\DeviceCollection\Deploy Apps\$f_AppName`"" -ForegroundColor Yellow
            New-CMDeviceCollection  -Name "DA | VIP | All WKS | $f_site | $f_AppName | Prod | $VIPInstallType"`
                                    -LimitingCollectionName "INV | All WKS | $f_site | $f_AppName"`
                                    -Comment "Коллекция для установки ""продуктивной"" версии ПО $f_AppName на все рабочие станции в домене $f_site"`
                                    -RefreshSchedule $Schedule | Move-CMObject -FolderPath $psite":\DeviceCollection\Deploy Apps\$f_AppName"

            Write-Host "Add-CMDeviceCollectionIncludeMembershipRule -CollectionName `"DA | VIP | All WKS | $f_site | $f_AppName | Prod | $VIPInstallType`" -IncludeCollectionName `"DEP+VIP | $f_site | ALL`"" -ForegroundColor Yellow
            Add-CMDeviceCollectionIncludeMembershipRule -CollectionName "DA | VIP | All WKS | $f_site | $f_AppName | Prod | $VIPInstallType"`
                                    -IncludeCollectionName "DEP+VIP | $f_site | ALL"

            foreach ($collection_exclude in $collections_exclude_sites)
            {
                Write-Host "Add-CMDeviceCollectionExcludeMembershipRule -CollectionName ""DA | VIP | All WKS | $f_site | $f_AppName | Prod | $VIPInstallType"" -ExcludeCollectionName `"$collection_exclude`"" -ForegroundColor Yellow
                Add-CMDeviceCollectionExcludeMembershipRule -CollectionName "DA | VIP | All WKS | $f_site | $f_AppName | Prod | $VIPInstallType"`
                                                            -ExcludeCollectionName $collection_exclude
            }
        }
    }
    if ($AppType -eq "Standard")
    {
        if ($Action -eq "Add new application")
        {
            # Создаём основную коллекцию для All WKS
            Write-Host "Create collection 'DA | VIP | All WKS | $f_AppName | Prod | $VIPInstallType'" -ForegroundColor Green
            Write-Host "New-CMDeviceCollection  -Name `"DA | VIP | All WKS | $f_AppName | Prod | $VIPInstallType`" -LimitingCollectionName `"All WKS`" -Comment `"Коллекция для установки `"`"продуктивной`"`" версии ПО $f_AppName на все рабочие станции.`" -RefreshSchedule `$Schedule | Move-CMObject -FolderPath $psite`":\DeviceCollection\Deploy Apps\$f_AppName`"" -ForegroundColor Yellow
            New-CMDeviceCollection  -Name "DA | VIP | All WKS | $f_AppName | Prod | $VIPInstallType"`
                                    -LimitingCollectionName "All WKS"`
                                    -Comment "Коллекция для установки ""продуктивной"" версии ПО $f_AppName на все рабочие станции."`
                                    -RefreshSchedule $Schedule | Move-CMObject -FolderPath $psite":\DeviceCollection\Deploy Apps\$f_AppName"
            Write-Host "Add-CMDeviceCollectionIncludeMembershipRule -CollectionName `"DA | VIP | All WKS | $f_AppName | Prod | $VIPInstallType`" -IncludeCollectionName `"DEP+VIP | ALL`"" -ForegroundColor Yellow
            Add-CMDeviceCollectionIncludeMembershipRule -CollectionName "DA | VIP | All WKS | $f_AppName | Prod | $VIPInstallType"`
                                    -IncludeCollectionName "DEP+VIP | ALL"
            
            foreach ($collection_exclude in $collections_exclude)
            {
                Write-Host "Add-CMDeviceCollectionExcludeMembershipRule -CollectionName `"DA | VIP | All WKS | $f_AppName | Prod | $VIPInstallType`" -ExcludeCollectionName `"$collection_exclude`"" -ForegroundColor Yellow
                Add-CMDeviceCollectionExcludeMembershipRule -CollectionName "DA | VIP | All WKS | $f_AppName | Prod | $VIPInstallType"`
                                                            -ExcludeCollectionName $collection_exclude
            }
        }

        # Создаём коллекции для каждой площадки
        ForEach ($f_site in $sites_VIP)
        {
            $collections_exclude_sites = "Exclude | DA | _Install Apps", "Exclude | DA | $f_AppName", "DA | VIP | All WKS | $f_AppName | Migration | $VIPInstallType"
            Write-Host "Create collection 'DA | VIP | All WKS | $f_site | $f_AppName | Prod | $VIPInstallType'" -ForegroundColor Green
            Write-Host "New-CMDeviceCollection  -Name `"DA | VIP | All WKS | $f_site | $f_AppName | Prod | $VIPInstallType`" -LimitingCollectionName `"All WKS | $f_site`" -Comment `"Коллекция для установки `"`"продуктивной`"`" версии ПО $f_AppName на все рабочие станции в домене $f_site`" -RefreshSchedule `$Schedule | Move-CMObject -FolderPath $psite`":\DeviceCollection\Deploy Apps\$f_AppName`"" -ForegroundColor Yellow
            New-CMDeviceCollection  -Name "DA | VIP | All WKS | $f_site | $f_AppName | Prod | $VIPInstallType"`
                                    -LimitingCollectionName "All WKS | $f_site"`
                                    -Comment "Коллекция для установки ""продуктивной"" версии ПО $f_AppName на все рабочие станции в домене $f_site"`
                                    -RefreshSchedule $Schedule | Move-CMObject -FolderPath $psite":\DeviceCollection\Deploy Apps\$f_AppName"
            
            Write-Host "Add-CMDeviceCollectionIncludeMembershipRule -CollectionName `"DA | VIP | All WKS | $f_site | $f_AppName | Prod | $VIPInstallType`" -IncludeCollectionName `"DEP+VIP | $f_site | ALL`"" -ForegroundColor Yellow
            Add-CMDeviceCollectionIncludeMembershipRule -CollectionName "DA | VIP | All WKS | $f_site | $f_AppName | Prod | $VIPInstallType"`
                                    -IncludeCollectionName "DEP+VIP | $f_site | ALL"
            
            foreach ($collection_exclude in $collections_exclude_sites)
            {
                Write-Host "Add-CMDeviceCollectionExcludeMembershipRule -CollectionName `"DA | VIP | All WKS | $f_site | $f_AppName | Prod | $VIPInstallType`" -ExcludeCollectionName `"$collection_exclude`"" -ForegroundColor Yellow
                Add-CMDeviceCollectionExcludeMembershipRule -CollectionName "DA | VIP | All WKS | $f_site | $f_AppName | Prod | $VIPInstallType"`
                                                            -ExcludeCollectionName $collection_exclude
            }
        }
    }
}