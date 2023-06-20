<#
    .SYNOPSIS
    Exports / imports a CSV with the given SearchBase for changing AD-User Properties
    Maik Luedeke
    Version 1.0
    2023-06
    .DESCRIPTION
#>
param(
   [parameter(Position=0,
               Mandatory=$true,
               ValueFromPipeline=$false,
               HelpMessage='Action (export/import)')][string]$action,
    [parameter(Position=1,
               Mandatory=$true,
               ValueFromPipeline=$false,
               HelpMessage='OU (e.g. OU=MainSite,OU=Company)..., domain (e.g. DC=Example,DC=COM) will be added automatically')][string]$OU,
    [parameter(Position=2,
               Mandatory=$false,
               ValueFromPipeline=$false,
               HelpMessage='Only test import, the Set-ADUser will be executed with "WhatIf"')][string]$testOnly
    )

Set-StrictMode -version 2
Import-Module ActiveDirectory
# to get all possible Properties use ==> Get-ADUser <user> -Properties * | Get-Member
$Properties = "Name,Title,Department,Description,Company,OfficePhone,Office,WWWHomePage,Manager,sAMAccountName"
# getting current Domain
$currentDomain = (Get-ADDomain -Current LocalComputer).DistinguishedName
$searchBase = "$OU,$currentDomain"
# name of the export file, location is currently the same as the script folder
$csv = "ADUserExport.csv"
# adding WhatIf (and another color) to the command, when $testOnly ist set to $true
if($testOnly -eq $true)
{
    $whatIf = " -WhatIf"
    $color = "DarkYellow"
}
else
{
    $whatIf = ""
    $color = "Green"
}

# actions export/import
switch ($action)
{
    export
    {
        # getting all Properties from AD for SearchBase $OU
        $ADUsers = Get-ADUser -Filter { Enabled -eq $true } -Properties ($Properties -split ",") -SearchBase $searchBase -SearchScope Subtree

        Write-Host "Reading Properties from $($ADUsers.Count) AD-accounts and exporting them to '$($csv)' for OU '$OU'" -ForegroundColor $color
        # creating an export array
        $exportCSV = @()

        foreach ($ADUser in $ADUsers)
        {
            $values = [ordered]@{} # hastable for holding all Properties

            foreach($property in $Properties -split ",")
            {
                # only the Manager name is extracted, usally "Manager" contains the DistinguishedName

                if ('Manager' -eq $property -and $null -ne $ADUser.Manager)
                {
                    $manager = Get-ADUser($ADUser.Manager)|Select-Object Name, Enabled
                    if ($manager.Enabled -eq $true)
                    {
                        $values.add('Manager', $manager.Name)
                    }
                    else
                    {
                        $values.add('Manager', $null)
                    }
                }
                else
                {
                    # all normal values are added to the array
                    $values.add($property, $ADUser.$property)
                }
            }
            # the array is added to the exportCSV variable
            $exportCSV += $values
        }
        # the whole array is exported to a CSV
        $exportCSV | ForEach-Object{ [pscustomobject]$_ } | Export-CSV -encoding utf8 -notypeinformation -Delimiter ";" $csv
    }
    import
    {
        Write-Host("Trying to import $($csv). Comparing values with existing ones") -ForegroundColor $color
        Write-Host("----------------------------------------------------------------") -ForegroundColor $color

        # importing the file with the same name, as the export
        Import-CSV -Path $csv -Delimiter ";" -Encoding UTF8 | Foreach-Object {
            #requesting user from AD for comparison
            $ADUser = Get-ADUser $_.sAMAccountName -Properties ($Properties -split ",")

            if($ADUser)
            {
                foreach($property in $Properties -split ",")
                {
                    $command = ""
                    if([bool]($_.PSobject.Properties.name -match $property) -and "" -ne $_.$property)
                    {
                        if ($property -eq "Manager")
                        {
                            # check if the manager is existent and enabled
                            $manager = (Get-ADUser -LdapFilter "(Name=$($_.Manager))(useraccountcontrol=512)")

                            if ($manager -and $manager.Enabled -eq $true -and $manager.DistinguishedName -ne $ADUser.Manager)
                            {
                                $command = "Get-ADUser $($_.sAMAccountName) | Set-ADUser -Manager '$($manager.DistinguishedName)' $whatIf"
                            }
                        }
                        else
                        {
                            #test if property is changed or new, if so set property to AD Account
                            if ($_.$property -ne $ADUser.$property)
                            {
                                # if testOnly is true, a -whatIf will be attached, so it will only dry perform the update action
                                $command = "Get-ADUser $($_.sAMAccountName) | Set-ADUser -$property '$($_.$property)' $whatIf"
                            }
                        }

                        if("" -ne $_.$property -and "" -ne $command)
                        {
                            # it is not possible to use Set-ADUser with dynamic property setting e.g. Set-ADUser -$property ...
                            # so the command is created as a string and then executed by Invoke-Expression
                            Write-Host('ADUser: ' + $ADUser.Name  + " => Property: " + $property + " =>(old) '" + $ADUser.$property + "' >>(new) '" + $_.$property + "'") -ForegroundColor $color
                            Invoke-Expression $command
                        }
                    }
                }

            }
        }
    }
    default
    {
        Write-Host("Action must be one of 'export'/'import'") -ForegroundColor Red
    }
}
