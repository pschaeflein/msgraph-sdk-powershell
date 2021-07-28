# ------------------------------------------------------------------------------
#  Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
# ------------------------------------------------------------------------------

Set-StrictMode -Version 2

<#
.SYNOPSIS
Find permissions for authorization against Microsoft Graph
#>

. "$psscriptroot/common/Permissions.ps1"

# Signing into any AAD organization

# Import-Module -Force .\Microsoft.Graph.Authentication.psd1

function Find-MgGraphPermission {
    [cmdletbinding()]
    param (
        [string] $SearchString,
        [switch] $online
    )

    if ($online.IsPresent){

        $Permissions_GetOauthData = Permissions_GetOauthData -online
        $Permissions_GetAppRolesData = Permissions_GetAppRolesData -online

    } else {
        
        $Permissions_GetOauthData = Permissions_GetOauthData
        $Permissions_GetAppRolesData = Permissions_GetAppRolesData

    }

    # Creating a table specifically for Oauth2permissions data 
    $oauthData = @()
    $oauthData += $Permissions_GetOauthData | Where-Object Name -like *$SearchString*
    # $oauthData += Permissions_GetOauthData | Where-Object Description -like *$search* 
    $oauthData | Sort-Object -Property Name

    # Creating a table specifically for appRoles data 
    $appRolesData = @()
    $appRolesData += $Permissions_GetAppRolesData | Where-Object Name -like *$SearchString*
    # $appRolesData += Permissions_GetAppRolesData | Where-Object Description -like *$search*
    $appRolesData | Sort-Object -Property Name

}
