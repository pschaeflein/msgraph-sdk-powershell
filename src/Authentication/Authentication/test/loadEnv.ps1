# ----------------------------------------------------------------------------------
#
# Copyright Microsoft Corporation
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
# http://www.apache.org/licenses/LICENSE-2.0
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.
# ----------------------------------------------------------------------------------
$envFile = 'env.json'
if ($TestMode -eq 'live') {
    $envFile = 'localEnv.json'
}

if (Test-Path -Path (Join-Path $PSScriptRoot $envFile)) {
    $envFilePath = Join-Path $PSScriptRoot $envFile
} else {
    $envFilePath = Join-Path $PSScriptRoot '..\$envFile'
}
$env = @{}
if (Test-Path -Path $envFilePath) {
    # Load dummy auth configuration.
    $env = Get-Content (Join-Path $PSScriptRoot $envFile) | ConvertFrom-Json -AsHashTable
    $env:DEFAULTUSERID = $env.DefaultUserIdentifier
} else{
    $env.TenantIdentifier = ${env:TENANTIDENTIFIER}
    $env.ClientIdentifier = ${env:CLIENTIDENTIFIER}
    $env.CertificateThumbprint = ${env:CERTIFICATETHUMBPRINT}
}
$PSDefaultParameterValues+=@{"Connect-MgGraph:TenantId"=$env.TenantIdentifier; "Connect-MgGraph:ClientId"=$env.ClientIdentifier; "Connect-MgGraph:CertificateThumbprint"=$env.CertificateThumbprint}
[Microsoft.Graph.PowerShell.Authentication.GraphSession]::Instance.AuthContext = New-Object Microsoft.Graph.PowerShell.Authentication.AuthContext -Property @{
    ClientId              = $env.ClientIdentifier
    TenantId              = $env.TenantIdentifier
    CertificateThumbprint = $env.CertificateThumbprint
}