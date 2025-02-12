﻿Describe 'Invoke-MgGraphRequest Command' -Skip {
     BeforeAll {
          $loadEnvPath = Join-Path $PSScriptRoot 'loadEnv.ps1'
          if (-Not (Test-Path -Path $loadEnvPath)) {
               $loadEnvPath = Join-Path $PSScriptRoot '..\loadEnv.ps1'
          }
          . ($loadEnvPath)
          $ModuleName = "Microsoft.Graph.Authentication"
          $ModulePath = Join-Path $PSScriptRoot "..\artifacts\$ModuleName.psd1"
          Import-Module $ModulePath -Force
          $PSDefaultParameterValues = @{"Connect-MgGraph:TenantId" = ${env:TENANTIDENTIFIER}; "Connect-MgGraph:ClientId" = ${env:CLIENTIDENTIFIER}; "Connect-MgGraph:CertificateThumbprint" = ${env:CERTIFICATETHUMBPRINT} }
          Connect-MgGraph
     }
     Context 'Collection Results' {
          It 'ShouldReturnPsObject' {
               Invoke-MgGraphRequest -OutputType PSObject -Uri "https://graph.microsoft.com/v1.0/users" | Should -BeOfType [System.Management.Automation.PSObject]
          }

          It 'ShouldReturnHashTable' {
               Invoke-MgGraphRequest -OutputType Hashtable -Uri "https://graph.microsoft.com/v1.0/users" | Should -BeOfType [System.Collections.Hashtable]
          }

          It 'ShouldReturnJsonString' {
               Invoke-MgGraphRequest -OutputType Json -Uri "https://graph.microsoft.com/v1.0/users" | Should -BeOfType [System.String]
          }

          It 'ShouldReturnHttpResponseMessage' {
               Invoke-MgGraphRequest -OutputType HttpResponseMessage -Uri "https://graph.microsoft.com/v1.0/users" | Should -BeOfType [System.Net.Http.HttpResponseMessage]
          }

          It 'ShouldReturnPsObjectBeta' {
               Invoke-MgGraphRequest -OutputType PSObject -Uri "https://graph.microsoft.com/beta/users" | Should -BeOfType [System.Management.Automation.PSObject]
          }

          It 'ShouldReturnHashTableBeta' {
               Invoke-MgGraphRequest -OutputType Hashtable -Uri "https://graph.microsoft.com/beta/users" | Should -BeOfType [System.Collections.Hashtable]
          }

          It 'ShouldReturnJsonStringBeta' {
               Invoke-MgGraphRequest -OutputType Json -Uri "https://graph.microsoft.com/beta/users" | Should -BeOfType [System.String]
          }

          It 'ShouldReturnHttpResponseMessageBeta' {
               Invoke-MgGraphRequest -OutputType HttpResponseMessage -Uri "https://graph.microsoft.com/beta/users" | Should -BeOfType [System.Net.Http.HttpResponseMessage]
          }
     }
     Context 'Single Entity' {
          It 'ShouldReturnPsObject' {
               Invoke-MgGraphRequest -OutputType PSObject -Uri "https://graph.microsoft.com/v1.0/users/${env:DEFAULTUSERID}" | Should -BeOfType [System.Management.Automation.PSObject]
          }

          It 'ShouldReturnHashTable' {
               Invoke-MgGraphRequest -OutputType Hashtable -Uri "https://graph.microsoft.com/v1.0/users/${env:DEFAULTUSERID}" | Should -BeOfType [System.Collections.Hashtable]
          }

          It 'ShouldReturnJsonString' {
               Invoke-MgGraphRequest -OutputType Json -Uri "https://graph.microsoft.com/v1.0/users/${env:DEFAULTUSERID}" | Should -BeOfType [System.String]
          }

          It 'ShouldReturnHttpResponseMessage' {
               Invoke-MgGraphRequest -OutputType HttpResponseMessage -Uri "https://graph.microsoft.com/v1.0/users/${env:DEFAULTUSERID}" | Should -BeOfType [System.Net.Http.HttpResponseMessage]
          }

          It 'ShouldReturnPsObjectBeta' {
               Invoke-MgGraphRequest -OutputType PSObject -Uri "https://graph.microsoft.com/beta/users/${env:DEFAULTUSERID}" | Should -BeOfType [System.Management.Automation.PSObject]
          }

          It 'ShouldReturnHashTableBeta' {
               Invoke-MgGraphRequest -OutputType Hashtable -Uri "https://graph.microsoft.com/beta/users/${env:DEFAULTUSERID}" | Should -BeOfType [System.Collections.Hashtable]
          }

          It 'ShouldReturnJsonStringBeta' {
               Invoke-MgGraphRequest -OutputType Json -Uri "https://graph.microsoft.com/beta/users/${env:DEFAULTUSERID}" | Should -BeOfType [System.String]
          }

          It 'ShouldReturnHttpResponseMessageBeta' {
               Invoke-MgGraphRequest -OutputType HttpResponseMessage -Uri "https://graph.microsoft.com/beta/users/${env:DEFAULTUSERID}" | Should -BeOfType [System.Net.Http.HttpResponseMessage]
          }
     }

     Context 'Non-Json Responses' {
          It 'Should Throw when -OutputFilePath is not Specified and Request Returns a Non-Json Response' {
               { Invoke-MgGraphRequest -OutputType PSObject -Uri "https://graph.microsoft.com/v1.0/reports/getTeamsUserActivityUserDetail(period='D7')" } | Should -Throw
          }
          It 'Should Not Throw when -OutputFilePath is Specified' {
               { Invoke-MgGraphRequest -OutputType PSObject -Uri "https://graph.microsoft.com/v1.0/reports/getTeamsUserActivityUserDetail(period='D7')" -OutputFilePath ./data.csv } | Should -Not -Throw
          }
          It 'Should Not Throw when -InferOutputFilePath is Specified' {
               { Invoke-MgGraphRequest -OutputType PSObject -Uri "https://graph.microsoft.com/v1.0/reports/getTeamsUserActivityUserDetail(period='D7')" -InferOutputFileName } | Should -Not -Throw
          }
     }

     AfterAll {
          Disconnect-MgGraph
     }
}