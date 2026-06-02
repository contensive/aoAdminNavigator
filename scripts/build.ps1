#Requires -Version 5.1
[CmdletBinding()]
param(
    [string]   $LocalDeployTarget  = '',
    [hashtable]$RemoteDeployTarget = $null
)

$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot '..\..\Contensive5\scripts\contensive-build.psm1') -Force

$projectRoot = (Resolve-Path "$PSScriptRoot\..").Path

Invoke-ContensiveBuild `
    -CollectionName    'Admin Navigator' `
    -CollectionPath    "$projectRoot\collections\Admin Navigator" `
    -SolutionPath      "$projectRoot\source\adminNavigator.sln" `
    -BinPath           "$projectRoot\source\bin\Release\netstandard2.0" `
    -DeploymentRoot    'C:\Deployments\aoAdminNavigator' `
    -CleanFolders      @(
                           "$projectRoot\source\bin"
                           "$projectRoot\source\obj"
                       ) `
    -UiPath            "$projectRoot\ui" `
    -LocalDeployTarget  $LocalDeployTarget `
    -RemoteDeployTarget $RemoteDeployTarget
