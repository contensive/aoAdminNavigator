#Requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$CollectionName,
    [Parameter(Mandatory)][string]$CollectionPath,
    [Parameter(Mandatory)][string]$DeploymentPath,
    [string]$UiPath = '',
    [string]$TargetDomain = 'www.contensive.com'
)
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot '..\..\Contensive5\scripts\deploy-to-addon-library.psm1') -Force
Invoke-AddonLibraryDeploy `
    -CollectionName $CollectionName `
    -CollectionPath $CollectionPath `
    -DeploymentPath $DeploymentPath `
    -UiPath         $UiPath `
    -TargetDomain   $TargetDomain
