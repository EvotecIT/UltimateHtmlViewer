@echo off
setlocal

REM Thin wrapper so people can run "deploy-uhv" even on case-insensitive filesystems.
REM Usage:
REM   .\scripts\deploy-uhv.cmd -AppCatalogUrl "https://<tenant>.sharepoint.com/sites/appcatalog" [-TenantWide]

pwsh -NoProfile -ExecutionPolicy Bypass -File "%~dp0Deploy-UHV-Wrapper.ps1" %*

