@echo off
rem Opens an elevated PowerShell window and runs ScanSerialPorts.ps1 from this folder
PowerShell -NoProfile -ExecutionPolicy Bypass -Command "Start-Process PowerShell -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File ""%~dp0ScanSerialPorts.ps1""' -Verb RunAs"
