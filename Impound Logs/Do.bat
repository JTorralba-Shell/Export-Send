@echo off

CScript Export.vbs

PowerShell -ExecutionPolicy ByPass -file "Send.ps1"

@echo on