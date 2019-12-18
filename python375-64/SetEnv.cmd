@echo off
set currPath=%~dp0
rem setx PATH "%PATH%;%currPath%Python37;%currPath%Python37\Scripts"
setx PATH "%PATH%;%currPath%;%currPath%Scripts"