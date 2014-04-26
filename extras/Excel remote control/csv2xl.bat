@echo off
rem -----------------------------------------------------------------------------
rem  CSV2XL.bat  Import of a (special) CSV file into Excel via Add-In "GeoTools"
rem -----------------------------------------------------------------------------
rem  
rem  Parameter 1: - optional
rem               - CSV filename
rem -----------------------------------------------------------------------------

start wscript xlM.vbs /M:Import_CSV /D:""%1"" /silent:true %2 %3
