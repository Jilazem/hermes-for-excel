@echo off
title Hermes Excel baslatiliyor...
rem Hata-AYIKLAMASIZ sideload: add-in yuklenir ama "hata ayikla" penceresi CIKMAZ.
npx --yes office-addin-dev-settings sideload "C:\HermesExcel\manifest.xml" desktop --app Excel
