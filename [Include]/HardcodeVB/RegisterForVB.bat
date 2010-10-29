@Echo Off
Exes\RegTlb /s Release\Win.tlb
If Exist Release\WinU.tlb Exes\RegTlb /s Release\WinU.tlb
If Exist Release\Cards32.dll Move Release\Cards32.dll %windir%
If Exist Release\PSAPI.dll Move Release\PSAPI.dll %windir%
