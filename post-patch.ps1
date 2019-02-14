#PowerShell Scrpt for UMW Windows server patching 
#Created by Shaun Melendy smelendy@umw.edu  http://github.com/smelendy

$styler = @"
<style>
TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}
TD {border-width: 1px; padding: 3px; border-style: solid; border-color: black;}
TH {border-width: 1px; padding: 3px; border-style: solid; border-color: black; background-color: #6495ED;}
tr:nth-child(even) {background-color: #f2f2f2;}

</style>
"@
		# Get Base Strings from ENV
	   $date = (Get-Date)
	   $SWDATE =(Get-Date -format "yyyymmdd")
	   $IntDate = [INT]$SWDATE
	   
	   $ComputerNM = $env:computername
	   $ComputerUSR = $env:username
	   $ComputerDomain = $env:userdomain
	   $ComputerSource = $env:clientname
	     # $ComputerInfo = New-Object System.Object 
       "<HTML><head> <title> Post Patch Report </title> $head </head> <style> body {background-color: lightgrey;} </style> <center>" > $ComputerNM-postpatch.html
#	   $header| Out-File -FilePath $ComputerNM-postpatch.html
#	    $ComputerSW =Get-WmiObject -Class win32_product|Where { $_.Installdate -eq (Get-Date -format #"yyyymmdd")} |select-object Name,Vendor,Version,Caption,InstallDate
	    $ComputerSW =Get-WmiObject -Class win32_product|Where { $_.Installdate -gt ($INTDate -1)} |select-object Name,Vendor,Version,Caption,InstallDate

		
	   #$ComputerHot = Get-hotfix |select Description,HotFixID,InstalledOn,InstalledBy|sort-object -property HotFixID
     #  $ComputerHot = Get-hotfix |select Description,HotFixID,InstalledOn,InstalledBy|sort-object InstalledOn -Descending
	   $ComputerHW =Get-WmiObject -Class Win32_ComputerSystem | select Manufacturer, Model, @{Expression={$_.TotalPhysicalMemory / 1GB};Label="TotalPhysicalMemoryGB"} 
       $ComputerCPU =Get-WmiObject win32_processor  | select DeviceID, Name, Manufacturer, NumberOfCores, NumberOfLogicalProcessors 
       $ComputerDisks =Get-WmiObject -Class Win32_LogicalDisk -Filter "DriveType=3" |select DeviceID,VolumeName, @{Expression={$_.Size / 1GB};Label="SizeGB"}  
	   $ComputerBios =Get-WmiObject Win32_Bios| select SMBIOSBIOSVersion,Manufacturer,Name,SerialNumber 
	   "<H2> Post Patch Report for $ComputerNM by $ComputerUSR on $date</H2>" >> $ComputerNM-postpatch.html
	   #$html +=$ComputerInfo | ConvertTo-Html -fragment 
	   $ComputerHW | ConvertTo-Html  -Head $styler |Out-File -append  -FilePath $ComputerNM-postpatch.html
	   $ComputerCPU | ConvertTo-Html -Head $styler |Out-File -append -FilePath $ComputerNM-postpatch.html
	   $ComputerBios | ConvertTo-Html -Head $styler|Out-File -append  -FilePath $ComputerNM-postpatch.html
       "<H3> Drives </H3>" >> $ComputerNM-postpatch.html
	   $ComputerDisks | ConvertTo-Html -Head $styler |Out-File  -append -FilePath $ComputerNM-postpatch.html
	   "<H3> Hotfixes installed this patching  </H3>" >> $ComputerNM-postpatch.html
	  # $ComputerHot | ConvertTo-Html -Head $styler |Out-File -append -FilePath $AComputerNM-postpatch.html
	  Get-WmiObject win32_quickfixengineering| Where { $_.InstalledOn -gt (Get-Date).AddDays(-2) } |sort InstalledOn | select-object Description,HotFixID,InstalledOn,InstalledBy |  ConvertTo-Html -Head $styler |Out-File  -append -FilePath $ComputerNM-postpatch.html

	   "<H3> Registered Software changes this patching </H3>" >> $ComputerNM-postpatch.html
	   $ComputerSW | ConvertTo-Html -Head $styler |Out-File -append -FilePath $ComputerNM-postpatch.html
	   "END of Report</center>  </body> </html>"  >> $ComputerNM-postpatch.html
	   