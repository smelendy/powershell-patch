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
	   $ComputerNM = $env:computername
	   $ComputerUSR = $env:username
	   $ComputerDomain = $env:userdomain
	   $ComputerSource = $env:clientname
	     # $ComputerInfo = New-Object System.Object 
       "<HTML><head> <title> Pre-Patch Report </title> $head </head> <style> body {background-color: lightgrey;} </style> <center>" > $ComputerNM-prepatch.html
#	   $header| Out-File -FilePath $ComputerNM-prepatch.html
	    $ComputerSW =Get-WmiObject -Class win32_product|select Name,Vendor,Version,Caption,InstallDate |sort-object Name 
	   $ComputerHot = Get-hotfix |select Description,HotFixID,InstalledOn,InstalledBy|sort-object InstalledOn -Descending 
	   $ComputerHW =Get-WmiObject -Class Win32_ComputerSystem | select Manufacturer, Model, @{Expression={$_.TotalPhysicalMemory / 1GB};Label="TotalPhysicalMemoryGB"} 
       $ComputerCPU =Get-WmiObject win32_processor  | select DeviceID, Name, Manufacturer, NumberOfCores, NumberOfLogicalProcessors 
       $ComputerDisks =Get-WmiObject -Class Win32_LogicalDisk -Filter "DriveType=3" |select DeviceID,VolumeName, @{Expression={$_.Size / 1GB};Label="SizeGB"},@{Expression={$_.freespace / 1GB};Label="FreeSpaceGB"}  
	   $ComputerBios =Get-WmiObject Win32_Bios| select SMBIOSBIOSVersion,Manufacturer,Name,SerialNumber 
	   "<H2> Pre-Patch Report for $ComputerNM by $ComputerUSR on $date</H2>" >> $ComputerNM-prepatch.html
	   #$html +=$ComputerInfo | ConvertTo-Html -fragment 
	   $ComputerHW | ConvertTo-Html  -Head $styler |Out-File -append  -FilePath $ComputerNM-prepatch.html
	   $ComputerCPU | ConvertTo-Html -Head $styler |Out-File -append -FilePath $ComputerNM-prepatch.html
	   $ComputerBios | ConvertTo-Html -Head $styler|Out-File -append  -FilePath $ComputerNM-prepatch.html
       "<H3> Drives </H3>" >> $ComputerNM-prepatch.html
	   $ComputerDisks | ConvertTo-Html -Head $styler |Out-File  -append -FilePath $ComputerNM-prepatch.html
	   "<H3> Hotfix info </H3>" >> $ComputerNM-prepatch.html
	   $ComputerHot | ConvertTo-Html -Head $styler |Out-File -append -FilePath $ComputerNM-prepatch.html
	   "<H3> Registered Software</H3>" >> $ComputerNM-prepatch.html
	   $ComputerSW | ConvertTo-Html -Head $styler |Out-File -append -FilePath $ComputerNM-prepatch.html
	   "END of Report </center></body> </html>"  >> $ComputerNM-prepatch.html
	   