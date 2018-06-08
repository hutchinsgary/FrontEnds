# FrontEnds
FrontEnd for creating VM shell






Code
 <#	


	.NOTES


	===========================================================================


	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2015 v4.2.98


	 Created on:   	15/04/2016 14:56


	 Created by:	Gary Hutchins   	 


	 Organization:		 


	 Filename:     	


	===========================================================================


	.DESCRIPTION


		This script presents a front-end to the user which allows them to provision a 


		VM shell based on drop downs which are retrived from a central SQL database


		#>



################################################################################################################


#Functions#


################################################################################################################


function Connect-SqlandPopulate {


	#Get info from the FrontEnd Database SQL Table to populate the dropdowns


	##############################################################################################################


	


	Clear-Variable -name mdtDatabase -errorAction SilentlyContinue


	


	$SQLServer = "sqlservername,1504"


	$SQLDBName = "MDT_PRO"


	$SqlQuery = "select * from FrontEndVMShell"


	


	$global:mdSiteLConnection = New-Object System.Data.SqlClient.SqlConnection


	$mdSiteLConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; Integrated Security = True"


	$mdSiteLConnection.Open()


	


	$SqlCmd = New-Object System.Data.SqlClient.SqlCommand


	$SqlCmd.CommandText = $SqlQuery


	$SqlCmd.Connection = $global:mdSiteLConnection


	


	$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter


	$SqlAdapter.SelectCommand = $SqlCmd


	


	$DataSet = New-Object System.Data.DataSet


	$SqlAdapter.Fill($DataSet)


	


	$OS = $DataSet.Tables[0] | select -ExpandProperty OS


	$Environment = $DataSet.Tables[0] | select -ExpandProperty Environment


	$Location = $DataSet.Tables[0] | select -ExpandProperty Location


	$CPU = $DataSet.Tables[0] | select -ExpandProperty CPU


	$Memory = $DataSet.Tables[0] | select -ExpandProperty Memory


	$Disk1 = $DataSet.Tables[0] | select -ExpandProperty Disk1


	$Disk2 = $DataSet.Tables[0] | select -ExpandProperty Disk2


	$Network = $DataSet.Tables[0] | select -ExpandProperty Network


	$AdminNetwork = $DataSet.Tables[0] | select -ExpandProperty Admin_Network


	$VMFolder = $DataSet.Tables[0] | select -ExpandProperty VMFolder


	$DevVMFolder = $DataSet.Tables[0] | select -ExpandProperty DevVMFolder


	$PreProdVMFolder = $DataSet.Tables[0] | select -ExpandProperty PreProdVMFolder


	$ProdDRVMFolder = $DataSet.Tables[0] | select -ExpandProperty ProdDRVMFolder


	$DSCluster = $DataSet.Tables[0] | select -ExpandProperty DatastoreCluster


		


	


	##############################################################################################################


	# Populate the Dropdowns on the form


	##############################################################################################################


	# Populate OS Dropdown


	$objOSArray = @()


	$I = 0


	


	foreach ($O in $OS)


	{


		If ($O -ne "")


		{


			$obj = new-object System.Object


			$obj | add-member -type NoteProperty -name OS -value $OS[$I]


			$I++


			$objOSArray += $obj


		}


	}


	


	$VMOS_Combobox.items.addrange($objOSArray)


	$VMOS_Combobox.DisplayMember = "OS"


	$obj = ""


	# END Populate OS Dropdown


	


	# Populate Environment Dropdown


	$global:objEnvArray = @()


	$I = 0


	


	foreach ($E in $Environment)


	{


		If ($E -ne "")


		{


			$obj = new-object System.Object


			$obj | add-member -type NoteProperty -name Environment -value $Environment[$I]


			$I++


			$global:objEnvArray += $obj


		}


	}


	


	$VMEnvironment_Combobox.items.addrange($global:objEnvArray)


	$VMEnvironment_Combobox.DisplayMember = "Environment"


	$obj = ""


	# END Populate Environment Dropdown


	


	# Populate Location Dropdown


	$global:objLocationArray = @()


	$I = 0


	


	foreach ($L in $Location)


	{


		If ($L -ne "")


		{


			$obj = new-object System.Object


			$obj | add-member -type NoteProperty -name Location -value $Location[$I]


			$I++


			$global:objLocationArray += $obj


		}


	}


	


	$VMLocation_Combobox.items.addrange($global:objLocationArray)


	$VMLocation_Combobox.DisplayMember = "Location"


	$obj = ""


	#END Populate Location Dropdown


	


	# Populate CPU Dropdown


	$objCPUArray = @()


	$I = 0


	


	foreach ($C in $CPU)


	{


		If ($C -ne "")


		{


			$obj = new-object System.Object


			$obj | add-member -type NoteProperty -name CPU -value $CPU[$I]


			$I++


			$objCPUArray += $obj


		}


	}


	


	$VMCPU_Combobox.items.addrange($objCPUArray)


	$VMCPU_Combobox.DisplayMember = "CPU"


	$obj = ""


	# END Populate CPU Dropdown


	


	# Populate Memory Dropdown


	$objMemoryArray = @()


	$I = 0


	


	foreach ($M in $Memory)


	{


		If ($M -ne "")


		{


			$obj = new-object System.Object


			$obj | add-member -type NoteProperty -name Memory -value $Memory[$I]


			$I++


			$objMemoryArray += $obj


		}


	}


	


	$VMMemory_Combobox.items.addrange($objMemoryArray)


	$VMMemory_Combobox.DisplayMember = "Memory"


	$obj = ""


	# END Populate Memory Dropdown


	


	# Populate Disk1 Dropdown


	$objDisk1Array = @()


	$I = 0


	


	foreach ($D in $Disk1)


	{


		If ($D -ne "")


		{


			$obj = new-object System.Object


			$obj | add-member -type NoteProperty -name Disk1 -value $Disk1[$I]


			$I++


			$objDisk1Array += $obj


		}


	}


	


	$VMDisk1_Combobox.items.addrange($objDisk1Array)


	$VMDisk1_Combobox.DisplayMember = "Disk1"


	$obj = ""


	# END Populate Disk1 Dropdown


	


	# Populate Disk2 Dropdown


	$objDisk2Array = @()


	$I = 0


	


	foreach ($D in $Disk2)


	{


		If ($D -ne "")


		{


			$obj = new-object System.Object


			$obj | add-member -type NoteProperty -name Disk2 -value $Disk2[$I]


			$I++


			$objDisk2Array += $obj


		}


	}


	


	$VMDisk2_Combobox.items.addrange($objDisk2Array)


	$VMDisk2_Combobox.DisplayMember = "Disk2"


	$obj = ""


	# END Populate Disk2 Dropdown


	


	# Populate Network Dropdown


	$global:objNetworkArray = @()


	$I = 0


	


	foreach ($N in $Network)


	{


		If ($N -ne "")


		{


			$obj = new-object System.Object


			$obj | add-member -type NoteProperty -name Network -value $Network[$I]


			$I++


			$global:objNetworkArray += $obj


		}


	}


	


	$VMNetwork_Combobox.items.addrange($global:objNetworkArray)


	$VMNetwork_Combobox.DisplayMember = "Network"


	$obj = ""


	# END Populate Network Dropdown


	


	# Populate Admin Network Dropdown


	$global:objAdminNetworkArray = @()


	$I = 0


	


	foreach ($N in $AdminNetwork)


	{


		If ($N -ne "")


		{


			$obj = new-object System.Object


			$obj | add-member -type NoteProperty -name Admin_Network -value $AdminNetwork[$I]


			$I++


			$global:objAdminNetworkArray += $obj


		}


	}


	


	$VMAdminNetwork_Combobox.items.addrange($global:objAdminNetworkArray)


	$VMAdminNetwork_Combobox.DisplayMember = "Admin_Network"


	$obj = ""


	# END Populate Admin Network Dropdown


	


	# Populate VM Folder Dropdown


	$global:objVMFolderArray = @()


	$I = 0


	


	foreach ($V in $VMFolder)


	{


		If ($V -ne "")


		{


			$obj = new-object System.Object


			$obj | add-member -type NoteProperty -name VMFolder -value $VMFolder[$I]


			$I++


			$global:objVMFolderArray += $obj


		}


	}


	


	$VMFolder_Combobox.items.addrange($global:objVMFolderArray)


	$VMFolder_Combobox.DisplayMember = "VMFolder"


	$obj = ""


	# END Populate VM Folder Dropdown


	


	# Populate Pre-Prod VM Folder Dropdown


	$global:objPreProdVMFolderArray = @()


	$I = 0


	


	foreach ($V in $PreProdVMFolder)


	{


		If ($V -ne "")


		{


			$obj = new-object System.Object


			$obj | add-member -type NoteProperty -name VMFolder -value $PreProdVMFolder[$I]


			$I++


			$global:objPreProdVMFolderArray += $obj


		}


	}


	# END Populate Pre-Prod VM Folder Dropdown


	


	# Populate DEV VM Folder Dropdown


	$global:objDevVMFolderArray = @()


	$I = 0


	


	foreach ($V in $DevVMFolder)


	{


		If ($V -ne "")


		{


			$obj = new-object System.Object


			$obj | add-member -type NoteProperty -name VMFolder -value $DevVMFolder[$I]


			$I++


			$global:objDevVMFolderArray += $obj


		}


	}


	# END Populate DEV VM Folder Dropdown


	


	# Populate Production-DR VM Folder Dropdown


	$global:objProdDRVMFolderArray = @()


	$I = 0


	


	foreach ($V in $ProdDRVMFolder)


	{


		If ($V -ne "")


		{


			$obj = new-object System.Object


			$obj | add-member -type NoteProperty -name VMFolder -value $ProdDRVMFolder[$I]


			$I++


			$global:objProdDRVMFolderArray += $obj


		}


	}


	# END Populate Production-DR VM Folder Dropdown


	


	# Populate DS Cluster Dropdown


	$global:objDSClusterArray = @()


	$I = 0


	


	foreach ($D in $DSCluster)


	{


		If ($D -ne "")


		{


			$obj = new-object System.Object


			$obj | add-member -type NoteProperty -name DSCluster -value $DSCluster[$I]


			$I++


			$global:objDSClusterArray += $obj


		}


	}


	


	$VMDSCluster_Combobox.items.addrange($global:objDSClusterArray)


	$VMDSCluster_Combobox.DisplayMember = "DSCluster"


	$obj = ""


	# END Populate DS Cluster Dropdown


	


	


	################################################################################################################


	#Pre-req check for PowerCLI


	################################################################################################################


	if (!(Test-Path “C:\Program Files (x86)\VMware\Infrastructure\PowerCLI\Scripts\Initialize-PowerCLIEnvironment.ps1”))


	{


		$VMOutput_Richtextbox.Text = $VMOutput_Richtextbox.Text.Clear


		$VMOutput_Richtextbox.ForeColor = [Drawing.Color]::Red


		$VMOutput_Richtextbox.AppendText("PowerCLI not found!")


		$VMOutput_Richtextbox.AppendText("`n")


		$VMOutput_Richtextbox.AppendText("PowerCLI required, please run from computer with it installed.")


	}


	


}



function Exit-Form


{


	$VM_Provisioning_Front_End.Close()


}



function Clear-Form {


	#Clearing all fields


	$VMHostname_Textbox.Text = $VMHostname_Textbox.Text.Clear


	$VMHostname_Textbox.BackColor = 'LightSalmon'


	$VMOS_Combobox.Items.Clear()


	$VMOS_Combobox.BackColor = 'LightSalmon'


	$VMEnvironment_Combobox.Items.Clear()


	$VMEnvironment_Combobox.BackColor = 'LightSalmon'


	$VMLocation_Combobox.Items.Clear()


	$VMLocation_Combobox.BackColor = 'LightSalmon'


	$VMCPU_Combobox.Items.Clear()


	$VMCPU_Combobox.BackColor = 'LightSalmon'


	$VMMemory_Combobox.Items.Clear()


	$VMMemory_Combobox.BackColor = 'LightSalmon'


	$VMDisk1_Combobox.Items.Clear()


	$VMDisk1_Combobox.BackColor = 'LightSalmon'


	$VMDisk2_Combobox.Items.Clear()


	$VMDisk2_Combobox.BackColor = 'LightSalmon'


	$VMNetwork_Combobox.Items.Clear()


	$VMNetwork_Combobox.BackColor = 'LightSalmon'


	$VMAdminNetwork_Combobox.Items.Clear()


	$VMAdminNetwork_Combobox.BackColor = 'LightSalmon'


	$VMFolder_Combobox.Items.Clear()


	$VMFolder_Combobox.BackColor = 'LightSalmon'


	$VMDSCluster_Combobox.Items.Clear()


	$VMDSCluster_Combobox.BackColor = 'LightSalmon'


	


	################


	#status message#


	################


	$VMOutput_Richtextbox.Text = $VMOutput_Richtextbox.Text.Clear


	$VMOutput_Richtextbox.ForeColor = [Drawing.Color]::Green


	$VMOutput_Richtextbox.AppendText("Form reset")


	


	Connect-SqlandPopulate


	


	#Renable locked combo boxes


	$VMDSCluster_Combobox.Enabled = $true


	$VMAdminNetwork_Combobox.Enabled = $true


}



function Harden-VM


{


	################################################################################################################


	#VM Hardening variables


	################################################################################################################


	#18.1.4 - Prevention of unsafe device disconnections


	$value1 = "isolation.device.connectable.disable"


	$value2 = "isolation.device.edit.disable"


	#18.1.5 - Prevent denial of service on the virtual disk changes


	$value3 = "isolation.tools.diskShrink.disable"


	$value4 = "isolation.tools.diskWiper.disable"


	#18.1.8 - Limit access to information from the physical host


	$value6 = "tools.guestlib.enableHostInfo"


	#18.1.9 - Disable copy/paste operations


	$value7 = "isolation.tools.copy.disable"


	$value8 = "isolation.tools.dnd.disable"


	$value9 = "isolation.tools.setGUIOptions.enable"


	$value10 = "isolation.tools.paste.disable"


	#18.1.10 - Limit the number of simultaneous connections


	$value11 = "RemoteDisplay.maxConnections"


	#18.1.11 - Limit maximum size of the configuration file


	$value12 = "tools.setInfo.sizeLimit"


	#18.1.12 - Disable BIOS BBS 


	$value13 = "isolation.bios.bbs.disable"


	#18.1.13 - Disable Guest Host Interaction Protocol Handler


	$value14 = "isolation.tools.ghi.protocolhandler.info.disable"


	#18.1.14 - Disable Unity Taskbar


	$value15 = "isolation.tools.unity.taskbar.disable"


	#18.1.15 - Disable Unity Active


	$value16 = "isolation.tools.unityActive.disable"


	#18.1.16 - Disable Unity Window Contents


	$value17 = "isolation.tools.unity.windowContents.disable"


	#18.1.17 - Disable Unity Push Update 


	$value18 = "isolation.tools.unity.push.update.disable"


	#18.1.18 - Disable Drag and Drop Version Get


	$value19 = "isolation.tools.vmxDnDVersionGet.disable"


	#18.1.19 - Disable Drag and Drop Version Set 


	$value20 = "isolation.tools.guestDnDVersionSet.disable"


	#18.1.20 - Disable Shell Action 


	$value21 = "isolation.ghi.host.shellAction.disable"


	#18.1.21 - Disable Request Disk Topology


	$value22 = "isolation.tools.dispTopoRequest.disable"


	#18.1.22 - Disable Trash Folder State


	$value23 = "isolation.tools.trashFolderState.disable"


	#18.1.23 - Disable Guest Host Interaction Tray Icon


	$value24 = "isolation.tools.ghi.trayicon.disable"


	#18.1.24 - Disable Unity


	$value25 = "isolation.tools.unity.disable"


	#18.1.25 - Disable Unity Interlock


	$value26 = "isolation.tools.unityInterlockOperation.disable"


	#18.1.26 - Disable GetCreds


	$value27 = "isolation.tools.getCreds.disable"


	#18.1.27 - Disable Host Guest File System Server 


	$value28 = "isolation.tools.hgfsServerSet.disable"


	#18.1.28 - Disable Host Guest Interaction Launch Menu 


	$value29 = "isolation.tools.ghi.launchmenu.change"


	#18.1.29 - Disable memSchedFakeSampleStats 


	$value30 = "isolation.tools.memSchedFakeSampleStats.disable"


	#18.1.30 - Limit number of VM log files and VM log file size


	$value31 = "log.keepOld"


	$value32 = "log.rotateSize"


	


	New-AdvancedSetting -Entity $VMHostname -Name $value1 -value $true -Confirm: $false -Force -WarningAction SilentlyContinue


	New-AdvancedSetting -Entity $VMHostname -Name $value2 -value $true -Confirm: $false -Force -WarningAction SilentlyContinue


	New-AdvancedSetting -Entity $VMHostname -Name $value3 -value $true -Confirm: $false -Force -WarningAction SilentlyContinue


	New-AdvancedSetting -Entity $VMHostname -Name $value4 -value $true -Confirm: $false -Force -WarningAction SilentlyContinue


	New-AdvancedSetting -Entity $VMHostname -Name $value6 -value $false -Confirm: $false -Force -WarningAction SilentlyContinue


	New-AdvancedSetting -Entity $VMHostname -Name $value7 -value $true -Confirm: $false -Force -WarningAction SilentlyContinue


	New-AdvancedSetting -Entity $VMHostname -Name $value8 -value $true -Confirm: $false -Force -WarningAction SilentlyContinue


	New-AdvancedSetting -Entity $VMHostname -Name $value9 -value $false -Confirm: $false -Force -WarningAction SilentlyContinue


	New-AdvancedSetting -Entity $VMHostname -Name $value10 -value $true -Confirm: $false -Force -WarningAction SilentlyContinue


	New-AdvancedSetting -Entity $VMHostname -Name $value11 -value 1 -Confirm: $false -Force -WarningAction SilentlyContinue


	New-AdvancedSetting -Entity $VMHostname -Name $value12 -value 1048576 -Confirm: $false -Force -WarningAction SilentlyContinue


	New-AdvancedSetting -Entity $VMHostname -Name $value13 -value $true -Confirm: $false -Force -WarningAction SilentlyContinue


	New-AdvancedSetting -Entity $VMHostname -Name $value14 -value $true -Confirm: $false -Force -WarningAction SilentlyContinue


	New-AdvancedSetting -Entity $VMHostname -Name $value15 -value $true -Confirm: $false -Force -WarningAction SilentlyContinue


	New-AdvancedSetting -Entity $VMHostname -Name $value16 -value $true -Confirm: $false -Force -WarningAction SilentlyContinue


	New-AdvancedSetting -Entity $VMHostname -Name $value17 -value $true -Confirm: $false -Force -WarningAction SilentlyContinue


	New-AdvancedSetting -Entity $VMHostname -Name $value18 -value $true -Confirm: $false -Force -WarningAction SilentlyContinue


	New-AdvancedSetting -Entity $VMHostname -Name $value19 -value $true -Confirm: $false -Force -WarningAction SilentlyContinue


	New-AdvancedSetting -Entity $VMHostname -Name $value20 -value $true -Confirm: $false -Force -WarningAction SilentlyContinue


	New-AdvancedSetting -Entity $VMHostname -Name $value21 -value $true -Confirm: $false -Force -WarningAction SilentlyContinue


	New-AdvancedSetting -Entity $VMHostname -Name $value22 -value $true -Confirm: $false -Force -WarningAction SilentlyContinue


	New-AdvancedSetting -Entity $VMHostname -Name $value23 -value $true -Confirm: $false -Force -WarningAction SilentlyContinue


	New-AdvancedSetting -Entity $VMHostname -Name $value24 -value $true -Confirm: $false -Force -WarningAction SilentlyContinue


	New-AdvancedSetting -Entity $VMHostname -Name $value25 -value $true -Confirm: $false -Force -WarningAction SilentlyContinue


	New-AdvancedSetting -Entity $VMHostname -Name $value26 -value $true -Confirm: $false -Force -WarningAction SilentlyContinue


	New-AdvancedSetting -Entity $VMHostname -Name $value27 -value $true -Confirm: $false -Force -WarningAction SilentlyContinue


	New-AdvancedSetting -Entity $VMHostname -Name $value28 -value $true -Confirm: $false -Force -WarningAction SilentlyContinue


	New-AdvancedSetting -Entity $VMHostname -Name $value29 -value $true -Confirm: $false -Force -WarningAction SilentlyContinue


	New-AdvancedSetting -Entity $VMHostname -Name $value30 -value $true -Confirm: $false -Force -WarningAction SilentlyContinue


	New-AdvancedSetting -Entity $VMHostname -Name $value31 -value 10 -Confirm: $false -Force -WarningAction SilentlyContinue


	New-AdvancedSetting -Entity $VMHostname -Name $value32 -value 1024000 -Confirm: $false -Force -WarningAction SilentlyContinue


	


	


}



##############################################################################################################


#Load Frontend


##############################################################################################################


$VM_Provisioning_Front_End_Load = {


	##############################################################################################################


	#Configure and Populate form


	##############################################################################################################


		Connect-SqlandPopulate


}



##############################################################################################################


#Clear Form button click


##############################################################################################################


$VMClearForm_Button_Click = {


	Clear-Form


}



##############################################################################################################


#Exit button click


##############################################################################################################


$VMExit_Button_Click = {


	Exit-Form


}



##############################################################################################################


#Create Shell button click


##############################################################################################################


$VMProvision_Button_Click = {


	


	#clear variables


	$Validated = $null


	


	# Clear Textbox


	$VMOutput_Richtextbox.Text = $VMOutput_Richtextbox.Text.Clear


	$VMOutput_Richtextbox.font = "lucida console"


	


	#validate form fields before submitting, does not proceed until form valadation succeeds


	if ($VMHostname_Textbox.BackColor -eq 'PaleGreen' -and $VMOS_Combobox.BackColor -eq 'PaleGreen'`


	-and $VMLocation_Combobox.BackColor -eq 'PaleGreen' -and $VMCPU_Combobox.BackColor -eq 'PaleGreen'`


	-and $VMMemory_Combobox.BackColor -eq 'PaleGreen' -and $VMDisk1_Combobox.BackColor -eq 'PaleGreen'`


	-and $VMDisk2_Combobox.BackColor -eq 'PaleGreen' -and $VMNetwork_Combobox.BackColor -eq 'PaleGreen'`


	-and $VMFolder_Combobox.BackColor -eq 'PaleGreen' -and $VMDSCluster_Combobox.BackColor -eq 'PaleGreen'`


	-and $VMAdminNetwork_Combobox.BackColor -eq 'PaleGreen' -and $VMEnvironment_Combobox.BackColor -eq 'PaleGreen')


	{ $Validated = "True" }


	else { $Validated = "False" }


	


	#continues shell creation if validation above passes with "True"


	if ($Validated -eq "True")


	{


		#reset variable


		$build = $null


		


		#apply username/if entered


		$ADUserName = $ADUserName_Textbox.Text


		$ADUserPassword = $ADUserPassword_Textbox.Text


		


		#check hostname entered is a vaild value, currently 


		$VMHostname = $VMHostname_Textbox.Text


		if ($VMHostname -match "" -or $VMHostname -match "" -or $VMHostname -match "" -or $VMHostname -match "")


		{


			


			#loading PowerCLI


			#pssnappin deprecated


			#Add-PSSnapin vmware.vimautomation.core


			if (!(Get-Module -Name VMware.VimAutomation.Core -ErrorAction SilentlyContinue))


			{


				#. “C:\Program Files (x86)\VMware\Infrastructure\vSphere PowerCLI\Scripts\Initialize-PowerCLIEnvironment.ps1”


				Import-Module -Name VMware.VimAutomation.Core


				Import-Module -Name VMware.VimAutomation.Vds


			}


			Set-PowerCLIConfiguration -InvalidCertificateAction ignore -confirm:$false


			Set-PowerCLIConfiguration -DisplayDeprecationWarnings $false -Confirm:$false


			


			################


			#status message#


			################


			$VMOutput_Richtextbox.ForeColor = [Drawing.Color]::Green


			$VMOutput_Richtextbox.AppendText("Connecting to vCenter")


			


			# Populate Variables based on form input & connect to relevant environment and VC


			$VMHostname = $VMHostname_Textbox.Text.ToUpper()


			$VMEnvironment = $VMEnvironment_Combobox.Text


			$VMLocation = $VMLocation_Combobox.Text


				Try


				{


				#Connecting to vCenter server based on location entered in form, also checks to see if password has been entered into form


				if ($ADUserPassword -ne "")


				{


					if ($VMLocation -match "Site") { Connect-VIServer vcenter1 -User $ADUserName -Password $ADUserPassword -Force -ErrorAction Stop -WarningAction SilentlyContinue }


					elseif ($VMLocation -match "Site") { Connect-VIServer vcenter2 -User $ADUserName -Password $ADUserPassword -Force -ErrorAction Stop -WarningAction SilentlyContinue }


					elseif ($VMLocation -match "Site" -and $VMEnvironment -match "Development") { Connect-VIServer vcenter3 -User $ADUserName -Password $ADUserPassword -Force -ErrorAction Stop -WarningAction SilentlyContinue }


					elseif ($VMLocation -match "Site" -and $VMEnvironment -match "Pre-Production") { Connect-VIServer vcenter4 -User $ADUserName -Password $ADUserPassword -Force -ErrorAction Stop -WarningAction SilentlyContinue }


					elseif ($VMLocation -match "Site" -and $VMEnvironment -match "Production-DR") { Connect-VIServer vcenter5 -User $ADUserName -Password $ADUserPassword -Force -ErrorAction Stop -WarningAction SilentlyContinue }


				}


				else


				{


					if ($VMLocation -match "Site") { Connect-VIServer vcenter1 -Force -ErrorAction Stop -WarningAction SilentlyContinue }


					elseif ($VMLocation -match "Site") { Connect-VIServer vcenter2 -Force -ErrorAction Stop -WarningAction SilentlyContinue }


					elseif ($VMLocation -match "Site" -and $VMEnvironment -match "Development") { Connect-VIServer vcenter3 -Force -ErrorAction Stop -WarningAction SilentlyContinue }


					elseif ($VMLocation -match "Site" -and $VMEnvironment -match "Pre-Production") { Connect-VIServer vcenter4 -Force -ErrorAction Stop -WarningAction SilentlyContinue }


					elseif ($VMLocation -match "Site" -and $VMEnvironment -match "Production-DR") { Connect-VIServer vcenter5 -Force -ErrorAction Stop -WarningAction SilentlyContinue }


				}


			}


			Catch


			{


				###############


				#Error message#


				###############


				$VMOutput_Richtextbox.Text = $VMOutput_Richtextbox.Text.Clear


				$VMOutput_Richtextbox.ForeColor = [Drawing.Color]::Red


				$ErrorMessage = $_.Exception.Message


				$VMOutput_Richtextbox.Text = $ErrorMessage


				Return


			}


			$VMOS = $VMOS_Combobox.Text


				#Passing OS switch to VM script depending on selection


				switch ($VMOS)


				{


					"Windows 2012 R2" { $OS = "windows8Server64Guest" }


					"Windows 2016" { $OS = "windows9Server64Guest" }


				}


			$VMLocation = $VMLocation_Combobox.Text


			$VMEnvironment = $VMEnvironment_Combobox.Text


			$VMCPU = $VMCPU_Combobox.Text


			$VMMemory = $VMMemory_Combobox.Text


			$VMDisk1 = $VMDisk1_Combobox.Text


			$VMDisk2 = $VMDisk2_Combobox.Text


			$VMNetwork = $VMNetwork_Combobox.Text


			$VMAdminNetwork = $VMAdminNetwork_Combobox.Text


			#Prod folders


			$VMFolder = $VMFolder_Combobox.Text


				#Convert selected folder for VM script


				if ($VMLocation -match "Site")


				{


					switch ($VMFolder)


					{


						"Unspecified" { $VMFolder = "" }


						"Unspecified" { $VMFolder = (Get-Folder -Id "Folder-group-v357") }


						"Unspecified" { $VMFolder = (Get-Folder -Id "Folder-group-v621") }


						"Unspecified" { $VMFolder = (Get-Folder -Id "Folder-group-v1055") }


						"Unspecified" { $VMFolder = (Get-Folder -Id "Folder-group-v18231") }


						"Unspecified" { $VMFolder = (Get-Folder -Id "Folder-group-v226") }


						"Unspecified" { $VMFolder = (Get-Folder -Id "Folder-group-v352") }


						"Unspecified" { $VMFolder = (Get-Folder -Id "Folder-group-v1950") }


				}


			}


			if ($VMLocation -match "Site")


				{


					switch ($VMFolder)


					{


						"Unspecified" { $VMFolder = "" }


						"Unspecified" { $VMFolder = (Get-Folder -Id "Folder-group-v141") }


						"Unspecified" { $VMFolder = (Get-Folder -Id "Folder-group-v325") }


						"Unspecified" { $VMFolder = (Get-Folder -Id "Folder-group-v903") }


						"Unspecified" { $VMFolder = (Get-Folder -Id "Folder-group-v7712") }


						"Unspecified" { $VMFolder = (Get-Folder -Id "Folder-group-v98") }


						"Unspecified" { $VMFolder = (Get-Folder -Id "Folder-group-v142") }


						"Unspecified" { $VMFolder = (Get-Folder -Id "Folder-group-v1799") }


					}


				}


			#Pre-Prod folders


			$VMPreProdFolder = $VMFolder_Combobox.Text


			#Convert selected folder for VM script


			if ($VMLocation -match "Site" -and $VMEnvironment -match "Pre-Production")


			{


				switch ($VMPreProdFolder)


				{


					"Unspecified" { $VMFolder = "" }


					"Unspecified" { $VMFolder = (Get-Folder -Id "Folder-group-v107") }


					"Unspecified" { $VMFolder = (Get-Folder -Id "Folder-group-v108") }


					"Unspecified" { $VMFolder = (Get-Folder -Id "Folder-group-v112") }


					"Unspecified" { $VMFolder = (Get-Folder -Id "Folder-group-v111") }


					"Unspecified" { $VMFolder = (Get-Folder -Id "Folder-group-v1165") }


					"Unspecified" { $VMFolder = (Get-Folder -Id "Folder-group-v84") }


					"Unspecified" { $VMFolder = (Get-Folder -Id "Folder-group-v109") }


				}


			}


			#Dev folders


			$VMDevFolder = $VMFolder_Combobox.Text


				#Convert selected folder for VM script


				if ($VMLocation -match "Site" -and $VMEnvironment -match "Development")


				{


					switch ($VMDevFolder)


					{


						"Unspecified" { $VMFolder = "" }


						"Unspecified" { $VMFolder = (Get-Folder -Id "Folder-group-v114") }


						"Unspecified" { $VMFolder = (Get-Folder -Id "Folder-group-v115") }


						"Unspecified" { $VMFolder = (Get-Folder -Id "Folder-group-v822") }


						"Unspecified" { $VMFolder = (Get-Folder -Id "Folder-group-v116") }


						"Unspecified" { $VMFolder = (Get-Folder -Id "Folder-group-v117") }


						"Unspecified" { $VMFolder = (Get-Folder -Id "Folder-group-v118") }


						"Unspecified" { $VMFolder = (Get-Folder -Id "Folder-group-v119") }


						"Unspecified" { $VMFolder = (Get-Folder -Id "Folder-group-v16141") }


						"Unspecified" { $VMFolder = (Get-Folder -Id "Folder-group-v257") }


						"Unspecified" { $VMFolder = (Get-Folder -Id "Folder-group-v120") }


						"Unspecified" { $VMFolder = (Get-Folder -Id "Folder-group-v123") }


						"Unspecified" { $VMFolder = (Get-Folder -Id "Folder-group-v121") }


					}


				}


			#Production-DR folders


			$VMProdDRFolder = $VMFolder_Combobox.Text


			#Convert selected folder for VM script


			if ($VMLocation -match "Site" -and $VMEnvironment -match "Production-DR")


			{


				switch ($VMProdDRFolder)


				{


					"Unspecified" { $VMFolder = "" }


					"Unspecified" { $VMFolder = (Get-Folder -Id "Folder-group-v256") }


					"Unspecified" { $VMFolder = (Get-Folder -Id "Folder-group-v258") }


					"Unspecified" { $VMFolder = (Get-Folder -Id "Folder-group-v261") }


					"Unspecified" { $VMFolder = (Get-Folder -Id "Folder-group-v2328") }


					"Unspecified" { $VMFolder = (Get-Folder -Id "Folder-group-v260") }


					"Unspecified" { $VMFolder = (Get-Folder -Id "Folder-group-v581") }


					"Unspecified" { $VMFolder = (Get-Folder -Id "Folder-group-v2008") }


					"Unspecified" { $VMFolder = (Get-Folder -Id "Folder-group-v2512") }


					"Unspecified" { $VMFolder = (Get-Folder -Id "Folder-group-v2470") }


				}


			}


			$VMDSCluster = $VMDSCluster_Combobox.Text


			if ($VMDSCluster -match "Servers")


				{$VMDSCluster = "StoragePod-group-p255"}


				


				################


				#status message#


				################


				$VMOutput_Richtextbox.AppendText("`n")


				$VMOutput_Richtextbox.AppendText("Connected to vCenter")


				$VMOutput_Richtextbox.AppendText("`n")


				$VMOutput_Richtextbox.AppendText("Creating Shell")


			


			##Provision VM in relevant DC


			Try


			{


				##Provision VM Site


				if ($VMLocation -eq "Site")


				{


					################


					#status message#


					################


					$VMOutput_Richtextbox.AppendText("`n")


					$VMOutput_Richtextbox.AppendText("Creating Shell in Site")


					$cluster1 = Get-Cluster "Cluster_Name" | Get-VMHost | Where-Object { $_.ConnectionState -eq 'Connected' } | Sort-Object MemoryUsageGB


					New-VM -Name $VMHostname -Location $VMFolder -VMHost $cluster1[0] -CD -Datastore $VMDSCluster -Version v10 -GuestId $OS -NumCpu $VMCPU -MemoryGB $VMMemory -DiskGB $VMDisk1, $VMDisk2 -DiskStorageFormat Thick -Portgroup (Get-VDPortGroup -Name $VMNetwork), (Get-VDPortGroup -Name $VMAdminNetwork) -ErrorAction Stop


					Get-VM -Name $VMHostname | Get-NetworkAdapter | Set-NetworkAdapter -Type Vmxnet3 -Confirm:$false -ErrorAction Stop


					$vmmac = Get-VM -Name $VMHostname | Get-NetworkAdapter


					$VMOutput_Richtextbox.AppendText("`n")


					$VMOutput_Richtextbox.AppendText("MAC: $($vmmac[0].MacAddress)")


					#Hardening


					$VMOutput_Richtextbox.AppendText("`n")


					$VMOutput_Richtextbox.AppendText("Hardening VM")


					Harden-VM


					Disconnect-VIServer -Server * -Force -Confirm:$false -ErrorAction SilentlyContinue -WarningAction SilentlyContinue


				}


				


				##Provision VM Site


				if ($VMLocation -eq "Site")


				{


					################


					#status message#


					################


					$VMOutput_Richtextbox.AppendText("`n")


					$VMOutput_Richtextbox.AppendText("Creating Shell in Site")


					$cluster2 = Get-Cluster "Cluster_Name" | Get-VMHost | Where-Object { $_.ConnectionState -eq 'Connected' } | Sort-Object MemoryUsageGB


					New-VM -Name $VMHostname -Location $VMFolder -VMHost $cluster2[0] -CD -Datastore $VMDSCluster -Version v10 -GuestId $OS -NumCpu $VMCPU -MemoryGB $VMMemory -DiskGB $VMDisk1, $VMDisk2 -DiskStorageFormat Thick -Portgroup (Get-VDPortGroup -Name $VMNetwork), (Get-VDPortGroup -Name $VMAdminNetwork) -ErrorAction Stop


					Get-VM -Name $VMHostname | Get-NetworkAdapter | Set-NetworkAdapter -Type Vmxnet3 -Confirm:$false


					$vmmac = Get-VM -Name $VMHostname | Get-NetworkAdapter


					$VMOutput_Richtextbox.AppendText("`n")


					$VMOutput_Richtextbox.AppendText("MAC: $($vmmac[0].MacAddress)")


					#Hardening


					$VMOutput_Richtextbox.AppendText("`n")


					$VMOutput_Richtextbox.AppendText("Hardening VM")


					Harden-VM


					Disconnect-VIServer -Server * -Force -Confirm:$false -ErrorAction SilentlyContinue -WarningAction SilentlyContinue


				}


				


				##Provision VM Site Dev


				if ($VMLocation -eq "Site" -and $VMEnvironment -eq "Development")


				{


					################


					#status message#


					################


					$VMOutput_Richtextbox.AppendText("`n")


					$VMOutput_Richtextbox.AppendText("Creating Shell in Site")


					$cluster1 = Get-Cluster "Cluster_Name" | Get-VMHost | Where-Object { $_.ConnectionState -eq 'Connected' } | Sort-Object MemoryUsageGB


					New-VM -Name $VMHostname -Location $VMFolder -VMHost $cluster1[0] -CD -Datastore $VMDSCluster -Version v10 -GuestId $OS -NumCpu $VMCPU -MemoryGB $VMMemory -DiskGB $VMDisk1, $VMDisk2 -DiskStorageFormat Thick -Portgroup (Get-VDPortGroup -Name $VMNetwork) -ErrorAction Stop


					Get-VM -Name $VMHostname | Get-NetworkAdapter | Set-NetworkAdapter -Type Vmxnet3 -Confirm:$false


					$vmmac = Get-VM -Name $VMHostname | Get-NetworkAdapter


					$VMOutput_Richtextbox.AppendText("`n")


					$VMOutput_Richtextbox.AppendText("MAC: $($vmmac[0].MacAddress)")


					#Hardening


					$VMOutput_Richtextbox.AppendText("`n")


					$VMOutput_Richtextbox.AppendText("Hardening VM")


					Harden-VM


					Disconnect-VIServer -Server * -Force -Confirm:$false -ErrorAction SilentlyContinue -WarningAction SilentlyContinue


				}


				


				##Provision VM Site Pre-Production


				if ($VMLocation -eq "Site" -and $VMEnvironment -eq "Pre-Production")


				{


					################


					#status message#


					################


					$VMOutput_Richtextbox.AppendText("`n")


					$VMOutput_Richtextbox.AppendText("Creating Shell in Site")


					$cluster1 = Get-Cluster "Cluster_Name" | Get-VMHost | Where-Object { $_.ConnectionState -eq 'Connected' } | Sort-Object MemoryUsageGB


					New-VM -Name $VMHostname -Location $VMFolder -VMHost $cluster1[0] -CD -Datastore $VMDSCluster -Version v10 -GuestId $OS -NumCpu $VMCPU -MemoryGB $VMMemory -DiskGB $VMDisk1, $VMDisk2 -DiskStorageFormat Thick -Portgroup (Get-VDPortGroup -Name $VMNetwork), (Get-VDPortGroup -Name $VMAdminNetwork) -ErrorAction Stop


					Get-VM -Name $VMHostname | Get-NetworkAdapter | Set-NetworkAdapter -Type Vmxnet3 -Confirm:$false


					$vmmac = Get-VM -Name $VMHostname | Get-NetworkAdapter


					$VMOutput_Richtextbox.AppendText("`n")


					$VMOutput_Richtextbox.AppendText("MAC: $($vmmac[0].MacAddress)")


					#Hardening


					$VMOutput_Richtextbox.AppendText("`n")


					$VMOutput_Richtextbox.AppendText("Hardening VM")


					Harden-VM


					Disconnect-VIServer -Server * -Force -Confirm:$false -ErrorAction SilentlyContinue -WarningAction SilentlyContinue


				}


				


				##Provision VM Site Production-DR


				if ($VMLocation -eq "Site" -and $VMEnvironment -eq "Production-DR")


				{


					################


					#status message#


					################


					$VMOutput_Richtextbox.AppendText("`n")


					$VMOutput_Richtextbox.AppendText("Creating Shell in Site")


					$cluster1 = Get-Cluster "Cluster_Name" | Get-VMHost | Where-Object { $_.ConnectionState -eq 'Connected' } | Sort-Object MemoryUsageGB


					New-VM -Name $VMHostname -Location $VMFolder -VMHost $cluster1[0] -CD -Datastore (Get-DatastoreCluster -Id $VMDSCluster) -Version v10 -GuestId $OS -NumCpu $VMCPU -MemoryGB $VMMemory -DiskGB $VMDisk1, $VMDisk2 -DiskStorageFormat Thick -Portgroup (Get-VDPortGroup -Name $VMNetwork), (Get-VDPortGroup -Name $VMAdminNetwork | Where-Object { $_.Datacenter -match 'DC_Site_PRD_1' }) -ErrorAction Stop


					Get-VM -Name $VMHostname | Get-NetworkAdapter | Set-NetworkAdapter -Type Vmxnet3 -Confirm:$false


					$vmmac = Get-VM -Name $VMHostname | Get-NetworkAdapter


					$VMOutput_Richtextbox.AppendText("`n")


					$VMOutput_Richtextbox.AppendText("MAC: $($vmmac[0].MacAddress)")


					#Hardening


					$VMOutput_Richtextbox.AppendText("`n")


					$VMOutput_Richtextbox.AppendText("Hardening VM")


					Harden-VM


					Disconnect-VIServer -Server * -Force -Confirm:$false -ErrorAction SilentlyContinue -WarningAction SilentlyContinue


				}


			}


			Catch


			{


				###############


				#Error message#


				###############


				$VMOutput_Richtextbox.Text = $VMOutput_Richtextbox.Text.Clear


				$VMOutput_Richtextbox.ForeColor = [Drawing.color]::Red


				$ErrorMessage = $_.Exception.Message


				$VMOutput_Richtextbox.Text = $ErrorMessage


				Return


			}


			


				################


				#status message#


				################


				$VMOutput_Richtextbox.AppendText("`n")


				$VMOutput_Richtextbox.AppendText("Shell created successfully")


			}


			else


			{


				################


				#status message#


				################


				$VMOutput_Richtextbox.Text = $VMOutput_Richtextbox.Text.Clear


				$VMOutput_Richtextbox.ForeColor = [Drawing.Color]::Red


				$VMOutput_Richtextbox.AppendText("Hostname should contain xx, xxx or xxxx. Please amend and try again")


			}


		}


		elseif ($Validated -eq "False")


		{


			################


			#status message#


			################


			$VMOutput_Richtextbox.Text = $VMOutput_Richtextbox.Text.Clear


			$VMOutput_Richtextbox.ForeColor = [Drawing.Color]::Red


			$VMOutput_Richtextbox.AppendText("Validation failed, please check fields and resubmit")


		}


	}



##############################################################################################################


#Colour format cells and create cell dependencies(Events)


##############################################################################################################


#Hostname value entered


$VMHostname_Textbox_TextChanged = {


	$VMHostname_Textbox.BackColor = 'PaleGreen'


}



#OS value selected


$VMOS_Combobox_TextChanged = {


	$VMOS_Combobox.BackColor = 'PaleGreen'


}



#Environment value selected


$VMEnvironment_Combobox_TextChanged = {


	$VMEnvironment_Combobox.BackColor = 'PaleGreen'


	$EnvTextCapture = $VMEnvironment_Combobox.Items


	switch ($VMEnvironment_Combobox.Text)


	{


		"Production" {


			#restrict location for Production


			$VMLocation_Combobox.Items.Clear()


			$VMLocation_Combobox.BackColor = 'LightSalmon'


			$VMLocation_Combobox.Items.AddRange($global:objLocationArray)


			$VMLocation_Combobox.Items.RemoveAt("2")


			#restrict networks for Production


			$VMNetwork_Combobox.Items.Clear()


			$VMNetwork_Combobox.Items.AddRange($global:objNetworkArray)


			$VMNetwork_Combobox.Items.RemoveAt("1")


			$VMNetwork_Combobox.Items.RemoveAt("1")


			$VMNetwork_Combobox.Items.RemoveAt("1")


			$VMNetwork_Combobox.SelectedIndex = "0"


			#restrict admin networks for Production


			$VMAdminNetwork_Combobox.Items.Clear()


			$VMAdminNetwork_Combobox.Items.AddRange($global:objAdminNetworkArray)


			$VMAdminNetwork_Combobox.Items.RemoveAt("1")


			$VMAdminNetwork_Combobox.SelectedIndex = "0"


			#restrict datastores for Production


			$VMDSCluster_Combobox.Items.Clear()


			$VMDSCluster_Combobox.Items.AddRange($global:objDSClusterArray)


			$VMDSCluster_Combobox.Items.RemoveAt("2")


			$VMDSCluster_Combobox.Items.RemoveAt("2")


			$VMDSCluster_Combobox.SelectedIndex = "0"


			#restrict folders for Production


			$VMFolder_Combobox.Items.Clear()


			$VMFolder_Combobox.BackColor = 'LightSalmon'


			$VMFolder_Combobox.Items.AddRange($global:objVMFolderArray)


		}


		


		"Pre-Production" {


			#restrict location for Pre-Production


			$VMLocation_Combobox.Items.Clear()


			$VMLocation_Combobox.Items.AddRange($global:objLocationArray)


			$VMLocation_Combobox.Items.RemoveAt("0")


			$VMLocation_Combobox.Items.RemoveAt("0")


			$VMLocation_Combobox.SelectedIndex = "0"


			#restrict network for Pre-Production


			$VMNetwork_Combobox.Items.Clear()


			$VMNetwork_Combobox.Items.AddRange($global:objNetworkArray)


			$VMNetwork_Combobox.Items.RemoveAt("0")


			$VMNetwork_Combobox.Items.RemoveAt("0")


			$VMNetwork_Combobox.Items.RemoveAt("1")


			$VMNetwork_Combobox.SelectedIndex = "0"


			#restrict admin networks for Pre-Production


			$VMAdminNetwork_Combobox.Items.Clear()


			$VMAdminNetwork_Combobox.Items.AddRange($global:objAdminNetworkArray)


			$VMAdminNetwork_Combobox.Items.RemoveAt("0")


			$VMAdminNetwork_Combobox.SelectedIndex = "0"


			#restrict datastores for Pre-Production


			$VMDSCluster_Combobox.Items.Clear()


			$VMDSCluster_Combobox.Items.AddRange($global:objDSClusterArray)


			$VMDSCluster_Combobox.Items.RemoveAt("1")


			$VMDSCluster_Combobox.Items.RemoveAt("1")


			$VMDSCluster_Combobox.Items.RemoveAt("1")


			$VMDSCluster_Combobox.SelectedIndex = "0"


			#Restrict folders for Pre-Prod


			$VMFolder_Combobox.Items.Clear()


			$VMFolder_Combobox.BackColor = 'LightSalmon'


			$VMFolder_Combobox.Items.AddRange($global:objPreProdVMFolderArray)


						 }


		


		"Development" {


			#restrict location for Development


			$VMLocation_Combobox.Items.Clear()


			$VMLocation_Combobox.Items.AddRange($global:objLocationArray)


			$VMLocation_Combobox.Items.RemoveAt("0")


			$VMLocation_Combobox.Items.RemoveAt("0")


			$VMLocation_Combobox.SelectedIndex = "0"


			#restrict network for Development


			$VMNetwork_Combobox.Items.Clear()


			$VMNetwork_Combobox.Items.AddRange($global:objNetworkArray)


			$VMNetwork_Combobox.Items.RemoveAt("0")


			$VMNetwork_Combobox.Items.RemoveAt("1")


			$VMNetwork_Combobox.Items.RemoveAt("1")


			$VMNetwork_Combobox.SelectedIndex = "0"


			$VMAdminNetwork_Combobox.Items.Clear()


			#restrict datastores for Development


			$VMDSCluster_Combobox.Items.Clear()


			$VMDSCluster_Combobox.Items.AddRange($global:objDSClusterArray)


			$VMDSCluster_Combobox.Items.RemoveAt("0")


			$VMDSCluster_Combobox.Items.RemoveAt("0")


			$VMDSCluster_Combobox.Items.RemoveAt("1")


			$VMDSCluster_Combobox.SelectedIndex = "0"


			#restrict folders for Development


			$VMFolder_Combobox.Items.Clear()


			$VMFolder_Combobox.BackColor = 'LightSalmon'


			$VMFolder_Combobox.Items.AddRange($global:objDevVMFolderArray)


			   			}


		


		"Production-DR" {


			#restrict location for Production-DR


			$VMLocation_Combobox.Items.Clear()


			$VMLocation_Combobox.Items.AddRange($global:objLocationArray)


			$VMLocation_Combobox.Items.RemoveAt("0")


			$VMLocation_Combobox.Items.RemoveAt("0")


			$VMLocation_Combobox.SelectedIndex = "0"


			#restrict network for Production-DR


			$VMNetwork_Combobox.Items.Clear()


			$VMNetwork_Combobox.Items.AddRange($global:objNetworkArray)


			$VMNetwork_Combobox.Items.RemoveAt("0")


			$VMNetwork_Combobox.Items.RemoveAt("1")


			$VMNetwork_Combobox.Items.RemoveAt("0")


			$VMNetwork_Combobox.SelectedIndex = "0"


			$VMAdminNetwork_Combobox.Items.Clear()


			$VMAdminNetwork_Combobox.Items.AddRange($global:objAdminNetworkArray)


			$VMAdminNetwork_Combobox.Items.RemoveAt("0")


			$VMAdminNetwork_Combobox.SelectedIndex = "0"


			#restrict datastores for Production-DR


			$VMDSCluster_Combobox.Items.Clear()


			$VMDSCluster_Combobox.Items.AddRange($global:objDSClusterArray)


			$VMDSCluster_Combobox.Items.RemoveAt("0")


			$VMDSCluster_Combobox.Items.RemoveAt("0")


			$VMDSCluster_Combobox.Items.RemoveAt("0")


			$VMDSCluster_Combobox.SelectedIndex = "0"


			#restrict folders for Production-DR


			$VMFolder_Combobox.Items.Clear()


			$VMFolder_Combobox.BackColor = 'LightSalmon'


			$VMFolder_Combobox.Items.AddRange($global:objProdDRVMFolderArray)


		}


	}


}



#Location value selected


$VMLocation_Combobox_TextChanged = {


	$VMLocation_Combobox.BackColor = 'PaleGreen'


	switch ($VMLocation_Combobox.Text)


	{


		"Site" {


			#restrict datastores for DC


			$VMDSCluster_Combobox.Items.Clear()


			$VMDSCluster_Combobox.Items.AddRange($global:objDSClusterArray)


			$VMDSCluster_Combobox.Items.RemoveAt("1")


			$VMDSCluster_Combobox.Items.RemoveAt("1")


			$VMDSCluster_Combobox.Items.RemoveAt("1")


			$VMDSCluster_Combobox.SelectedIndex = "0"


			#restrict networks for DC


			$VMNetwork_Combobox.Items.Clear()


			$VMNetwork_Combobox.Items.AddRange($global:objNetworkArray)


			$VMNetwork_Combobox.Items.RemoveAt("1")


			$VMNetwork_Combobox.Items.RemoveAt("1")


			$VMNetwork_Combobox.Items.RemoveAt("1")


			$VMNetwork_Combobox.SelectedIndex = "0"


			#restrict folders for Production


			$VMFolder_Combobox.Items.Clear()


			$VMFolder_Combobox.BackColor = 'LightSalmon'


			$VMFolder_Combobox.Items.AddRange($global:objVMFolderArray)					


			}


		"Site" {


			#restrict datastores for DC


			$VMDSCluster_Combobox.Items.Clear()


			$VMDSCluster_Combobox.Items.AddRange($global:objDSClusterArray)


			$VMDSCluster_Combobox.Items.RemoveAt("0")


			$VMDSCluster_Combobox.Items.RemoveAt("1")


			$VMDSCluster_Combobox.Items.RemoveAt("1")


			$VMDSCluster_Combobox.SelectedIndex = "0"


			#restrict networks for DC


			$VMNetwork_Combobox.Items.Clear()


			$VMNetwork_Combobox.Items.AddRange($global:objNetworkArray)


			$VMNetwork_Combobox.Items.RemoveAt("1")


			$VMNetwork_Combobox.Items.RemoveAt("1")


			$VMNetwork_Combobox.Items.RemoveAt("1")


			$VMNetwork_Combobox.SelectedIndex = "0"


			#restrict folders for Production


			$VMFolder_Combobox.Items.Clear()


			$VMFolder_Combobox.BackColor = 'LightSalmon'


			$VMFolder_Combobox.Items.AddRange($global:objVMFolderArray)


			}


		"Site" {


			#restrict datastores for DC


			$VMDSCluster_Combobox.Items.Clear()


			$VMDSCluster_Combobox.Items.AddRange($global:objDSClusterArray)


			$VMDSCluster_Combobox.Items.RemoveAt("1")


			$VMDSCluster_Combobox.Items.RemoveAt("1")


			$VMDSCluster_Combobox.SelectedIndex = "0"


			#restrict networks for DC


			$VMNetwork_Combobox.Items.Clear()


			$VMNetwork_Combobox.Items.AddRange($global:objNetworkArray)


			$VMNetwork_Combobox.Items.RemoveAt("0")


			$VMNetwork_Combobox.SelectedIndex = "0"


			#restrict admin networks for DC


			$VMAdminNetwork_Combobox.Items.Clear()


			$VMAdminNetwork_Combobox.Text = "Delete this at some point"


			#restrict folders for DC


			$VMFolder_Combobox.Items.Clear()


			$VMFolder_Combobox.BackColor = 'LightSalmon'


			$VMFolder_Combobox.Items.AddRange($global:objDevVMFolderArray)


		}


	}


	


}



#CPU value selected


$VMCPU_Combobox_TextChanged = {


	$VMCPU_Combobox.BackColor = 'PaleGreen'


}



#Memory value selected


$VMMemory_Combobox_TextChanged = {


	$VMMemory_Combobox.BackColor = 'PaleGreen'


}



#Disk1 value selected


$VMDisk1_Combobox_TextChanged = {


	$VMDisk1_Combobox.BackColor = 'PaleGreen'


}



#Disk2 value selected


$VMDisk2_Combobox_TextChanged = {


	$VMDisk2_Combobox.BackColor = 'PaleGreen'


}



#Network value selected


$VMNetwork_Combobox_TextChanged = {


	$VMNetwork_Combobox.BackColor = 'PaleGreen'


}



#Admin Network value selected


$VMAdminNetwork_Combobox_TextChanged = {


	$VMAdminNetwork_Combobox.BackColor = 'PaleGreen'


}



#Folder value selected


$VMFolder_Combobox_TextChanged = {


	$VMFolder_Combobox.BackColor = 'PaleGreen'


}



#DSCluster value selected


$VMDSCluster_Combobox_TextChanged = {


	$VMDSCluster_Combobox.BackColor = 'PaleGreen'


}



############################################


#Username/Password fields


############################################


$ADUserName_Textbox.Text = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name


$ADUserName_Textbox.BackColor = 'PaleGreen'



$ADUserName_Textbox_TextChanged = {


	$ADUserName_Textbox.BackColor = 'PaleGreen'


}



$ADUserPassword_Textbox_TextChanged = {


	$ADUserPassword_Textbox.BackColor = 'PaleGreen'


	$ADUserPassword_Textbox.UseSystemPasswordChar = $true


} 
