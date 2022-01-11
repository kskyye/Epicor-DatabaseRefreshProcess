# >(^_^)>  >(^_^)> --- SCRIPT SCOPE, PURPOSE & DETAILS --- <(^_^)<  <(^_^)<
#region Scope and Purpose

#Name:             Epicor Environment Refresh Partial Automation Script v1.1
#Date:             2018 06 08 
#Author:           Kristin Anderson (all code written by me unless noted)
#Purpose:          To save time by automating environment refresh tasks assigned to Infrastructure by minimizing user iteration and 
#                  eliminating need for IIS/MMC/etc GUIs.

#Methods of Use:   Unattended - script can be setup to run as a scheduled task, SSIS external application, or other similar ways by using the command line below:
#                     %windir%\system32\WindowsPowerShell\v1.0\powershell.exe -executionpolicy bypass -nologo -file "c:\scripts\epicor\TEST-Automate-EpicorRefresh.ps1" -targetEnvName DEV 
#                     The above example would run refresh tasks against the DEV environment only. (NOTE: this script will not automate the entire process due to there being several
#                     steps that must be done from Epicor GUI's)

#                  Manual - script can be run manually wherein which the user is presented a menu of options including running refresh tasks on individual environments, an option to run tasks
#                    on all lower environments and a task to restart services post-database refresh for all environments.
#>
#endregion
### <(^_^)> --- END SCOPE & PURPOSE REGION --- <(^_^)>


### >(^_^)>  >(^_^)> --- MODULES --- <(^_^)<  <(^_^)<
#region Modules
Import-Module WebAdministration
#endregion
### <(^_^)> --- END MODULES REGION --- <(^_^)>


### >(^_^)>  >(^_^)> --- VARIABLES --- <(^_^)<  <(^_^)<
#region Variables

#Set default value of flag to determine if current environment refresh has been completed
$refreshComplete = "N"

#Date used for folder and log file timestamping
$date = Get-Date -format yyyy.M.d

#Define variables including servers on which certain components live (i.e. the Integration Services)
<#1 Epicor DEV
Foundation DEV -> Epicor DEV#>
$epicorIntSvcDEV = 'DEV-FLWA1.sufs.local'
$appPoolDev = 'EpicorDev10'

<#2 Epicor TEST
Foundation UAT -> Epicor TEST#>
$epicorIntSvcTEST = 'UAT-FLSV1.sufs.local'
$appPoolTest = 'EpicorTest10'

<#3 Epicor QA
Foundation TEST -> Epicor QA#>
$epicorIntSvcQA = 'TEST-FLWA1.sufs.local'
$appPoolQA = 'EpicorQA10'

<#4 Epicor PROD
Foundation PROD -> Epicor PROD#>
$epicorIntSvcPROD = 'PROD-FLSV1.sufs.local'
$appPoolProd = 'Prod-Epicor_AppPool'

<#5 Epicor Test-PROD
Foundation UAT -> Epicor Test-PROD#>
$epicorIntSvcTestPROD = 'DC-AE-ADMIN1.sufs.local'
$appPoolTestPROD = 'ERP10Test'

<#Other Epicor PROD web servers (currently not a part of the refresh process)
$foundationWSFLWA1 = 'PROD-FLWA1.sufs.local'
$foundationWSFLWA2 = 'PROD-FLWA2.sufs.local'#>

#Define Epicor application and task agent servers:
#Lower Env application and task agent server
$epicorAppSvrAllLower = 'PROD-EPIAPP1.sufs.local'
$epicorTskSvrAllLower = 'PROD-EPIAPP1.sufs.local'
#Lower Env application and task agent server
$epicorTskSvrPROD = 'PROD-EPICORTSK1.sufs.local'
$epicorAppSvrPROD = 'PROD-EPICOR1.sufs.local'
#Test-Prod application and task agent server
$epicorTskSvrTestPROD = 'DC-AE-ADMIN1.sufs.local'
$epicorAppSvrTestPROD = 'DC-AE-ADMIN1.sufs.local'

#Current version of Task Agent Service.  This variable can be updated as Epicor is upgraded to newer versions.
$taskAgentService = 'Epicor ICE Task Agent 3.1.600.0'


#Create an array to use in ForEach loops that includes pertinent server information for all lower environments.  The values for each object can be altered as new lower environments are created, services moved, etc.
#Order of Array columns is "EnvName"  |  "IntServer" (Intergration Service Server)"  |  "TskAgent" (Task agent server)  |  "WebServer" (typically the Epicor Application Server)  |  "AppPool" (name of IIS application pool)"
#The first commented line under @() is just an example.  To set the values of this array change them in the similar line at the top of each item in the below SWITCH statement
$epicorEnvironments = @()
#endregion
### <(^_^)> --- END VARIABLES REGION --- <(^_^)>


### >(^_^)>  >(^_^)> --- FUNCTIONS --- <(^_^)`<  <(^_^)<
#region Functions
#Create function to pause script (Source: https://adamstech.wordpress.com/2011/05/12/how-to-properly-pause-a-powershell-script/)
Function Pause($M="Press any key to continue . . . "){If($psISE){$S=New-Object -ComObject "WScript.Shell";$B=$S.Popup("Click OK to continue to proceed to the next step...",0,"Script Paused",0);Return};Write-Host -NoNewline $M;$I=16,17,18,20,91,92,93,144,145,166,167,168,169,170,171,172,173,174,175,176,177,178,179,180,181,182,183;While($K.VirtualKeyCode -Eq $Null -Or $I -Contains $K.VirtualKeyCode){$K=$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")};Write-Host}

#Create function to define array containing all lower environment 
Function New-UserObject ($EnvName, $IntServer, $TskAgent, $WebServer, $AppPool) 
{
    New-Object PsObject -Property @{
        EnvName = $EnvName;
        IntServer = $IntServer;
        TskAgent = $TskAgent;
        WebServer = $WebServer;
        AppPool= $AppPool
   }
}

#Create function to generate text displayed for case statement
Function Show-Menu
{
     param 
     (
           [string]$Title = ' --- Epicor Refresh Automation  --- '
     )
        Clear-Host
        Write-Host -for DarkMagenta "Script Label Color Coding:"
        Write-Host -for DarkGray "--------------------------------------------------"
        Write-Host -for Yellow "Menus/Menu Header"
        Write-Host -for DarkYellow "Menu Items/Informational"
        Write-Host -for Cyan "Action starting..."
        Write-Host -for Green "Action completed Sucessfully"
        Write-Host -for Red "Action failed/Error/Important Message"
        Write-Host -for DarkGray "--------------------------------------------------"     
        Write-Host -for Yellow ">(^_^)> --- $Title --- <(^_^)<" 
        Write-Host -for Yellow "" 
        Write-Host -for Yellow "- Please Select the Target Environment to Begin *PRE* Database-Refresh Tasks -" 
        Write-Host -for DarkYellow "1: Epicor DEV."
        Write-Host -for DarkYellow "2: Epicor TEST."
        Write-Host -for DarkYellow "3: Epicor QA."
        Write-Host -for DarkYellow "4: Epicor PROD."
        Write-Host -for DarkYellow "5: Epicor Test-PROD."
        Write-host -for DarkYellow "A: for ALL Epicor Lower Epicor Environments"
        Write-host -for DarkYellow "R: for *POST* DB-Refresh Tasks for Current Target"
        Write-Host -for DarkYellow "Q: Press 'Q' to finish."
}

###STEP 0 - DISABLE INTEGRATION SERVICES USED BY THE ENVIRONMENT TO BE REFRESHED
#Display general information of current target environment
Function DisplayEnvInfo()
{
    ForEach ($envrnmnt in $epicorEnvironments)
    {
        Write-Host -for Yellow "***EPICOR Environment Name***"
        $envrnmnt.EnvName
        Write-Host -for DarkYellow "Integration Services Server:"
        $envrnmnt.IntServer
        Write-Host -for DarkYellow "Task Agent Server:"
        $envrnmnt.TskAgent
        Write-Host -for DarkYellow "Web Server and Application Pool:"
        $envrnmnt.WebServer
        $envrnmnt.AppPool
        Write-Host -for Gray "-------------------------------------------"
        }
}

###STEP 2 - DISABLE INTEGRATION SERVICES USED BY THE ENVIRONMENT TO BE REFRESHED
Function DisableIntegrationSvcs()
{
    ForEach ($envrnmnt in $epicorEnvironments)
    {    
    #Shutdown Integration Services for current environment in ForEach loop
    Write-Host -for Cyan "Stopping Epicor Integration Service(s) on $($envrnmnt.IntServer)...."


    #Get date for Do While statement to use as timeout value
    $startDate = Get-Date

    do 
    {
        #SET COMMAND...stop STOP Epicor Intetgration Services
        Get-Service -Name 'Epicor Integration Service' -ComputerName $envrnmnt.IntServer | Set-Service -Status Stopped
        Get-Service -Name 'Epicor PLSA Integration Service' -ComputerName $envrnmnt.IntServer | Set-Service -Status Stopped

        #10 second delay to wait for service to stop
        Start-Sleep -Seconds 10

        #Get the current status of each service to detemine if its in STOPPED state.  
        $intServiceStatus = Get-Service -Name 'Epicor Integration Service' -ComputerName $envrnmnt.IntServer | %{$_.Status}
        $intServiceStatusPLSA = Get-Service -Name 'Epicor PLSA Integration Service' -ComputerName $envrnmnt.IntServer | %{$_.Status}
    } 
    
    #$startDate.AddMinutes() value can be increased to any specified number of minutes to increase the timeout value
    while ($intServiceStatus -match "Running" -or $intServiceStatusPLSA -match "Running" -and $startDate.AddMinutes(3) -gt (Get-Date))


    #Do final check to see if services stopped sucessfully
    If ($intServiceStatus -match "Stopped" -and $intServiceStatusPLSA -match "Stopped")
    {
        #Report to user that both services have stopped sucessfully
        Write-Host -for Green "Successful!  Current status of Integration Services for $($envrnmnt.EnvName) on $($envrnmnt.IntServer):"
        Write-Host -for DarkRed "STOPPED (Epicor Integration Service)"
        Write-Host -for DarkRed "STOPPED (Epicor PLSA Integration Service)"
        Write-Host -for DarkGray "--------------------------------------------------"
    } 
    Else 
    {
        #Report to user that one or both services failed to stop and manual intervention is required.
        Write-Host -for Red "One or both services timed out while attempting to stop.  Manual remediation is required to stop Integration Services for $($envrnmnt.EnvName) on $($envrnmnt.IntServer)."
        Write-Host -for Red "IMPORTANT: Please ensure both services are stopped before continuing."
        Write-Host -for DarkGray "--------------------------------------------------"

    #PAUSE script and wait for user to remediate issues
    Pause
    }
}
}

###STEP 3 - STOP TASK AGENT SERVICE USED BY THE TARGET ENVIRONMENT
Function StopTaskAgent()
{
    ForEach ($envrnmnt in $epicorEnvironments)
    {    
        #Inform the user of the third pre refresh step
        Write-Host -for Cyan "STOPPING task agent service ($taskAgentService) on $($envrnmnt.TskAgent)...."

        #Get current status of service
        $taskAgentSvcStatus = Get-Service -Name "$($taskAgentService)" -ComputerName "$($envrnmnt.TskAgent)"
        
        #Do final check to see if task agent service has been STOPPED.
        If ($taskAgentSvcStatus.Status -match "Stopped")
            {
                #Report to user that Task Agent Service has been STOPPED
                Write-Host -for Green "Successful!  Task Agent Service ($taskAgentService) on $($envrnmnt.TskAgent) has STOPPED sucessfully.  Current status:"
                Write-Host -for DarkRed "STOPPED ($taskAgentService)"
                Write-Host -for DarkGray "--------------------------------------------------"
            } 
        ElseIf ($taskAgentSvcStatus.Status -match "Running")
            {
            #Get date for Do While statement to use as timeout value
            $startDate = Get-Date
            #Do While loop will retry to stop the service but will timeout after number of minutes set in $startDate.AddMinutes() variable below
            do 
                {
                    #SET COMMAND - STOP the Epicor Task Agent Service
                    Get-Service -Name "$($taskAgentService)" -ComputerName "$($envrnmnt.TskAgent)" | Set-Service -Status Stopped
                    #10 second delay to wait for service to stop
                    Start-Sleep -Seconds 10

                    #Get new status of service
                    $taskAgentSvcStatus = Get-Service -Name "$($taskAgentService)" -ComputerName "$($envrnmnt.TskAgent)"
                }
            #$startDate.AddMinutes() value can be increased to any specified number of minutes to increase the timeout value
            while ($taskAgentSvcStatus.Status -match "Running" -and $startDate.AddMinutes(2) -gt (Get-Date))  

            #Report to user that Task Agent Service has been STOPPED
            Write-Host -for Green "Successful!  Task Agent Service ($taskAgentService) on $($envrnmnt.TskAgent) has STOPPED sucessfully.  Current status:"
            Write-Host -for DarkRed "STOPPED ($taskAgentService)"
            Write-Host -for DarkGray "--------------------------------------------------"
            }
         Else
            {
                #Report to user that service may have failed to stop and manual intervention is required.
                Write-Host -for Red "!!!WARNING!!! Task Agent Service ($taskAgentService) on $($envrnmnt.TskAgent) FAILED to STOP.  Manual remediation is required.
                Do NOT click OK to proceed until this service has been STOPPED."
                Write-Host -for DarkGray "--------------------------------------------------"

                #PAUSE script and wait for user to remediate issues
                Pause
            }
    }
}

###STEP 5 - STOP IIS APP POOLS RUNNING FOR EACH TARGET ENVIRONMENT
Function StopIISAppPool()
{
    ForEach ($envrnmnt in $epicorEnvironments)
    {   
        #Stop IIS Application Pool for current environment in ForEach loop
        Write-Host -for Cyan "Stopping Web Application pool $($envrnmnt.AppPool) on $($envrnmnt.WebServer)...."

        #Get date for Do While statement to use as timeout value
        $startDate = Get-Date
                        
        #Do While loop will retry to STOP the current environment's AppPool but will timeout after number of minutes set in $startDate.AddMinutes() variable below
        do 
        {
            #Get name of current environment App Pool
            $appPoolName = "$($envrnmnt.AppPool)"

            #SET COMMAND - STOP App Pool
            Invoke-Command -ComputerName "$($envrnmnt.WebServer)" { param($apn) Stop-WebAppPool $apn} -Args $appPoolName

            #10 second delay to wait for service to stop
            Start-Sleep -Seconds 10

            #Get new state of App Pool
            $appPoolStatus = Invoke-Command -ComputerName "$($envrnmnt.WebServer)" { param($apn) Get-WebAppPoolState $apn} -Args $appPoolName
         }
        #$startDate.AddMinutes() value can be increased to any specified number of minutes to increase the timeout value
        while ($appPoolStatus.Value -match "Started" -or $appPoolStatus.Value -match "Stopping" -and $startDate.AddMinutes(2) -gt (Get-Date))

        #Do final check to see if AppPool for current environment has been stopped
        If ($appPoolStatus.Value -match "Stopped")
            {
                #Report to user that Application Pool is stopped
                Write-Host -for Green "Successful!  Web Application Pool for $($envrnmnt.EnvName) on $($envrnmnt.WebServer) has STOPPED sucessfully.  Current status:"
                Write-Host -for DarkRed "STOPPED ($($envrnmnt.AppPool))"
                Write-Host -for DarkGray "--------------------------------------------------"
            } 
        Else 
            {
                #Report to user that App Pool for current environment has failed attempting to STOP
                Write-Host -for Red "!!!WARNING!!! Application Pool timed out while attempting to STOP.  Manual remediation is required to stop Web Application Pool $($envrnmnt.AppPool) on $($envrnmnt.WebServer)."
                Write-Host -for DarkGray "--------------------------------------------------"
                
                #PAUSE script and wait for user to remediate issues
                Pause
            }
        }
}

###STEP 6 to 10 - DBA TEAM RESPONSIBLE FOR THESE ITEMS (SSIS PACKAGE)
Function DBATeamRefreshPackage()
{
    Write-Host -for Green "!!COMPLETED!! All Integration services and app pools for *ALL* lower environments have been stopped.  DBA team can now execute refresh package."
    Write-Host -for Cyan "After DBA confirms refresh for ALL environments is complete, click OK and select option R to complete steps 11-17"
}

###STEP 11 - REGENERATE DATA MODEL (MANUAL STEP)
Function RegenerateDataModel()
{
    #Inform the user of the first post refresh step
    Write-Host -for Yellow "!!IMPORTANT!! Verify with Data and Implementation team that the deployment/refresh has been completed before continuing."
    #PAUSE script and wait for user to verify the refresh has been completed
    Pause                
               
    Write-Host -for Cyan "Regenerating Database Data Models...."
    Write-Host -for Yellow "The next step is a manual process."
    Write-Host -for DarkYellow "1.	From target environment(s) Windows Server, open Epicor Administration console."
    Write-Host -for DarkYellow "2.	Expand Database Server Management > PROD-SQLEPI1, and select the target environment’s database. "
    Write-Host -for DarkYellow "3.	Right click and select Regenerate Data Model. "
    Write-Host -for DarkYellow "4.	Completion may take between 10-30 minutes and the completion dialog box may pop up under the other windows. "
    Write-Host -for DarkYellow "5.  This must be done for the database belonging to each environment that was refreshed."
    Write-Host -for Yellow "!!IMPORTANT!! Do not continue until the Regenerate Data Model process for EVERY target environment has been completed!"
    Write-Host -for DarkGray "--------------------------------------------------"
    Pause

    Write-Host -for Cyan "User confirmed database data models were regenerated...."
}

###STEP 12 - START IIS APP POOLS RUNNING FOR EACH TARGET ENVIRONMENT
Function StartIISAppPool()
{
    Write-Host -for Cyan "Restarting App Pools for ALL environments...."

    ForEach ($envrnmnt in $epicorEnvironments)
        {
            #Report to user what environment is currently being manipulated
            Write-Host -for Cyan "Starting Web Application pool $($envrnmnt.AppPool) on $($envrnmnt.WebServer)...."
                       
            #Get date for Do While statement to use as timeout value
            $startDate = Get-Date
            do 
                {
                    #Get name of current environment App Pool
                    $appPoolName = "$($envrnmnt.AppPool)"
                                
                    #SET COMMAND - START App Pool
                    Invoke-Command -ComputerName "$($envrnmnt.WebServer)" { param($apn) Start-WebAppPool $apn} -Args $appPoolName
                                
                    #10 second delay to wait for service to stop
                    Start-Sleep -Seconds 10

                    #Get new state of App Pool
                    $appPoolStatus = Invoke-Command -ComputerName "$($envrnmnt.WebServer)" { param($apn) Get-WebAppPoolState $apn} -Args $appPoolName                                
                }
                #$startDate.AddMinutes() value can be increased to any specified number of minutes to increase the timeout value
            while ($appPoolStatus.Value -match "Stopped" -or $appPoolStatus.Value -match "Starting" -and $startDate.AddMinutes(2) -gt (Get-Date))          
                        
                        
            #Do final check to see if AppPool for current environment has been started
            If ($appPoolStatus.Value -match "Started")
                {
                    #Report to user that Application Pool has started
                    Write-Host -for Green "Web Application Pool for $($envrnmnt.EnvName) on $($envrnmnt.WebServer):"
                    Write-Host -for DarkGreen "STARTED ($($envrnmnt.AppPool)"
                    Write-Host -for DarkGray "--------------------------------------------------"
                } 
            Else 
                {
                    #Report to user that one or both services failed to stop and manual intervention is required.
                    Write-Host -for Red "!!!WARNING!!! Application Pool timed out while attempting to start.  Manual remediation is required to START Web Application Pool for $($envrnmnt.AppPool) on $($envrnmnt.WebServer)."
                    Write-Host -for Red "Do NOT click OK to proceed until application pools have been started."
                    Write-Host -for DarkGray "--------------------------------------------------"
                                                                
                    #PAUSE script and wait for user to remediate issues
                    Pause
                }
        }
}

###STEP 13 - START TASK AGENT SERVICE                
Function StartTaskAgent()
{
    ForEach ($envrnmnt in $epicorEnvironments)
    {
        #Inform the user of the 13th step in the process
        Write-Host -for Cyan "STARTING task agent service for ALL lower Environments."                                                
        #Get current status of service
        $taskAgentSvcStatus = Get-Service -Name "$($taskAgentService)" -ComputerName "$($envrnmnt.TskAgent)"
                        
        #Do initial check to see if task agent service has been STARTED.
        If ($taskAgentSvcStatus.Status -match "Running")
            {
                #Report to user that Task Agent Service has been STARTED
                Write-Host -for Green "Successful!  Task Agent Service ($taskAgentService) on $($envrnmnt.TskAgent) has STARTED sucessfully.  Current status:"
                Write-Host -for DarkGreen "                      STARTED ($taskAgentService)"
                Write-Host -for DarkGray "--------------------------------------------------"
            } 
        ElseIf ($taskAgentSvcStatus.Status -match "Stopped")
            {
                #Get date for Do While statement to use as timeout value
                $startDate = Get-Date
                #Do While loop will retry to START the service but will timeout after number of minutes set in $startDate.AddMinutes() variable below
                do 
                    {
                        #SET COMMAND - START the Epicor Task Agent Service
                        Get-Service -Name "$($taskAgentService)" -ComputerName "$($envrnmnt.TskAgent)" | Set-Service -Status Running
                        #10 second delay to wait for service to stop
                        Start-Sleep -Seconds 10
                                
                        #Get new status of service
                        $taskAgentSvcStatus = Get-Service -Name "$($taskAgentService)" -ComputerName "$($envrnmnt.TskAgent)"
                    }
                #$startDate.AddMinutes() value can be increased to any specified number of minutes to increase the timeout value
                while ($taskAgentSvcStatus.Status -match "Stopped" -and $startDate.AddMinutes(2) -gt (Get-Date))
                #Report to user that Task Agent Service has been STARTED
                Write-Host -for Green "Successful!  Task Agent Service ($taskAgentService) on $($envrnmnt.TskAgent) has STARTED sucessfully.  Current status:"
                Write-Host -for DarkGreen "STARTED ($taskAgentService)"
                Write-Host -for DarkGray "--------------------------------------------------"
            } 
        Else
            {
                #Report to user that task agent service may have failed to stop and manual intervention is required.
                Write-Host -for Red "!!!WARNING!!! Task Agent Service ($taskAgentService) on $($envrnmnt.TskAgent) FAILED to START.  Manual remediation is required.
                Do NOT click OK to proceed until this service has STARTED."
                Write-Host -for DarkGray "--------------------------------------------------"
                                                                                                
                #PAUSE script and wait for user to remediate issues
                Pause
            }
    }
}

###STEP 14 - START INTEGRATION SERVICES USED BY THE ENVIRONMENT TO BE REFRESHED
Function StartIntegrationServices()
{
ForEach ($envrnmnt in $epicorEnvironments)
    {
        #Startup Integration Services for current environment in ForEach loop
        Write-Host -for Cyan "Starting Epicor Integration Service on $($envrnmnt.IntServer)...."
                                                                        
        #Get date for Do While statement to use as timeout value
        $startDate = Get-Date
        do 
            {
                #SET COMMAND - START integration services
                Get-Service -Name 'Epicor Integration Service' -ComputerName $envrnmnt.IntServer | Set-Service -Status Running
                Get-Service -Name 'Epicor PLSA Integration Service' -ComputerName $envrnmnt.IntServer | Set-Service -Status Running
                            
                #10 second delay to wait for service to stop
                Start-Sleep -Seconds 10
                            
                #Get the current status of each service to detemine if its in STARTED state.  Do While loop will retry to stop the service but will timeout after 10 minutes.  
                $intServiceStatus = Get-Service -Name 'Epicor Integration Service' -ComputerName $envrnmnt.IntServer | %{$_.Status}
                $intServiceStatusPLSA = Get-Service -Name 'Epicor PLSA Integration Service' -ComputerName $envrnmnt.IntServer | %{$_.Status}

            } 
        #$startDate.AddMinutes() value can be increased to any specified number of minutes to increase the timeout value
        while ($intServiceStatus -match "Stopped" -or $intServiceStatusPLSA -match "Stopped" -and $startDate.AddMinutes(3) -gt (Get-Date))
                           

        #Do final check to see if services STARTED sucessfully
        If ($intServiceStatus -match "Running" -and $intServiceStatusPLSA -match "Running")
            {
                #Report to user that both services have STARTED sucessfully
                Write-Host -for Green "Integration Services for $($envrnmnt.EnvName) on $($envrnmnt.IntServer):"
                Write-Host -for DarkGreen "STARTED (Epicor Integration Service)"
                Write-Host -for DarkGreen "STARTED (Epicor PLSA Integration Service)"
                Write-Host -for DarkGray "--------------------------------------------------"
            } 
        Else 
            {
                #Report to user that one or both services failed to START and manual intervention is required.
                Write-Host -for Red "One or both services timed out while attempting to START.  Manual remediation is required to START Integration Services for $($envrnmnt.EnvName) on $($envrnmnt.IntServer)."
                Write-Host -for Red "IMPORTANT: Please ensure both services are STARTED before continuing."
                Write-Host -for DarkGray "--------------------------------------------------"

                #PAUSE script and wait for user to remediate issues
                Pause
            }

    }
}

###STEP 15 - PURGE BPM Artifacts (MANUAL STEP)
Function PurgeBPMArtifacts()
{
    Write-Host -for Cyan "Purging BPM Artifacts...."
    <#New-Item "\\$($envrnmnt.WebServer)\D$\Backup\BPM\$($envrnmnt.EnvName)\$($date)" -ItemType directory#>

    Write-Host -for Yellow "The next step is a manual process."
    Write-Host -for DarkYellow "1.	From target environment(s) Windows Server, copy contents of D:\Websites\$($envrnment.EnvName)\Server\BPM."
    Write-Host -for DarkYellow "    to D:\Backup\BPM\[ENVIRONMENT]\$($date)"
    Write-Host -for DarkYellow "2.	Delete contents of D:\Websites\$($envrnment.EnvName)\Server\BPM folder."
    Write-Host -for DarkGray "--------------------------------------------------"
    Pause

    Write-Host -for Cyan "User confirmed BPM Artifacts were purged...."
}

###STEP 16 - Recompile BPMs (MANUAL STEP)
Function RecompileBPMs()
{
    Write-Host -for Cyan "Recompiling BPMs...."
    Write-Host -for Yellow "The next step is a manual process."
    Write-Host -for DarkYellow "1.	From target environment(s) Windows Server, open Epicor the Epicor client software."
    Write-Host -for DarkYellow "2.	Navigate to System Management > Business Process Management > Directive Update."
    Write-Host -for DarkYellow "3.	Click the Directive Recompile Setup tab."
    Write-Host -for DarkYellow "4.	Select the Both outdated and up to date directive check box."
    Write-Host -for DarkYellow "5.  Select the Refresh Signatures check box."
    Write-Host -for DarkYellow "5.  Click the Start Recompile button. (a dialog box displays indicating the BMP Directives are recompiled)."
    Write-Host -for DarkGray "--------------------------------------------------"
    Pause

    Write-Host -for Cyan "User confirmed BPM Artifacts were recompiled...."
}

###STEP 17 - Deploy Dashboard UIs
Function DeployDashboardUIs()
{
    Write-Host -for Cyan "Deploying Dashboard UIs...."
    Write-Host -for Yellow "The next step is a manual process."
    Write-Host -for DarkYellow "1.	Navigate to System Management > Upgrade/Mass Regeneration > Dashboard Maintenance."
    Write-Host -for DarkYellow "2.	Click the Dashboard ID button.  In the search form, click Options.  Select the Return All Rows check box to override the default of 100.  Click OK."
    Write-Host -for DarkYellow "3.	Click Search to display the list of deployed dashboards.  In the results list click Select All. Click OK."
    Write-Host -for DarkYellow "4.	From the actions menu, select Deploy All UI Applications.  The selected dashboards are redeployed."
    Write-Host -for DarkYellow "5.  Repeat the above steps for each company in your Epicor ERP application."
    Write-Host -for Yellow "!!IMPORTANT!! Do not continue until these steps have been completed for all companies in the environment."
    Write-Host -for DarkGray "--------------------------------------------------"
    Pause

    Write-Host -for Cyan "User confirmed database data models were regenerated...."
}

#endregion
### <(^_^)> --- END FUNCTIONS REGION --- <(^_^)>


### >(^_^)>  >(^_^)> --- CORE RUN CODE --- <(^_^)<  <(^_^)<
#region CoreRunCode
do
{
     #Call Show-Menu and record input from user to determine which switch option to use
     Show-Menu
     $input = Read-Host "Please make a selection (to start over at any time hit CTRL+C.)"
     switch ($input)
     {
           '1' {
                Clear-Host
                <#1 Epicor DEV
                Foundation DEV --> Epicor DEV#>

                #Clear values set in $epicorEnvironments array so that it can be used in another iteration of code for a different environment
                $epicorEnvironments = @();

                #Set target environment notification text
                $targetDescription = "Foundation DEV --> Epicor DEV"

                #Set values for target environment
                $epicorEnvironments += New-UserObject "EPI DEV" "$($epicorIntSvcDEV)" "$($epicorTskSvrAllLower)" "$($epicorAppSvrAllLower)" "$($appPoolDev)"

           } '2' {
                Clear-Host
                <#2 Epicor TEST
                Foundation UAT --> Epicor TEST#>

                #Clear values set in $epicorEnvironments array so that it can be used in another iteration of code for a different environment
                $epicorEnvironments = @();

                #Set target environment notification text
                $targetDescription = "Foundation UAT --> Epicor TEST"
                
                #Set values for target environment
                $epicorEnvironments += New-UserObject "EPI TEST" "$($epicorIntSvcTEST)" "$($epicorTskSvrAllLower)" "$($epicorAppSvrAllLower)" "$($appPoolTest)"

           } '3' {
                Clear-Host
                <#3 Epicor QA
                Foundation TEST --> Epicor QA#>

                #Clear values set in $epicorEnvironments array so that it can be used in another iteration of code for a different environment
                $epicorEnvironments = @();
                
                #Set target environment notification text
                $targetDescription = "Foundation TEST --> Epicor QA"

                #Set values for target environment
                $epicorEnvironments += New-UserObject "EPI QA" "$($epicorIntSvcQA)" "$($epicorTskSvrAllLower)" "$($epicorAppSvrAllLower)" "$($appPoolQA)"

           } '4' {
                Clear-Host
                <#4 Epicor PROD
                Foundation PROD --> Epicor PROD#>

                #Clear values set in $epicorEnvironments array so that it can be used in another iteration of code for a different environment
                $epicorEnvironments = @();

                #Set target environment notification text
                $targetDescription = "Foundation PROD --> Epicor PROD"

                #Set values for target environment
                $epicorEnvironments += New-UserObject "EPI PROD" "$($epicorIntSvcProd)" "$($epicorTskSvrProd)" "$($epicorAppSvrProd)" "$($appPoolProd)"

           } '5' {
                Clear-Host
                <#1 Epicor TEST-PROD
                Foundation UAT --> Epicor Test-Prod#>

                #Clear values set in $epicorEnvironments array so that it can be used in another iteration of code for a different environment
                $epicorEnvironments = @();

                #Set target environment notification text
                $targetDescription = "Foundation UAT --> Epicor Test-PROD"

                #Set values for target environment
                $epicorEnvironments += New-UserObject "EPI Test-Prod" "$($epicorIntSvcTestPROD)" "$($epicorTskSvrTestPROD)" "$($epicorAppSvrTestPROD)" "$($appPoolTestPROD)"

           } 'A' {
                Clear-Host
                <#1 Epicor ALL Lower Environments (DEV, TEST, QA)
                Foundation DEV --> Epicor DEV
                Foundation TEST --> Epicor TEST
                Foundation UAT --> Epicor QA#>

                #Clear values set in $epicorEnvironments array so that it can be used in another iteration of code for a different environment
                $epicorEnvironments = @();

                #Set target environment
                $targetDescription = "*ALL* Lower Environments: Foundation DEV --> Epicor DEV ... Foundation TEST --> Epicor TEST ... Foundation UAT --> Epicor QA"
                
                #Set values for target environment
                $epicorEnvironments += New-UserObject "EPI DEV" "$($epicorIntSvcDEV)" "$($epicorTskSvrAllLower)" "$($epicorAppSvrAllLower)" "$($appPoolDev)"
                $epicorEnvironments += New-UserObject "EPI TEST" "$($epicorIntSvcTEST)" "$($epicorTskSvrAllLower)" "$($epicorAppSvrAllLower)" "$($appPoolTest)"
                $epicorEnvironments += New-UserObject "EPI QA" "$($epicorIntSvcQA)" "$($epicorTskSvrAllLower)" "$($epicorAppSvrAllLower)" "$($appPoolQA)"
          } 'R' {
####BEGIN RUNNING *POST* DATABASE-REFRESH TASKS`
                #Check to see if post-refresh tasks have already been completed.
                Write-Host -for Magenta "Running *POST* Database-Refresh tasks for:"
                Write-Host -for Magenta $targetDescription

                ###STEP 12 - START IIS APP POOLS RUNNING FOR EACH TARGET ENVIRONMENT
                StartIISAppPool

                ###STEP 13 - START TASK AGENT SERVICE
                StartTaskAgent

                ###STEP 14 - START INTEGRATION SERVICES USED BY THE ENVIRONMENT TO BE REFRESHED
                StartIntegrationServices
                
                ###STEP 15 - PURGE BPM ARTIFACTS (MANUAL STEP)
                PurgeBPMArtifacts

                ###STEP 16 - PURGE BPM ARTIFACTS (MANUAL STEP)
                RecompileBPMs

                ###STEP 17 - DEPLOY DASHBOARD UIs (MANUAL STEP)
                DeployDashboardUIs

                #Set flag to indicate that the entire refresh process has been comoleted.
                $refreshComplete = "Y"
          } 'q' {
                return
           }
     }

####BEGIN RUNNING *PRE* DATABASE-REFRESH TASKS IF AN ENVIRONMENT REFRESH WAS COMPLETED
    
    If ($refreshComplete = "Y")
    {
        Write-Host -for Cyan "Refresh of Environment is complete."
    } 
    Else
    {
        ###STEP 0 - SET AND DISPLAY TARGET ENVIRONMENT INFORMATION                
        Write-Host -for Magenta $targetDescription
        #Inform the user what environment was selected and will be manipulated
        DisplayEnvInfo

        ###STEP 2 - DISABLE INTEGRATION SERVICES USED BY THE ENVIRONMENT TO BE REFRESHED
        DisableIntegrationSvcs

        ###STEP 3 - STOP TASK AGENT SERVICE USED BY THE TARGET ENVIRONMENT
        StopTaskAgent

        ###STEP 5 - STOP IIS APP POOLS RUNNING FOR EACH TARGET ENVIRONMENT
        StopIISAppPool

        ###STEP 6 to 10 - DBA TEAM RESPONSIBLE FOR THESE ITEMS (SSIS PACKAGE)
        DBATeamRefreshPackage

        ###STEP 11 - REGENERATE DATA MODEL (MANUAL STEP)
        RegenerateDataModel
    }
    pause
}
until ($input -eq 'q')
#endregion
### <(^_^)> --- END CORE CODE BLOCK --- <(^_^)<