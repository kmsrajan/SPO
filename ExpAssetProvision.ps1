 <#---------------------------------------------------------------------------#
# Global variables
#  
#---------------------------------------------------------------------------#> 

$global:SPOAdminCenter="https://statoilintegrationtest-admin.sharepoint.com"
$global:SPOHomeUrl="https://statoilintegrationtest.sharepoint.com/"
$global:logFolderpath="C:\LAB\ADG\PS\"
$global:AssetTemplate="C:\LAB\ExpPoC\AssetTemplate.xml"
$global:AssetConfigList="AssetConfigurations"
$global:EquinorContentType = "Equinor Document"
$global:RegXSiteTitle = "[^a-zA-Z0-9_-]"
$global:totalWaitingTimeInMinutes = 10
$global:sleepLengthSecondsWaitingForEIMSiteDesign = 2
$global:iterationsToWaitForEIMSiteDesign = ($global:totalWaitingTimeInMinutes/$global:sleepLengthSecondsWaitingForEIMSiteDesign) * 60

$global:FormatFolderPath = "C:\LAB\ExpPoC\JSON\"


# Enable loging  
$date= Get-Date -format MMddyyyyHHmmss  
start-transcript -path "C:\LAB\ADG\PS\logs\Log_$date.doc"  


#---------------------------------------------------------------------------#> 
#---------------------------------------------------------------------------#> 


 <#---------------------------------------------------------------------------#
# Create the team site if it doesn't exists
#  
#---------------------------------------------------------------------------#> 
Function CreateTeamSite{
    
    #Check if site already existing
    $newSite= Get-PnPTenantSite -Identity $global:siteUrl

    if($newSite){
        Write-Host "Site already exists"
        
    }
    else{
        $siteAlias = $global:SiteName -replace $RegXSiteTitle, ""
        Write-Host "Creating new teamsite" 
        Write-Host "   Title:"$global:SiteName
        Write-Host "   Alias:"$siteAlias
        Write-Host "   Classification:"$global:SiteSC
        
        try{
        if($global:IsPublic){
            $newSiteUrl = New-PnPSite -Type TeamSite -Title $global:SiteName -Alias $siteAlias  -IsPublic -ErrorAction Stop
            $global:siteUrl = $newSiteUrl
            }
            else{
            $newSiteUrl = New-PnPSite -Type TeamSite -Title $global:SiteName -Alias $siteAlias  -ErrorAction Stop
            $global:siteUrl = $newSiteUrl
            }

        }
        catch{ 
            Write-Host $_           
            Write-Host "Error creating the team site for $siteInformation.siteTitle. Error:" $_ -ForegroundColor Red 
            Exit           
        }

       
    }
    
}

<#---------------------------------------------------------------------------#
# To check EIM metadta which is applied by site design
#  
#---------------------------------------------------------------------------#> 
function checkEIMSiteDesign(){        
    $counter = 0
    Write-Host "   Check if $global:EquinorContentType content type exists on the teamsite"
    while (-not (Get-PnPContentType -Identity $global:EquinorContentType -ErrorAction SilentlyContinue)) {            
        Write-Progress -Activity 'Checking for content type' -PercentComplete ((($counter++) / $global:iterationsToWaitForEIMSiteDesign) * 100)
        if($counter -gt $global:iterationsToWaitForEIMSiteDesign){
            Write-Host "$global:EquinorContentType content type not found after $global:totalWaitingTimeInMinutes minute(s). Check if there are error on the new team site:" $newSiteUrl -ForegroundColor Red
            Exit
        }
        Start-Sleep -Seconds $global:sleepLengthSecondsWaitingForEIMSiteDesign        
    }

    Write-Progress -Activity 'Checking for content type' -Status "Ready" -Completed    
    Write-Host "      $EquinorContentType content type exists on the teamsite" -ForegroundColor Green

    $counter = 0
    Write-Host "   Check if Enterprise Metadata view exists on Documents library"
    while (-not (Get-PnPView -List "Documents" -Identity "Enterprise Metadata" -ErrorAction SilentlyContinue)) {            
        Write-Progress -Activity 'Checking for Enterprise Metadata view' -PercentComplete ((($counter++) / $global:iterationsToWaitForEIMSiteDesign) * 100)
        if($counter -gt $global:iterationsToWaitForEIMSiteDesign){
            Write-Host "Enterprise Metadata view exists on Documents library not found after $global:totalWaitingTimeInMinutes minute(s). Check if there are error on the new team site:" $newSiteUrl -ForegroundColor Red
            Exit
        }
        Start-Sleep -Seconds $global:sleepLengthSecondsWaitingForEIMSiteDesign        
    }

    Write-Progress -Activity 'Checking for Enterprise Metadata view' -Status "Ready" -Completed
    
    Write-Host "      Giving the site a minute to set things right" -ForegroundColor Green
    Start-Sleep -Seconds 60
    Write-Host "      Enterprise Metadata view exists on Documents library" -ForegroundColor Green
}

 <#---------------------------------------------------------------------------#
# To connect to SharePoint Online
#  
#---------------------------------------------------------------------------#> 
Function ConnectSPO{
    Param(
    # SharePoint Site URL
    [Parameter(Mandatory = $true)]
    [String]$SPOUrl
    )
    try{
        #Connect-PnPOnline -Url $SPOUrl -Credentials $credential # To  use inside Equinor network
        Connect-PnPOnline -Url $SPOUrl -UseWebLogin
        #Connect-PnPOnline -Url $SPOUrl -ClearTokenCache -SPOManagementShell # To use outside Equinor network
                
        if (-not (Get-PnPContext)) {
            Write-Host "Error connecting to site $SPOUrl Unable to establish context" $_ -ForegroundColor Red            
            WriteToLog -Type "Failed" "Error connecting to site $SPOUrl Unable to establish context"
            Exit
        }
    } catch {
        Write-Host "Error connecting to site [$SPOUrl] Error: " $_ -ForegroundColor Red
        WriteToLog -Type "Failed" "Error connecting to site [$SPOUrl] Error:"
        Exit
    }
    Write-Host "   Connected to site" $SPOUrl -ForegroundColor green
}

 <#---------------------------------------------------------------------------#
# Configure site assets
#  
#---------------------------------------------------------------------------#> 
Function ConfigureAsset{
    try{        
        #Get asset columns
        [XML]$assetConfig = Get-Content $global:AssetTemplate
                      
        #Create asset common columns
        CreateSiteColumns -assetFields $assetConfig.pnpProvisioningTemplate.pnpSiteFields

        #add common content type
        $assetCType=$assetConfig | Select-XML –Xpath "//*[@AssetType='Subsurface']"
        
        CreateContentType -ctId $assetCType.Node.ID -ctName $assetCType.Node.Name -ctGroup $assetCType.Node.Group
        
        #Add columns to the content type
        AddColumnToContentType -ctName $assetCType.Node.Name -assetFields $assetConfig.pnpProvisioningTemplate.pnpSiteFields
       
        #Get asset specific content type
        $cType=$assetConfig | Select-XML –Xpath "//*[@AssetType='SubsurfaceAsset']"
        CreateContentType -ctId $cType.Node.ID -ctName $cType.Node.Name -ctGroup $cType.Node.Group


        
        if($global:SiteAssetType.Equals("License")){
            CreateSiteColumns -assetFields $assetConfig.pnpProvisioningTemplate.pnpLicenseFields
            AddColumnToContentType -ctName $cType.Node.Name -assetFields $assetConfig.pnpProvisioningTemplate.pnpLicenseFields
        }
        if($global:SiteAssetType.Equals("Well")){
            CreateSiteColumns -assetFields $assetConfig.pnpProvisioningTemplate.pnpWellFields
           AddColumnToContentType -ctName $cType.Node.Name -assetFields $assetConfig.pnpProvisioningTemplate.pnpWellFields
        }
        if($global:SiteAssetType.Equals("Prospect")){
            CreateSiteColumns -assetFields $assetConfig.pnpProvisioningTemplate.pnpProspectFields
             AddColumnToContentType -ctName $cType.Node.Name -assetFields $assetConfig.pnpProvisioningTemplate.pnpProspectFields
        }
        if($global:SiteAssetType.Equals("Regional")){
            CreateSiteColumns -assetFields $assetConfig.pnpProvisioningTemplate.pnpRegionalFields
             AddColumnToContentType -ctName $cType.Node.Name -assetFields $assetConfig.pnpProvisioningTemplate.pnpRegionalFields
        }
        if($global:SiteAssetType.Equals("Dataroom")){
            CreateSiteColumns -assetFields $assetConfig.pnpProvisioningTemplate.pnpDataroomFields
             AddColumnToContentType -ctName $cType.Node.Name -assetFields $assetConfig.pnpProvisioningTemplate.pnpDataroomFields
        }
        if($global:SiteAssetType.Equals("Field")){
            CreateSiteColumns -assetFields $assetConfig.pnpProvisioningTemplate.pnpFieldFields
            AddColumnToContentType -ctName $cType.Node.Name -assetFields $assetConfig.pnpProvisioningTemplate.pnpFieldFields
        }
        if($global:SiteAssetType.Equals("Survey")){
            CreateSiteColumns -assetFields $assetConfig.pnpProvisioningTemplate.pnpSurveyFields
            AddColumnToContentType -ctName $cType.Node.Name -assetFields $assetConfig.pnpProvisioningTemplate.pnpSurveyFields
        }
        if($global:SiteAssetType.Equals("Project")){
            CreateSiteColumns -assetFields $assetConfig.pnpProvisioningTemplate.pnpProjectFields
            AddColumnToContentType -ctName $cType.Node.Name -assetFields $assetConfig.pnpProvisioningTemplate.pnpProjectFields
        }

        Write-Host "================================================================"
        Write-Host "Subsurface columns and content type are ready"
        Write-Host "================================================================"
        

         ConnectSPO -SPOUrl $global:SiteUrl
        $assetConfigList="AssetConfigurations"
        
        $lstExist=Get-PnPList -Identity $assetConfigList

        if(!$lstExist){
            #Create a list
            $configList=New-PnPList -Title  $assetConfigList -Template GenericList 
        }
        
        $ctInConfigList=Get-PnPContentType -Identity "Subsurface Asset Metadata" -List $assetConfigList
        If(!$ctInConfigList){
             #Add content type and make is default
            Add-PnPContentTypeToList -List $assetConfigList -ContentType "Subsurface Asset Metadata" -DefaultContentType
            Write-Host "Subsurface Asset Metadata content type is added as default content type "
        }  
        else
        {
            Set-PnPDefaultContentTypeToList -List $assetConfigList -ContentType "Subsurface Asset Metadata" 
            Write-Host "Subsurface Asset Metadata content type is updated as default content type "
        }
       
        #remove Item content type from list
        Remove-PnPContentTypeFromList -List $assetConfigList -ContentType "Item"
        
        Write-Host "Item content type is removed from list"

        Set-PnPList -Identity $assetConfigList -EnableContentTypes $false

        Write-Host "Disabled custom content type in the list"
       
    } catch {
      Write-Host $_
    }
}

 <#---------------------------------------------------------------------------#
# Create fields
#  
#---------------------------------------------------------------------------#> 
Function CreateSiteColumns{
param($assetFields)

#Create asset common columns
        foreach($assetFld in $assetFields.Field){            
            $field = Get-PnPField -Identity $assetFld.Name -ErrorAction SilentlyContinue
           
            if(!$field){
                if($assetFld.outerXML){
                    try{
                        
                        if($assetFld.Type -clike "*TaxonomyFieldType*"){
                           if($assetFld.Mult -eq "FALSE"){
                                Add-PnPTaxonomyField -Id $assetFld.ID -DisplayName $assetFld.DisplayName -InternalName $assetFld.Name -TermSetPath $assetFld.TermSetPath -Group $assetFld.Group | Out-Null
                            }
                            else{
                                Add-PnPTaxonomyField -Id $assetFld.ID -DisplayName $assetFld.DisplayName -InternalName $assetFld.Name -TermSetPath $assetFld.TermSetPath -Group $assetFld.Group -MultiValue | Out-Null
                            }
                            $txFld=Get-PnPField -Identity $assetFld.Name
                            $txFld.TermSetId=$assetFld.TermSetId
                            

                        }
                        else{
                         $field = Add-PnPFieldFromXml -FieldXml $assetFld.outerXML -ErrorAction Stop 
                        }             
                    }
                    catch{
                        Write-Host "Error creating new site column " + $assetFld.Name + ". Please check columnXml before running the script again." -ForegroundColor Red
                        Write-Host $_.Exception.GetType().FullName 
                        Write-Host $_.Exception.Message 
                        Write-Host $_.Exception.StackTrace                      
                        Exit
                    }                
                    Write-Host "  " $assetFld.Name "- Site column created" -ForegroundColor Green
                }
                else{
                    Write-Host "  "$assetFld.Name "- Site column not found or baseSettings.config is missing the columnXml for this column" -ForegroundColor Red
                }
            } 
        }

}

 <#---------------------------------------------------------------------------#
# Create Create content types
#  
#---------------------------------------------------------------------------#> 

Function CreateContentType{
param($ctId,$ctName,$ctGroup)

try{
        $ctExists=Get-PnPContentType -Identity $ctId
        if($ctExists){
            Write-Host "Content type already exists"
        }
        else{
            Add-PnPContentType -Name $ctName -ContentTypeId $ctId -Group $ctGroup | Out-Null
            Write-Host "New Content type created "    $ctName    
         }
}
catch{}

}
 
 
 <#---------------------------------------------------------------------------#
# Create fields
#  
#---------------------------------------------------------------------------#> 
Function AddColumnToContentType{
param($assetFields,$ctName)


#Create asset common columns
        foreach($assetFld in $assetFields.Field){
            
            $ctype = get-pnpcontenttype -Identity $ctName
            
            $flds=Get-PnPProperty -clientobject $ctype -property "Fields"
            $fldExist=$flds | Where-Object {$_.InternalName -eq $assetFld.Name}

            If(!$fldExist){
                $field = Get-PnPField -Identity $assetFld.ID -ErrorAction SilentlyContinue
                if($field){
                    if($assetFld.outerXML){
                        try{
                            #$field = Add-PnPFieldFromXml -FieldXml $assetFld.outerXML -ErrorAction Stop                 
                            Add-PnPFieldToContentType -Field $assetFld.Name -ContentType $ctName
                        }
                        catch{
                            Write-Host "Error creating new site column " + $assetFld.Name + ". Please check columnXml before running the script again." -ForegroundColor Red
                            Write-Host $_.Exception.GetType().FullName
                             Write-Host $_.Exception.Message
                            Exit
                        }                
                        Write-Host "  " $assetFld.Name "- Site column added to " $ctName -ForegroundColor Green
                    }
                    else{
                        Write-Host "  "  $assetFld.Name "- Site column not found or baseSettings.config is missing the columnXml for this column" -ForegroundColor Red
                    }
                }
            }
            else{
                Write-Host "Field " $assetFld.Name "already exists in content type " $ctName
            } 
        }

}

<#---------------------------------------------------------------------------#
# To create log file. Reusable function
#  
#---------------------------------------------------------------------------#>
Function Create-Log
{
    Param(
        # Log folder Root
        [Parameter(Mandatory = $true)]
        [String]$LogFolderRoot,
        # The function Log file for
        [Parameter(Mandatory = $true)]
        [String]$LogFunction
    )
    $logFolderPath = "$LogFolderRoot\logfiles"
    $folderExist = Test-Path "$logFolderPath"
    if (!$folderExist)
    {
        $folder = New-Item "$logFolderPath" -type directory
    }
    $date = Get-Date -Format 'MMddyyyy_HHmmss'
    $logfilePath = "$logFolderPath\Log_{0}_{1}.txt" -f $LogFunction, $date
    Write-Verbose "Log file is writen to: $logfilePath"
    $logfile = New-Item $logfilePath  -type file
    return $logfilePath
}

 
<#---------------------------------------------------------------------------#
# To log each actions. Reusable function
#  
#---------------------------------------------------------------------------#> 
Function Log
{
    Param(
        [Parameter(Mandatory = $true)]
        [String]$Message,
        [String]$Type = "Message"
    )
    $date = Get-Date -Format 'HH:mm:ss'
    $logInfo = $date + " - [$Type] " + $Message
    $logInfo | Out-File -FilePath $logfilePath -Append
    if ($Type -eq "Succeed") { Write-Host $logInfo -ForegroundColor Green }
    elseif ($Type -eq "Error") { Write-Host $logInfo -ForegroundColor Red }
    elseif ($Type -eq "Warning") { Write-Host $logInfo -ForegroundColor Yellow }
    elseif ($Type -eq "Start") { Write-Host $logInfo -ForegroundColor Cyan }
    else { Write-Verbose $logInfo }
}
Write-host "Message" -InformationAction SilentlyContinue

Function Main
{
    $logfilePath = Create-Log -LogFolderRoot $global:logFolderpath "M365"
}

<#---------------------------------------------------------------------------#
# Initiate the site provioning
#  
#---------------------------------------------------------------------------#> 

Function NewAsset{
    Param(
        [Parameter(Mandatory = $true)]
        [String]$SiteName,
         [Parameter(Mandatory = $true)]
        [String]$NameInURL,
        [Parameter(Mandatory = $true)]
        [String]$SecurityClassification,
         [Parameter(Mandatory = $true)]
        [boolean]$IsPublic,
        [Parameter(Mandatory = $true)]
        [String]$AssetType
        )

        $logfilePath = Create-Log -LogFolderRoot $global:logFolderpath "M365"

        #check Asset template existing
        try{
            $xmlIn=Test-Path $global:AssetTemplate
            if($xmlIn){
                
                #Connect to SharePoint Online Admin Center
                ConnectSPO -SPOUrl $global:SPOAdminCenter

                #set global variables for site properties
                $global:SiteName = $SiteName
                $global:SiteUrl = ($SPOHomeUrl+"Sites/"+$SiteName)
                $global:SiteSC = $SecurityClassification
                $global:SitePrivacy = $Privacy
                $global:SiteAssetType = $AssetType

                CreateTeamSite 

                ConnectSPO -SPOUrl $global:SiteUrl
               
                #To ensure whether EIM content type enabled in newly created site
                #checkEIMSiteDesign                
                              
                #Configure asset list
                ConfigureAsset

                #Disconnect-PnPOnline


                ConnectSPO -SPOUrl $global:SiteUrl
                FormatAssetView -AssetType $AssetType

                #Update Asset configurations list permission
                RestrictConfigListPermission -ListName $global:AssetConfigList

                Write-Host "Asset configuration completed"
                
            }
            else
            {
                 WriteToLog -Type "Failed" "Asset template is not accesible"
            }

        }
        catch{
            Write-Host $_
            WriteToLog -Type "Failed" "Asset template not exists "
        }


}


<#---------------------------------------------------------------------------#
# Update asset configurations list permissions
#  
#---------------------------------------------------------------------------#> 
Function RestrictConfigListPermission{
Param(
    [Parameter(Mandatory = $true)]
    [String]$ListName
    )
    try{
        Get-PnPListPermissions -Identity $global:AssetConfigList -PrincipalId (Get-PnPGroup | Where-Object {$_.LoginName -clike '*Owner*' }).Id
        $ownerGrp=Get-PnPGroup | Where-Object {$_.LoginName -clike '*Owner*' }
        
        $acList =Get-PnPList -Identity $global:AssetConfigList -Includes RoleAssignments
        
        $acList.BreakRoleInheritance($true,$false)

        # get all the users and groups who has access
        $roleAssignments = $acList.RoleAssignments
        
        foreach ($roleAssignment in $roleAssignments)
        {
           
            Get-PnPProperty -ClientObject $roleAssignment -Property RoleDefinitionBindings, Member
            
            
            if($roleAssignment.RoleDefinitionBindings[0].RoleTypeKind.ToString() -ieq "Administrator"){
               

                Set-PnPListPermission -Identity $global:AssetConfigList -Group $roleAssignment.Member.LoginName -RemoveRole 'Full Control'
                Set-PnPListPermission -Identity $global:AssetConfigList -Group $roleAssignment.Member.LoginName -AddRole 'Edit'

                #$roleAssignment.RoleDefinitionBindings

            }
            else {
                if($ownerGrp.Id -ne $roleAssignment.PrincipalId){
                    Write-Host "Read only"
                    Set-PnPListPermission -Identity $global:AssetConfigList -Group $roleAssignment.Member.LoginName -RemoveRole 'Edit'
                    Set-PnPListPermission -Identity $global:AssetConfigList -Group $roleAssignment.Member.LoginName -AddRole 'Read'

                    #$roleAssignment.RoleDefinitionBindings

                }
            }
        }
    }
    catch{
        Write-Error -Exception $_.Exception.Message -Category 'Exception'
        Write-Error -Exception $_.Exception.StackTrace -Category 'Exception'
    }
}

<#---------------------------------------------------------------------------#
# Update asset configurations list permissions
#  
#---------------------------------------------------------------------------#> 
Function FormatAssetView{
param($AssetType)
    try{
       
        #Get all asset formats
        $styleFileIn=$global:FormatFolderPath +$AssetType+".json"

        $JSONformat = Get-Content  $styleFileIn 
        $JSONformat | ConvertFrom-Json |Out-Null                
        
        $view = Get-PnPView -List $global:AssetConfigList -Identity "All Items" -Includes "CustomFormatter"
        $view.CustomFormatter = $JSONformat
        
        $view.Update()
        Invoke-PnPQuery

    }
    catch{
        Log -Message $($_.Exception.Message) -Type 'Error'
        Log -Message $($_.Exception.StackTrace) -Type 'Error'
    }
}

$aType="Regional"
$sName="PocRegionalA"
NewAsset -SiteName $sName -NameInURL $sName -SecurityClassification "Internal" -IsPublic $true -AssetType $aType

Disconnect-PnPOnline
Stop-Transcript  
