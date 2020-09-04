# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

<#
.Link 
https://go.microsoft.com/fwlink/?linkid=2107264
#>
param(
    # the operation to perform
    [string] $operation = $null,
    [string] $smtpAddress = $null
)

$ver = 1.4

write-host ""
write-host "OCM-Script (ver:$ver)"
write-host ""

function showValidOptionsForExport()
{
    Write-Host "       export-all-data : Export all data for contacts, companies and deals."
    Write-Host "       export-contact-templates : Export custom fields for contacts."
    Write-Host "       export-company-templates : Export custom fields for companies."
    Write-Host "       export-deal-templates : Export custom fields for deals template."
    Write-Host "       export-contacts-data : export data for contacts."
    Write-Host "       export-companies-data : export data for companies."
    Write-Host "       export-deals-data : export data for deals."
    Write-Host "       export-tasks-data : export data for tasks."
    Write-Host "       export-posts-data : export data for posts."
    Write-Host "       export-activities-data : export data for activities."
}

function showValidOptionsForPurge()
{
    Write-Host "       purge-all-data <SMTP address of user or group mailbox> : Purge all data for contacts, companies, deals, xrmactivitystream, xrminsights, xrmdeleteditems and templates."
    Write-Host "       purge-contacts-data <SMTP address of user or group mailbox> : Purge all data for contacts."
    Write-Host "       purge-companies-data <SMTP address of user or group mailbox> : Purge all data for companies."
    Write-Host "       purge-ocm-personmeta-data <SMTP address of user or group mailbox> : Purge all data for personmetadata."
    Write-Host "       purge-xrm-activity-stream-data <SMTP address of user or group mailbox> : Purge all data for xrmactivitystream."
    Write-Host "       purge-xrm-insights-data <SMTP address of user or group mailbox> : Purge all data for xrminsights."
    Write-Host "       purge-xrm-deleted-items-data <SMTP address of user or group mailbox> : Purge all data for xrmdeleteditems."
    Write-Host "       purge-templates-data <SMTP address of user or group mailbox> : Purge all data for templates."
    Write-Host "       purge-deals-data <SMTP address of user or group mailbox> : Purge all data for deals."
}

if (!$operation)
{
    Write-Host -ForegroundColor Yellow "Please specify the operation to perform"
    Write-Host "Valid choices:"
    Write-Host
    showValidOptionsForExport
    showValidOptionsForPurge
    Write-Host
    break
}

function getUserCreds($user)
{
    # only ask for credentials if we don't have any

    if ($null -eq $cred -or ($user -and $cred.UserName -ine $user))
    {
        # if this script needs to be run many times: copy, paste and run the following line before running
        write-host "Enter credentials for OCM account:"
        write-host 'To avoid this prompt on repeat runs, use the following prior to this script: '
        write-host '  $cred = get-credential'

        $cred = get-credential -Message "Enter credentials for OCM account:" -UserName $user
    }
    $user = $cred.UserName

    if (!$user -or !$cred)
    {
        # exit early if we have no user information
        write-host
        write-host -ForegroundColor Red "Cannot continue without user info or credentials."
        write-host
        break;
    }

    write-host "(connecting to Exchange as '$user')"
    return $cred
}

function CallStartXrmSessionV2()
{
    if (!($smtpAddress -like "*c036c9b0-d80c-423b-afed-8642ce2d6076*"))
    {
        # skip when not the group
        return
    }

    try
    {
        $starXrmSession ="<?xml version='1.0' encoding='utf-8'?>
        <soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:m='http://schemas.microsoft.com/exchange/services/2006/messages' xmlns:t='http://schemas.microsoft.com/exchange/services/2006/types' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>
            <soap:Header>
            <t:RequestServerVersion Version=''V2017_07_11' />
            </soap:Header>
            <soap:Body>
            <m:StartXrmSession>
                <m:FailOnError>false</m:FailOnError>
                <m:ForcePersonaBackfill>false</m:ForcePersonaBackfill>
                <m:FlightAndLicenseCheckOnly>false</m:FlightAndLicenseCheckOnly>
                <m:ProvisionGroupIdentity>default</m:ProvisionGroupIdentity>
            </m:StartXrmSession>
            </soap:Body>
        </soap:Envelope>"

        Write-Host  "Calling StartXrmSession on usermailbox: $smtpAddress ..."

        $response = Invoke-WebRequest $uri -Method Post -ContentType "text/xml" -Credential $cred -Body "$starXrmSession" -Headers @{ 'X-AnchorMailbox' = "$smtpAddress" }
        $responseXML = [xml]$response
        $responseCode = (select-xml -Xml $responseXML -Namespace $ns -XPath "//m:ResponseCode").Node.InnerText

        if ($responseCode -eq 'NoError')
        {
            Write-Host "StartXrmSession success."
        }
        else
        {
            $errorMessageText = (select-xml -Xml $responseXML -Namespace $ns -XPath "//m:MessageText").Node.InnerText
            Write-Host -ForegroundColor Red "StartXrmSession failed with error - $errorMessageText"
        }
    }
    catch
    {
        Write-Host -ForegroundColor Red "StartXrmSession failed with error."
        $_.Exception.Message
    }
}


function GetFolderName()
{
    $dataTime = (Get-Date -Format FileDateTime)
    $folderName = $cred.UserName + "_" + $dataTime
    New-Item -Path ".\" -Name $folderName -ItemType "directory"
}

function ExportTemplatesDataToCSV
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        $customFieldList,
        [Parameter(Mandatory=$true)]
        [string] $templateType
        )

    $filepath = $hashFilepath[$templateType];

    $customFieldArray = @()

    foreach ($customField in $customFieldList) {

        if ($customField.Label -eq 'Status' -or $customField.Label -eq 'Stage' ) {
            continue
        }
            
        $obj = New-Object -TypeName psobject
        $obj | Add-Member -MemberType NoteProperty -Name FieldId -Value $customField.Id
        $obj | Add-Member -MemberType NoteProperty -Name FieldLabel -Value $customField.Label
        $obj | Add-Member -MemberType NoteProperty -Name FieldType -Value $customField._Type.Label

        if ($customField._Type.Label -eq 'Choice' -or $customField._Type.Label -eq 'Currency')
        {
            $valueOptions = $customField.ValueOptions

            if($valueOptions) {
                $valueOptionsArray = $valueOptions | ForEach-Object -Process {$_.Label } 
                $valueOptionsString = $valueOptionsArray -join '|'
                $obj | Add-Member -MemberType NoteProperty -Name FieldInfo -Value $valueOptionsString
            }
        }
      
        $IsPropertyAdded = $obj.PSobject.Properties.Name -contains "FieldInfo"

        if($IsPropertyAdded -eq $False) {
            $obj | Add-Member -MemberType NoteProperty -Name FieldInfo -Value ''
        }

        $customFieldArray += $obj
    }
      
    
    if ($customFieldArray.Count -gt 0 ) 
    {
        $filepath = ".\" + $folderName.BaseName + "\" + $filepath
        $customFieldArray | Export-Csv  -Path  $filepath -NoTypeInformation
    }
    
}

function CallExportContactTemplatesData
{
    #
    # Make a REST call for contact templates
    #
    try {

        $contactTemplate = Invoke-WebRequest -Method Get -uri "https://outlook.office365.com/api/beta/me/XrmContactTemplate" -Headers @{ 'Authorization' = "Bearer $token" } 

        if($contactTemplate) {
            write-host -ForegroundColor Green "Exporting data for contacts template..."
            $contactContent = $contactTemplate.Content | ConvertFrom-Json
            $fieldListContact = $contactContent.value[0].Template.FieldList

            if($fieldListContact) {
                ExportTemplatesDataToCSV $fieldListContact 'ContactTemplate'
            }
            else{
                write-host -ForegroundColor Yellow "No custom fields available for contacts template"
            }
        }
        else {
            write-host -ForegroundColor Yellow "No data available for the contacts template"
        }

        write-host -ForegroundColor Green "Exporting data for contacts template complete"

    } catch {
        write-host -ForegroundColor Green "Exporting data for contacts template failed"
        $_.Exception.Message
    }
}

function CallExportCompanyTemplatesData
{
    #
    # Make a REST call for company templates
    #
    try {
        
        $companyTemplate = Invoke-WebRequest -Method Get -uri "https://outlook.office365.com/api/beta/Me/XrmOrganizationTemplate" -Headers @{ 'Authorization' = "Bearer $token" } 

        if($companyTemplate) {
            write-host -ForegroundColor Green "Exporting data for company template..."
            $companyContent = $companyTemplate.Content | ConvertFrom-Json
            $fieldListCompany = $companyContent.value[0].Template.FieldList

            if($fieldListCompany) {
                ExportTemplatesDataToCSV $fieldListCompany 'CompanyTemplate'
            }
            else{
                write-host -ForegroundColor Yellow "No custom fields available for company template"
            }
        }
        else {
            write-host -ForegroundColor Yellow "No data available for the company template"
        }

        write-host -ForegroundColor Green "Exporting data for company template complete"

    } catch {
        write-host -ForegroundColor Green "Exporting data for company template failed"
        $_.Exception.Message
    }
}

function CallExportDealTemplatesData
{
    #
    # Make a REST call for deal templates
    #

    try {

        $dealTemplate = Invoke-WebRequest -Method Get -uri "https://outlook.office365.com/api/beta/me/XrmDealTemplate" -Headers @{ 'Authorization' = "Bearer $token" } 

        if($dealTemplate) {
            write-host -ForegroundColor Green "Exporting data for deal template..."
            $dealsContent = $dealTemplate.Content | ConvertFrom-Json
            $fieldListDeals = $dealsContent.value[0].Template.FieldList
            
            $stageListDeals = $dealsContent.value[0].Template.StatusList

            if($stageListDeals) {
                foreach($stageItem in $stageListDeals)
                {
                    foreach($stage in $stageItem.Stages)
                    {
                        $hashStageIdToLabel[$stage.Id] = $stage.Label
                    }
                }
            }

            if($fieldListDeals) {
                ExportTemplatesDataToCSV $fieldListDeals 'DealTemplate' 
            }
            else{
                write-host -ForegroundColor Yellow "No custom fields available for deals template"
            }
         }
        else {
            write-host -ForegroundColor Yellow "No data available for the deals template"
        }

        write-host -ForegroundColor Green "Exporting data for deals template complete"

    } catch {
        write-host -ForegroundColor Green "Exporting data for deals template failed"
        $_.Exception.Message
    }
}

function CallExportAllData {
    CallExportContactTemplatesData
    CallExportCompanyTemplatesData
    CallExportDealTemplatesData
    CallExportContactsData
    CallExportCompaniesData
    CallExportDealsData
    CallExportTasksData
    CallExportPostsData
    CallExportActivitiesData
}

function AddPropertyObject($obj, $hashProperty, $propertyName)
{
    $values = $null

    if($hashProperty.Count -gt 0) {
        $values = $hashProperty | Select-Object -ExpandProperty Keys 
        $values = $values -join '|'
        $values = $values.Trim('|')
    }

    if($values -ne $null) {
        $IsPropertyAdded = $obj.PSobject.Properties.Name -contains $propertyName
            
        if($IsPropertyAdded -eq $False) {
            $obj | Add-Member -MemberType NoteProperty -Name $propertyName -Value $values
        }
        else
        {
            $obj | Add-Member -MemberType NoteProperty -Name $propertyName -Value $values -Force
        }
    }
}

function GetPropertyId($propertyItem) {

    $propertyId = $property.Value.ExtendedFieldURI.PropertyTag

    if($propertyId -eq $null) {
        $propertyId = $property.Value.ExtendedFieldURI.PropertyId
    }

    if($propertyId -eq $null) {
        $propertyId = $property.Value.ExtendedFieldURI.PropertyName
    }

    return $propertyId
}

function AddContactPropertiesToHashSets ($propertyId, $propertyValue) {

    if(!$hashContactProperties.ContainsKey($propertyId)) {
        $propertyValueArray = @{}
        $propertyValueArray[$propertyValue] = ''
        $hashContactProperties[$propertyId] = $propertyValueArray
    }
    else
    {
        if(!$hashContactProperties[$propertyId].ContainsKey($propertyValue))
        {
            $hashContactProperties[$propertyId][$propertyValue] = ''
        }
    }
}

function GetColumnList()
{
    [object[]]$columns = "PersonaId"
    
    $columns += "XrmId"
    $columns += "DisplayName"
    $columns += "GivenName"
    $columns += "MiddleName"
    $columns += "LastName"
    $columns += "BirthDay"
    $columns += "JobTitle"
    $columns += "CompanyName"
    $columns += "IsBusinessContact"
    $columns += "Shared"
    $columns += "Email1Address"
    $columns += "Email1DisplayName"
    $columns += "Email1OriginalDisplayName"
    $columns += "Email2Address"
    $columns += "Email2DisplayName"
    $columns += "Email2OriginalDisplayName"
    $columns += "Email3Address"
    $columns += "Email3DisplayName"
    $columns += "Email3OriginalDisplayName"
    $columns += "WorkAddressStreet"
    $columns += "WorkAddressCity"
    $columns += "WorkAddressStateOrProvince"
    $columns += "WorkAddressPostalCode"
    $columns += "WorkAddressCountry"
    $columns += "HomeAddressStreet"
    $columns += "HomeAddressCity"
    $columns += "HomeAddressStateOrProvince"
    $columns += "HomeAddressPostalCode"
    $columns += "HomeAddressCountry"
    $columns += "OtherAddressStreet"
    $columns += "OtherAddressCity"
    $columns += "OtherAddressStateOrProvince"
    $columns += "OtherAddressPostalCode"
    $columns += "OtherAddressCountry"
    $columns += "BusinessTelephoneNumber"
    $columns += "MobileTelephoneNumber"
    $columns += "HomeTelephoneNumber"
    $columns += "OtherTelephoneNumber"
    $columns += "Notes"
    $columns += "CompanyLinks"
    $columns += "DealLinks"
    $columns += "ItemId"

    foreach($template in $contactTemplateData){
        $columns += $template.FieldLabel
    }

    return $columns
}

function GetCustomPropertiesList()
{
    $extendedPropsListForContacts = New-Object System.Collections.ArrayList

    foreach($template in $contactTemplateData) {
        $extendedPropsforCustomFields = "<t:ExtendedFieldURI PropertySetId='1a417774-4779-47c1-9851-e42057495fca' PropertyName=" + "'" + $template.FieldId + "'" + " PropertyType=" + "'" + $customFiedsTypeMapping[$template.FieldType] + "'" + " xmlns:t='http://schemas.microsoft.com/exchange/services/2006/types' />"
        $extendedPropsListForContacts.Add($extendedPropsforCustomFields) | Out-Null
    }

    $extendedPropsListForContacts = $extendedPropsListForContacts -join "`r`n"

    return $extendedPropsListForContacts
}

function InitializeContactProperties($obj)
{
    $obj | Add-Member -MemberType NoteProperty -Name GivenName -Value ''
    $obj | Add-Member -MemberType NoteProperty -Name LastName -Value ''
    $obj | Add-Member -MemberType NoteProperty -Name CompanyName -Value ''
    $obj | Add-Member -MemberType NoteProperty -Name MiddleName -Value ''
    $obj | Add-Member -MemberType NoteProperty -Name BirthDay -Value ''
    $obj | Add-Member -MemberType NoteProperty -Name BusinessHomePage -Value ''
    $obj | Add-Member -MemberType NoteProperty -Name JobTitle -Value ''
    $obj | Add-Member -MemberType NoteProperty -Name BusinessTelephoneNumber -Value ''
    $obj | Add-Member -MemberType NoteProperty -Name HomeTelephoneNumber -Value ''
    $obj | Add-Member -MemberType NoteProperty -Name MobileTelephoneNumber -Value ''
    $obj | Add-Member -MemberType NoteProperty -Name OtherTelephoneNumber -Value ''

    $obj | Add-Member -MemberType NoteProperty -Name Notes -Value ''
 
    $obj | Add-Member -MemberType NoteProperty -Name HomeAddressStreet -Value ''
    $obj | Add-Member -MemberType NoteProperty -Name HomeAddressCity -Value ''
    $obj | Add-Member -MemberType NoteProperty -Name HomeAddressStateOrProvince -Value ''
    $obj | Add-Member -MemberType NoteProperty -Name HomeAddressPostalCode -Value ''
    $obj | Add-Member -MemberType NoteProperty -Name HomeAddressCountry -Value ''
 
    $obj | Add-Member -MemberType NoteProperty -Name Email1Address -Value ''
    $obj | Add-Member -MemberType NoteProperty -Name Email1DisplayName -Value ''
    $obj | Add-Member -MemberType NoteProperty -Name Email1OriginalDisplayName -Value ''

    $obj | Add-Member -MemberType NoteProperty -Name Email2Address -Value ''
    $obj | Add-Member -MemberType NoteProperty -Name Email2DisplayName -Value ''
    $obj | Add-Member -MemberType NoteProperty -Name Email2OriginalDisplayName -Value ''

    $obj | Add-Member -MemberType NoteProperty -Name Email3Address -Value ''
    $obj | Add-Member -MemberType NoteProperty -Name Email3DisplayName -Value ''
    $obj | Add-Member -MemberType NoteProperty -Name Email3OriginalDisplayName -Value ''

    $obj | Add-Member -MemberType NoteProperty -Name WorkAddressStreet -Value ''
    $obj | Add-Member -MemberType NoteProperty -Name WorkAddressCity -Value ''
    $obj | Add-Member -MemberType NoteProperty -Name WorkAddressStateOrProvince -Value ''
    $obj | Add-Member -MemberType NoteProperty -Name WorkAddressPostalCode -Value ''
    $obj | Add-Member -MemberType NoteProperty -Name WorkAddressCountry -Value ''

    $obj | Add-Member -MemberType NoteProperty -Name OtherAddressStreet -Value ''
    $obj | Add-Member -MemberType NoteProperty -Name OtherAddressCity -Value ''
    $obj | Add-Member -MemberType NoteProperty -Name OtherAddressStateOrProvince -Value ''
    $obj | Add-Member -MemberType NoteProperty -Name OtherAddressPostalCode -Value ''
    $obj | Add-Member -MemberType NoteProperty -Name OtherAddressCountry -Value ''
    
    $obj | Add-Member -MemberType NoteProperty -Name CompanyLinks -Value ''
    $obj | Add-Member -MemberType NoteProperty -Name DealLinks -Value ''

    $obj | Add-Member -MemberType NoteProperty -Name  IsBusinessContact -Value 'False'
    $obj | Add-Member -MemberType NoteProperty -Name  XrmId -Value ''
    $obj | Add-Member -MemberType NoteProperty -Name  Shared -Value 'False'

    $labelArray = $contactTemplateData.FieldLabel
    foreach($label in $labelArray)
    {
        $IsPropertyAdded = $obj.PSobject.Properties.Name -contains $label
            
        if($IsPropertyAdded -eq $False) {
            $obj | Add-Member -MemberType NoteProperty -Name $label -Value ' '
        }
    }

    return $obj | Out-Null
}

function CallExportContactsData {
    try 
    {
        $personaArray = @()
        $hashContactProperties = @{}

        $findFolder = "<?xml version='1.0' encoding='utf-8'?>
        <soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:m='http://schemas.microsoft.com/exchange/services/2006/messages' xmlns:t='http://schemas.microsoft.com/exchange/services/2006/types' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>
            <soap:Header>
            <t:RequestServerVersion Version='V2017_07_11' />
            </soap:Header>
            <soap:Body>
            <m:FindFolder Traversal='Shallow'>
                <m:FolderShape>
                <t:BaseShape>IdOnly</t:BaseShape>
                </m:FolderShape>
                <m:IndexedPageFolderView MaxEntriesReturned='100' Offset='0' BasePoint='Beginning' />
                <m:Restriction>
                <t:IsEqualTo>
                    <t:FieldURI FieldURI='folder:DisplayName' />
                    <t:FieldURIOrConstant>
                    <t:Constant Value='MyOCMContacts' />
                    </t:FieldURIOrConstant>
                </t:IsEqualTo>
                </m:Restriction>
                <m:ParentFolderIds>
                <t:DistinguishedFolderId Id='root' />
                </m:ParentFolderIds>
            </m:FindFolder>
            </soap:Body>
        </soap:Envelope>"

        $response = Invoke-WebRequest $uri -Method Post -ContentType "text/xml" -Credential $cred -Body "$findFolder"
        $responseXML = [xml]$response
        $responseCode = (select-xml -Xml $responseXML -Namespace $ns -XPath "//m:ResponseCode").Node.InnerText

        $folderId = $null

        if ($responseCode -eq 'NoError')
        {
            $folderId = (select-xml -Xml $responseXML -Namespace $ns -XPath "//t:FolderId").Node
            $folderId = $folderId.Id

            if($folderId)
            {

            write-host -ForegroundColor Green "Downloading data for contacts..."
    
            # Get custom property details from template CSV
            $contactTemplateData = $null
            $filepath = ".\" + $folderName.BaseName + "\" + $hashFilepath['ContactTemplate']
            $contactTemplateData = $null
         
            if (Test-Path -Path $filepath)
            {
                $contactTemplateData = Import-Csv -Path $filepath
            }
    
            $pageSize = 100
            $pageOffset = 0
            $extendedPropsListForContacts = $null

            if($contactTemplateData)
            {
                $extendedPropsListForContacts = GetCustomPropertiesList
            }

            #
            # FindPeople request
            #
            $findPeopleDefault = "<?xml version='1.0' encoding='UTF-8'?>
            <soap:Envelope xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'
                            xmlns:t='http://schemas.microsoft.com/exchange/services/2006/types'
                            xmlns:m='http://schemas.microsoft.com/exchange/services/2006/messages'>
                <soap:Header>
                <t:RequestServerVersion Version='V2017_07_11' />
                </soap:Header>
                <soap:Body >
                <m:FindPeople>
                    <m:PersonaShape>
                    <t:BaseShape>Default</t:BaseShape>
                    </m:PersonaShape>
                    <m:IndexedPageItemView BasePoint='Beginning' MaxEntriesReturned='1' Offset='0' />
                    <m:ParentFolderId>
                    <t:FolderId Id='$folderId'/>
                    </m:ParentFolderId>
                </m:FindPeople>
                </soap:Body>
            </soap:Envelope>"

            $response = Invoke-WebRequest $uri -Method Post -ContentType "text/xml" -Credential $cred -Body "$findPeopleDefault"
            $responseXML = [xml]$response
            $responseCode = (select-xml -Xml $responseXML -Namespace $ns -XPath "//m:ResponseCode").Node.InnerText

            if ($responseCode -eq 'NoError') 
            {
                $totalCount = (select-xml -Xml $responseXML -Namespace $ns -XPath "//m:TotalNumberOfPeopleInView").Node.InnerText

                if($totalCount -eq 0)
                {
                    write-host -ForegroundColor Yellow "No data available for the contacts"
                    return
                }

                $noOfRequests = ([Math]::Ceiling($totalCount / 100))
                $pageOffset = 0

                while($noOfRequests -gt 0)
                {
                    #
                    # FindPeople request
                    #
                    $findPeople = "<?xml version='1.0' encoding='UTF-8'?>
                    <soap:Envelope xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'
                                    xmlns:t='http://schemas.microsoft.com/exchange/services/2006/types'
                                    xmlns:m='http://schemas.microsoft.com/exchange/services/2006/messages'>
                        <soap:Header>
                        <t:RequestServerVersion Version='V2017_07_11' />
                        </soap:Header>
                        <soap:Body >
                        <m:FindPeople>
                            <m:PersonaShape>
                            <t:BaseShape>Default</t:BaseShape>
                            <t:AdditionalProperties>
                                <t:FieldURI FieldURI='persona:InlineLinks'/>
                                <t:FieldURI FieldURI='persona:ItemLinkIds'/>
                                <t:FieldURI FieldURI='persona:Attributions' xmlns:t='http://schemas.microsoft.com/exchange/services/2006/types' />
                                <t:ExtendedFieldURI PropertyTag='0x1000' PropertyType='String' />
                                <t:ExtendedFieldURI PropertyTag='0x3a06' PropertyType='String' />
                                <t:ExtendedFieldURI PropertyTag='0x3a11' PropertyType='String' />
                                <t:ExtendedFieldURI PropertyTag='0x3a16' PropertyType='String' />
                                <t:ExtendedFieldURI PropertyTag='0x3a44' PropertyType='String' />
                                <t:ExtendedFieldURI PropertyTag='0x3a42' PropertyType='SystemTime' />
                                <t:ExtendedFieldURI PropertyTag='0x3a17' PropertyType='String' />
                                <t:ExtendedFieldURI PropertyTag='0x3a08' PropertyType='String' />
                                <t:ExtendedFieldURI PropertyTag='0x3a09' PropertyType='String' />
                                <t:ExtendedFieldURI PropertyTag='0x3a1c' PropertyType='String' />
                                <t:ExtendedFieldURI PropertyTag='0x3a1f' PropertyType='String' />
                                <t:ExtendedFieldURI DistinguishedPropertySetId='Address' PropertyId='32899' PropertyType='String' />
                                <t:ExtendedFieldURI DistinguishedPropertySetId='Address' PropertyId='32896' PropertyType='String' />
                                <t:ExtendedFieldURI DistinguishedPropertySetId='Address' PropertyId='32900' PropertyType='String' />
                                <t:ExtendedFieldURI DistinguishedPropertySetId='Address' PropertyId='32915' PropertyType='String' />
                                <t:ExtendedFieldURI DistinguishedPropertySetId='Address' PropertyId='32912' PropertyType='String' />
                                <t:ExtendedFieldURI DistinguishedPropertySetId='Address' PropertyId='32916' PropertyType='String' />
                                <t:ExtendedFieldURI DistinguishedPropertySetId='Address' PropertyId='32931' PropertyType='String' />
                                <t:ExtendedFieldURI DistinguishedPropertySetId='Address' PropertyId='32928' PropertyType='String' />
                                <t:ExtendedFieldURI DistinguishedPropertySetId='Address' PropertyId='32932' PropertyType='String' />
                                <t:ExtendedFieldURI PropertyTag='0x3a5d' PropertyType='String' />
                                <t:ExtendedFieldURI PropertyTag='0x3a59' PropertyType='String' />
                                <t:ExtendedFieldURI PropertyTag='0x3a5c' PropertyType='String' />
                                <t:ExtendedFieldURI PropertyTag='0x3a5b' PropertyType='String' />
                                <t:ExtendedFieldURI PropertyTag='0x3a5a' PropertyType='String' />
                                <t:ExtendedFieldURI DistinguishedPropertySetId='Address' PropertyId='32837' PropertyType='String' />
                                <t:ExtendedFieldURI DistinguishedPropertySetId='Address' PropertyId='32838' PropertyType='String' />
                                <t:ExtendedFieldURI DistinguishedPropertySetId='Address' PropertyId='32839' PropertyType='String' />
                                <t:ExtendedFieldURI DistinguishedPropertySetId='Address' PropertyId='32840' PropertyType='String' />
                                <t:ExtendedFieldURI DistinguishedPropertySetId='Address' PropertyId='32841' PropertyType='String' />
                                <t:ExtendedFieldURI PropertyTag='0x3a63' PropertyType='String' />
                                <t:ExtendedFieldURI PropertyTag='0x3a5f' PropertyType='String' />
                                <t:ExtendedFieldURI PropertyTag='0x3a62' PropertyType='String' />
                                <t:ExtendedFieldURI PropertyTag='0x3a61' PropertyType='String' />
                                <t:ExtendedFieldURI PropertyTag='0x3a60' PropertyType='String' />
                                <t:ExtendedFieldURI DistinguishedPropertySetId='Address' PropertyName='CustomerBit' PropertyType='Boolean' />
                                <t:ExtendedFieldURI PropertySetId='1a417774-4779-47c1-9851-e42057495fca' PropertyName='XrmSourceMailboxGuid' PropertyType='CLSID' />
                                $extendedPropsListForContacts
                            </t:AdditionalProperties>
                            </m:PersonaShape>
                            <m:IndexedPageItemView BasePoint='Beginning' MaxEntriesReturned='$pageSize' Offset='$pageOffset'/>
                            <m:ParentFolderId>
                            <t:FolderId Id='$folderId'/>
                            </m:ParentFolderId>
                            <m:SortOrder>
                            <t:FieldOrder Order='Descending'>
                                <t:FieldURI FieldURI='persona:LastModifiedTime' />
                            </t:FieldOrder>
                            </m:SortOrder>
                        </m:FindPeople>
                        </soap:Body>
                    </soap:Envelope>"

                    $pageOffset += $pageSize
            
                    $response = Invoke-WebRequest $uri -Method Post -ContentType "text/xml" -Credential $cred -Body "$findPeople"

                    $responseXML = [xml]$response
                    $responseCode = (select-xml -Xml $responseXML -Namespace $ns -XPath "//m:ResponseCode").Node.InnerText

                    if ($responseCode -eq 'NoError') {
                        $personaList += select-xml -Xml $responseXML -Namespace $ns -XPath "//t:Persona"
                        $personaList = @($personaList)
                    }
                    
                    write-progress -Activity "Downloading"-Status ("{0} of $totalCount Contacts"-f $personaList.Count) -PercentComplete (($personaList.Count/$totalCount)*100)
                    
                    $noOfRequests--
                }

                write-progress -Activity "Download complete" -Completed

                if($personaList) 
                {
                    write-host -ForegroundColor Green "Exporting data for contacts..."

                    foreach($persona in $personaList) {
                        $idx = $personaList.IndexOf($persona)
                        write-progress -Activity "Processing"-Status ("{0} of $totalCount Contacts"-f $idx) -PercentComplete (($idx/$totalCount)*100)
                        
                        $obj = New-Object -TypeName psobject
                        InitializeContactProperties $obj

                        $personaId = $persona.Node.PersonaId.Id
                        $obj | Add-Member -MemberType NoteProperty -Name PersonaId -Value $personaId

                        $displayName = $persona.Node.DisplayName
                        $obj | Add-Member -MemberType NoteProperty -Name DisplayName -Value $displayName

                        $attributions = $persona.Node.Attributions.Attribution

                        $hashItemIds = @{}
                        $hashChangeKeys = @{}

                        $itemIds = New-Object System.Collections.ArrayList

                        foreach($attribution in $attributions)
                        {
                            if($attribution.IsHidden -eq $false)
                            {
                                $itemIds.Add($attribution.SourceId.Id) | Out-Null
                            }

                            $hashItemIds[$attribution.Id] = $attribution.SourceId.Id
                            $hashChangeKeys[$attribution.Id] = $attribution.SourceId.ChangeKey
                        }

                        if($itemIds.Count -gt 1)
                        {
                            $itemId = $itemIds -join '|'
                        }

                        $obj | Add-Member -MemberType NoteProperty -Name ItemId -Value $itemId

                        $itemLinkIds = $persona.Node.ItemLinkIds.StringArrayAttributedValue

                        foreach($itemLink in $itemLinkIds) {
                            $itemLinkId = $itemLink.Values.Value
                            AddContactPropertiesToHashSets "ItemLinkId" $itemLinkId
                        }

                        $inlineLinks = $persona.Node.InlineLinks.StringAttributedValue
                          
                        foreach($link in $inlineLinks) {
                            $relationships = $link.Value | ConvertFrom-Json
                            $relationships = $relationships.Relationships
                            $companyObj = $relationships | where ItemType -eq "IPM.Contact.Company" | select ItemLinkId
                
                            if ($companyObj) {
                                $companyObj | ForEach-Object -Process { AddContactPropertiesToHashSets "CompanyLinks" $_.ItemLinkId}
                            }
                
                            $dealObj = $relationships | where ItemType -eq "IPM.XrmProject.Deal" | select ItemLinkId

                            if ($dealObj) {
                                $dealObj | ForEach-Object -Process { AddContactPropertiesToHashSets "DealLinks" $_.ItemLinkId}
                            }
                        }

                        $extendedPropertiesList = $persona.Node.ExtendedProperties.ExtendedPropertyAttributedValue

                        foreach($property in $extendedPropertiesList)
                        {
                            $extendedField = $property.Value.ExtendedFieldURI
                            $propertyId = GetPropertyId $extendedField
                            $propertyValue = $property.Value.Value

                            
                            if($propertyValue)
                            {
                                #Handle 1601-01-01T00:00:00Z date format for birthday and custom date fields
                                
                                if($propertyId -eq "0x3a42" -and $propertyValue -eq $defaultUTCDateTime)
                                {
                                    $propertyValue = ''
                                }

                                $fieldType = $null
                                $fieldType = $contactTemplateData | Where-Object {$_.FieldId -eq $propertyId} | Select-Object -Property FieldType
                                
                                if($fieldType)
                                {
                                    if($fieldType.FieldType -eq "Date" -and $propertyValue -eq $defaultUTCDateTime)
                                    {
                                        $propertyValue = ''
                                    }
                                }
                            }

                            # Handle data truncation issue for Notes field
                            
                            if($propertyId -eq "0x1000" -and $propertyValue.Length -gt 254)
                            {
                                $notesAttributions = $property.Attributions.Attribution

                                foreach($notesAttributionId in $notesAttributions)
                                {
                                    $contactItemId = $hashItemIds[$notesAttributionId]
                                    $contactChangeKey = $hashChangeKeys[$notesAttributionId]

                                    break
                                }

                                $getItem = "<?xml version='1.0' encoding='utf-8'?>
                                <soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/' xmlns:t='http://schemas.microsoft.com/exchange/services/2006/types'>
                                    <soap:Header>
                                    <RequestServerVersion Version='V2017_07_11' xmlns='http://schemas.microsoft.com/exchange/services/2006/types' />
                                    </soap:Header>
                                    <soap:Body>
                                    <GetItem xmlns='http://schemas.microsoft.com/exchange/services/2006/messages'>
                                        <ItemShape>
                                        <t:BaseShape>IdOnly</t:BaseShape>
                                        <t:AdditionalProperties>
                                            <t:FieldURI FieldURI='item:Body' />
                                        </t:AdditionalProperties>
                                        </ItemShape>
                                        <ItemIds>
                                        <ItemId Id='$contactItemId' ChangeKey='$contactChangeKey' xmlns='http://schemas.microsoft.com/exchange/services/2006/types' />
                                        </ItemIds>
                                    </GetItem>
                                    </soap:Body>
                                </soap:Envelope>"

                                try
                                {
                                    $response = Invoke-WebRequest $uri -Method Post -ContentType "text/xml" -Credential $cred -Body "$getItem"

                                    $responseXML = [xml]$response
                                    $responseCode = (select-xml -Xml $responseXML -Namespace $ns -XPath "//m:ResponseCode").Node.InnerText

                                    if ($responseCode -eq 'NoError') {
                                        $propertyValue = (select-xml -Xml $responseXML -Namespace $ns -XPath "//t:Body").Node.InnerText
                                    }
                                }
                                catch {
                                    write-host -ForegroundColor Red "GetItem request call to fetch contact notes failed with error -"
                                    $_.Exception.Message
                                }
                            }
                            
                            AddContactPropertiesToHashSets $propertyId $propertyValue
                        }

                        foreach($hashContactProperty in $hashContactProperties.GetEnumerator()){
                
                            if($hashContactProperty.Key -eq 'XrmSourceMailboxGuid') {
                                $sharedMailbox = 'True'
                                foreach($item in $hashContactProperty.Value) {
                                    foreach($key in $item.Keys)
                                    {
                                        if($key -eq  '00000000-0000-0000-0000-000000000000' -or $key -eq  $null) {
                                            $sharedMailbox = 'False'
                                            break
                                        }
                                    }
                        
                                }
                                $obj | Add-Member -MemberType NoteProperty -Name Shared -Value $sharedMailbox -Force
                            }
                            elseif ($hashContactProperty.Key -eq 'CustomerBit') {
                                $isBusinessContact = 'False'
                                foreach($item in $hashContactProperty.Value) {
                                    foreach($key in $item.Keys)
                                    {
                                        if($key -eq $true)
                                        {
                                            $isBusinessContact =  'True'
                                            break
                                        }
                                    }
                                }
                                $obj | Add-Member -MemberType NoteProperty -Name IsBusinessContact -Value $isBusinessContact -Force
                            }
                            else
                            {
                                $idLabel = $hashContactPropertyIdToText[$hashContactProperty.Key]
                                if($idLabel -eq $null -and $contactTemplateData -ne $null) {
                                    $idLabel = $contactTemplateData | Where-Object {$_.FieldId -eq $hashContactProperty.Key} | Select-Object -Property FieldLabel
                                    $idLabel = $idLabel.FieldLabel
                                }
                    
                                if($idLabel -ne $null) {
                                    AddPropertyObject $obj $hashContactProperty.Value $idLabel
                                }
                    
                            }
                        }
            
                        $hashContactProperties.Clear()
                        $personaArray += $obj
                }

                $columns = GetColumnList
    
                $sortedList = $personaArray |  Select-Object -property $columns
                if($sortedList) {
                    $filepath = ".\" + $folderName.BaseName + "\" + $hashFilepath['Contact'] 
                    $sortedList | Export-Csv  -Path  $filepath -NoTypeInformation
                }

                write-host -ForegroundColor Green "Exporting data for contacts complete"
            }
            }
            else {
                write-host -ForegroundColor Yellow "No data available for the contacts"
            }
        }
        else
        {
            write-host -ForegroundColor Red "Failed to get the folder id for contacts..."
        }
    }
   } 
   catch {
        write-host -ForegroundColor Green "Exporting data for contacts failed"
        $_.Exception.Message
   }
}

function ParseCompanydata($dataList)
{
    $companyFieldArray = @()
    $totalCount = $dataList.value.length

    foreach ($dataObj in $dataList.value) {

        $idx = $dataList.value.IndexOf($dataObj)
        write-progress -Activity "Processing"-Status ("{0} of $totalCount Companies"-f $idx) -PercentComplete (($idx/$totalCount)*100)

        $obj = New-Object -TypeName psobject
        
        $obj | Add-Member -MemberType NoteProperty -Name XrmId -Value $dataObj.XrmId
        $obj | Add-Member -MemberType NoteProperty -Name DisplayName -Value $dataObj.DisplayName
        $obj | Add-Member -MemberType NoteProperty -Name BusinessHomePage -Value $dataObj.BusinessHomePage

        $obj | Add-Member -MemberType NoteProperty -Name BusinessPhone -Value $dataObj.BusinessPhones[0]

        $emailAddresses = $dataObj.EmailAddresses[0].Address
        $obj | Add-Member -MemberType NoteProperty -Name EmailAddress -Value $emailAddresses

        $obj | Add-Member -MemberType NoteProperty -Name BusinessAddressStreet -Value $dataObj.BusinessAddress.Street
        $obj | Add-Member -MemberType NoteProperty -Name BusinessAddressCity -Value $dataObj.BusinessAddress.City
        $obj | Add-Member -MemberType NoteProperty -Name BusinessAddressState -Value $dataObj.BusinessAddress.State
        $obj | Add-Member -MemberType NoteProperty -Name BusinessAddressPostalCode -Value $dataObj.BusinessAddress.PostalCode
        $obj | Add-Member -MemberType NoteProperty -Name BusinessAddressCountry/Region -Value $dataObj.BusinessAddress.CountryOrRegion

        $obj | Add-Member -MemberType NoteProperty -Name Notes -Value $dataObj.Notes

        if($dataObj.SourceMailboxGuid -eq  '00000000-0000-0000-0000-000000000000' -or $dataObj.SourceMailboxGuid -eq  $null) {
            $shared = 'False'
        }
        else {
            $shared = 'True'
        }

        $obj | Add-Member -MemberType NoteProperty -Name Shared -Value $shared
        $obj | Add-Member -MemberType NoteProperty -Name ItemId -Value $dataObj.Id


        $IsPropertyAdded = $obj.PSobject.Properties.Name -contains 'ContactLinks'
            
        if($IsPropertyAdded -eq $False) {
            $obj | Add-Member -MemberType NoteProperty -Name ContactLinks -Value ''
        }

        $IsPropertyAdded = $obj.PSobject.Properties.Name -contains 'DealLinks'
            
        if($IsPropertyAdded -eq $False) {
            $obj | Add-Member -MemberType NoteProperty -Name DealLinks -Value ''
        }

        $inlineLinks = $dataObj.InlineLinks.Relationships

        if($inlineLinks) {
            $contactLinks = New-Object System.Collections.ArrayList
            $dealLinks = New-Object System.Collections.ArrayList

            foreach($link in $inlineLinks)
            {
                if($link.ItemType -eq 'IPM.AbchPerson')
                {
                    $contactLinks.Add(($link.ItemLinkId)) | Out-Null
                }

                if($link.ItemType -eq 'IPM.XrmProject.Deal')
                {
                    $dealLinks.Add($link.ItemLinkId) | Out-Null
                }
            }

            if($contactLinks) {

                $contactLinks = $contactLinks -join '|'
                 $obj | Add-Member -MemberType NoteProperty -Name  ContactLinks -value $contactLinks -Force
            }
        
            if($dealLinks) {
                $dealLinks = $dealLinks -join '|'
                $obj | Add-Member -MemberType NoteProperty -Name DealLinks -value $dealLinks -Force
            }
        }

        #
        # Start with adding all the custom Field Labels to object
        #
        $labelArray = $templateData.FieldLabel
        foreach($label in $labelArray)
        {
            $IsPropertyAdded = $obj.PSobject.Properties.Name -contains $label
            
            if($IsPropertyAdded -eq $False) {
                $obj | Add-Member -MemberType NoteProperty -Name $label -Value ''
            }
        }

        $singleValueProperties = $dataObj.SingleValueExtendedProperties;

        $propertyLabel = $null

        foreach($prop in $singleValueProperties)
        {
            $propertyId = $prop.PropertyId.Split(' ')[3].TrimStart('{').TrimEnd('}')

            $propertyLabel = $templateData | Where-Object {$_.FieldId -eq $propertyId} | Select-Object -Property FieldLabel
            $propertyValue = $prop.Value

            #Handle 1601-01-01T00:00:00Z date format for custom date fields

            if($propertyValue)
            {
                $fieldType = $null
                $fieldType = $templateData | Where-Object {$_.FieldId -eq $propertyId} | Select-Object -Property FieldType

                if($fieldType)
                {
                    if($fieldType.FieldType -eq "Date" -and $propertyValue -eq $defaultUTCDateTime)
                    {
                        $propertyValue = ''
                    }
                }
            }
           

            $obj | Add-Member -MemberType NoteProperty -Name $propertyLabel.FieldLabel -Value $propertyValue -Force
        }

        $companyFieldArray += $obj
    }

    return $companyFieldArray;
}

function CallExportCompaniesData {
    #
    # Make a REST call for company templates
    #
    try {
        
        # Get custom property details from template CSV
        $filepath = ".\" + $folderName.BaseName + "\" + $hashFilepath['CompanyTemplate']
        $templateData = $null

        if (Test-Path -Path $filepath)
        {
            $templateData = Import-Csv -Path $filepath
        }

        if($templateData)
        {
            $extendedPropsList = New-Object System.Collections.ArrayList

            foreach($template in $templateData) {
                $extendedPropsforCustomFields = "(PropertyId eq " + "'" + $customFiedsTypeMapping[$template.FieldType] + " " + "{" + "1a417774-4779-47c1-9851-e42057495fca" + "}" + " Name " + $template.FieldId + "'" + ")"
                $extendedPropsList.Add($extendedPropsforCustomFields) | Out-Null
            }

            $extendedPropsList = $extendedPropsList -join " OR "
        
            $extendedPropertiesRequest = "$" + "expand" + "=" + "SingleValueExtendedProperties" + "(" + "$" + "filter" + "=" + $extendedPropsList + ")"
            $companyRequestURL = $baseURL + $entityTypeOrg + $filterTop + $pageSize + $filterSkip +$skipCount + $filterCount + $filterOrderByOrg + "&" + $extendedPropertiesRequest
        }
        else
        {
            $companyRequestURL = $baseURL + $entityTypeOrg + $filterTop + $pageSize + $filterSkip +$skipCount + $filterCount + $filterOrderByOrg
        }

        $companyData = @()

        write-host -ForegroundColor Green "Downloading data for companies..."

        $data = Invoke-WebRequest -Method Get -uri $companyRequestURL -Headers @{ 'Authorization' = "Bearer $token" } 
        $companyData = $data.Content| ConvertFrom-Json | Select-Object -Property  "value"

        $totalCount = $data.Content| ConvertFrom-Json | Select -ExpandProperty  "@odata.count"

        if($totalCount -eq 0)
        {
            write-host -ForegroundColor Yellow "No data available for the companies"
            return
        }

        write-progress -Activity "Download progress"-Status ("{0} of $totalCount"-f $companyData.Value.Count) -PercentComplete (($companyData.Value.Count/$totalCount)*100)

        $noOfRequests = ([Math]::Ceiling($totalCount / 100) - 1)
        
        while($noOfRequests -gt 0)
        {
            $skipCount += $pageSize
            $companyRequestURL = $baseURL + $entityTypeOrg + $filterTop + $pageSize + $filterSkip +$skipCount + $filterCount + $filterOrderByOrg + "&" + $extendedPropertiesRequest
            $data = Invoke-WebRequest -Method Get -uri $companyRequestURL -Headers @{ 'Authorization' = "Bearer $token" }

            $companyData = @($companyData )
            $companyData += $data.Content| ConvertFrom-Json | Select-Object -Property  "value"

            write-progress -Activity "Download progress"-Status ("{0} of $totalCount"-f $companyData.Value.Count) -PercentComplete (($companyData.Value.Count/$totalCount)*100)

            $noOfRequests--
        }

        write-progress -Activity "Download complete" -Completed

        if($companyData) {
            write-host -ForegroundColor Green "Exporting data for companies..."

            $parseData = @()
            $parseData = ParseCompanydata $companyData
                
            if($parseData) {
                $filepath = ".\" + $folderName.BaseName + "\" + $hashFilepath['Company'] 
                $parseData | Export-Csv  -Path  $filepath -NoTypeInformation
            }
        }
        else {
            write-host -ForegroundColor Yellow "No data available for the companies"
        }

        write-host -ForegroundColor Green "Exporting data for companies complete"

    } catch {
        write-host -ForegroundColor Green "Exporting data for companies failed"
        $_.Exception.Message
    }
}

function ParseDealsData($dataList)
{
    $dealsFieldArray = @()
    $totalCount = $dataList.value.length

    foreach ($dataObj in $dataList.value) {

        $idx = $dataList.value.IndexOf($dataObj)
        write-progress -Activity "Processing"-Status ("{0} of $totalCount Deals"-f $idx) -PercentComplete (($idx/$totalCount)*100)

        $obj = New-Object -TypeName psobject
        
        $obj | Add-Member -MemberType NoteProperty -Name XrmId -Value $dataObj.XrmId
        $obj | Add-Member -MemberType NoteProperty -Name DealName -Value $dataObj.Name
        $obj | Add-Member -MemberType NoteProperty -Name Amount -Value ("USD {0:n2}" -f $dataObj.Amount)
        $obj | Add-Member -MemberType NoteProperty -Name Priority -Value $dataObj.Priority
        $obj | Add-Member -MemberType NoteProperty -Name Status -Value $dataObj.Status

        $stage = $null
        if($dataObj.Stage -ne $null -and $hashStageIdToLabel.ContainsKey($dataObj.Stage))
        {
            $stage = $hashStageIdToLabel[$dataObj.Stage]
        }
        else
        {
            $stage = $dataObj.Stage
        }
        
        $obj | Add-Member -MemberType NoteProperty -Name Stage -Value $stage

        $closeDate = $null
        if ($dataObj.CloseTime)
        {
            $closeDate = ([datetime]::Parse($dataObj.CloseTime)).ToString("yyyy-MM-dd")

            if($defaultDateTime.Contains($closeDate))
            {
                $closeDate = ''
            }
        }

        $obj | Add-Member -MemberType NoteProperty -Name CloseDate -Value $closeDate
        $obj | Add-Member -MemberType NoteProperty -Name Owner -Value $dataObj.Owner
        $obj | Add-Member -MemberType NoteProperty -Name Shared -Value ($dataObj.SourceMailboxGuid -eq  '00000000-0000-0000-0000-000000000000' -or $dataObj.SourceMailboxGuid -eq  $null)
        $obj | Add-Member -MemberType NoteProperty -Name Notes -Value $dataObj.Notes

        $companyLink = $null
        $contactLinks = $null
        $inlineLinks = $dataObj.InlineLinks.Relationships

        if($inlineLinks) 
        {
            $companyObj = $inlineLinks | where ItemType -eq "IPM.Contact.Company" | select ItemLinkId
            if ($companyObj) {
                $companyLink = $companyObj.ItemLinkId -join "|"
            }

            $contactObjs = $inlineLinks | where ItemType -eq "IPM.AbchPerson" | select ItemLinkId

            if ($contactObjs) {
                $contactLinks = $contactObjs.ItemLinkId -join "|"
            }
        }

        $obj | Add-Member -MemberType NoteProperty -Name CompanyLink -Value $companyLink
        $obj | Add-Member -MemberType NoteProperty -Name ContactLinks -Value $contactLinks
        $obj | Add-Member -MemberType NoteProperty -Name ItemId -Value $dataObj.Id

        #
        # Start with adding all the custom Filed Labels to object
        #

        $labelArray = $dealTemplateData.FieldLabel

        foreach($label in $labelArray)
        {
            $IsPropertyAdded = $obj.PSobject.Properties.Name -contains $label
            
            if($IsPropertyAdded -eq $False) {
                $obj | Add-Member -MemberType NoteProperty -Name $label -Value ''
            }
        }

        if ($dataObj.SingleValueExtendedProperties) 
        {
            $singleValueProperties = $dataObj.SingleValueExtendedProperties;

            $propertyLabel = $null

            foreach($prop in $singleValueProperties)
            {
                $propertyId = $prop.PropertyId.Split(' ')[3]
                $propertyLabel = $dealTemplateData | Where-Object {$_.FieldId -eq $propertyId} | Select-Object -Property FieldLabel
                $propertyValue = $prop.Value

                #Handle 1601-01-01T00:00:00Z date format for custom date fields

                if($propertyValue)
                {
                    $fieldType = $null
                    $fieldType = $dealTemplateData | Where-Object {$_.FieldId -eq $propertyId} | Select-Object -Property FieldType

                    if($fieldType)
                    {
                        if($fieldType.FieldType -eq "Date" -and $propertyValue -eq $defaultUTCDateTime)
                        {
                            $propertyValue = ''
                        }
                    }
                }
                
                $obj | Add-Member -MemberType NoteProperty -Name $propertyLabel.FieldLabel -Value $propertyValue -Force
            }
        }

        $dealsFieldArray += $obj
    }

    return $dealsFieldArray;
}

function CallExportDealsData {
    #
    # Make a REST call for Deals templates
    #
    try {
        
        
        # Get custom property details from template CSV

        $filepath = ".\" + $folderName.BaseName + "\" + $hashFilepath['DealTemplate']
        $dealTemplateData = $null

        if (Test-Path -Path $filepath)
        {
            $dealTemplateData = Import-Csv -Path $filepath
        }

        if($dealTemplateData)
        {
            $extendedPropsList = New-Object System.Collections.ArrayList

            foreach($template in $dealTemplateData) {
                $extendedPropsforCustomFields = "(PropertyId eq " + "'" + $customFiedsTypeMapping[$template.FieldType] + " " + "{" + "1a417774-4779-47c1-9851-e42057495fca" + "}" + " Name " + $template.FieldId + "'" + ")"
                $extendedPropsList.Add($extendedPropsforCustomFields) | Out-Null
            }

            $extendedPropsList = $extendedPropsList -join " OR "
            $extendedPropertiesRequest = "$" + "expand" + "=" + "SingleValueExtendedProperties" + "(" + "$" + "filter" + "=" + $extendedPropsList + ")"
            $dealsRequestURL = $baseURL + $entityTypeDeal + $filterTop + $pageSize + $filterSkip +$skipCount + $filterCount + $filterOrderByDeal + "&" + $extendedPropertiesRequest
        }
        else
        {
            $dealsRequestURL = $baseURL + $entityTypeDeal + $filterTop + $pageSize + $filterSkip +$skipCount + $filterCount + $filterOrderByDeal + "&" + $extendedPropertiesRequest
        }
        
        $dealData = @()

        write-host -ForegroundColor Green "Downloading data for deals..."

        $data = Invoke-WebRequest -Method Get -uri $dealsRequestURL -Headers @{ 'Authorization' = "Bearer $token" } 
        $dealsData = $data.Content| ConvertFrom-Json | Select-Object -Property  "value"

        $totalCount = $data.Content| ConvertFrom-Json | Select -ExpandProperty  "@odata.count"
        if($totalCount -eq 0)
        {
            write-host -ForegroundColor Yellow "No data available for the deals"
            return
        }

        if ($dealsData)
        {
            write-progress -Activity "Download progress"-Status ("{0} of $totalCount Deals"-f $dealsData.Value.Count) -PercentComplete (($dealsData.Value.Count/$totalCount)*100)
        }

        $noOfRequests = ([Math]::Ceiling($totalCount / 100) - 1)
        
        while($noOfRequests -gt 0)
        {
            $skipCount += $pageSize
            $dealsRequestURL = $baseURL + $entityTypeDeal + $filterTop + $pageSize + $filterSkip +$skipCount + $filterCount + $filterOrderByDeal + "&" + $extendedPropertiesRequest
            $data = Invoke-WebRequest -Method Get -uri $dealsRequestURL -Headers @{ 'Authorization' = "Bearer $token" }

            $dealsData = @($dealsData )
            $dealsData += $data.Content| ConvertFrom-Json | Select-Object -Property  "value"
            write-progress -Activity "Download progress"-Status ("{0} of $totalCount Deals"-f $dealsData.Value.Count) -PercentComplete (($dealsData.Value.Count/$totalCount)*100)

            $noOfRequests--
        }

        write-progress -Activity "Download complete" -Completed

        if($dealsData) {
            write-host -ForegroundColor Green "Exporting data for deals..."

            $parseData = @()
            $parseData = ParseDealsData $dealsData
                
            if($parseData) {
                $filepath = ".\" + $folderName.BaseName + "\" + $hashFilepath['Deal'] 
                $parseData | Export-Csv  -Path  $filepath -NoTypeInformation
            }
        }
        else {
            write-host -ForegroundColor Yellow "No data available for the deals"
        }

        write-host -ForegroundColor Green "Exporting data for deals complete"

    } catch {
        write-host -ForegroundColor Green "Exporting data for deals failed"
        $_.Exception.Message
    }
}

function CallExportTasksData {

    try {

        $taskObjArray = @()

        $syncFolderItems = $syncFolderItemsStart + $syncFolderItemsEnd

        write-host -ForegroundColor Green "Downloading data for tasks..."

        $response = Invoke-WebRequest $uri -Method Post -ContentType "text/xml" -Credential $cred -Body "$syncFolderItems"
        $responseXML = [xml]$response

        $responseCode = (select-xml -Xml $responseXML -Namespace $ns -XPath "//m:ResponseCode").Node.InnerText

        if ($responseCode -eq 'NoError') {
    
            $taskItems = (select-xml -Xml $responseXML -Namespace $ns -XPath "//t:*[t:ItemId]")

            if ($taskItems)
            {
                write-progress -Activity "Download progress"-Status ("{0} Tasks"-f $taskItems.Count) 
            }

            $includesLastItemInRange = (select-xml -Xml $responseXML -Namespace $ns -XPath "//m:IncludesLastItemInRange").Node.InnerText

            if($includesLastItemInRange -eq $false)
            {
                $syncStateValue = (select-xml -Xml $responseXML -Namespace $ns -XPath "//m:SyncState").Node.InnerText
            }

            while($includesLastItemInRange -eq $false)
            {
                $synsState = "<m:SyncState>" + $syncStateValue + "</m:SyncState>"
                $syncFolderItems = $syncFolderItemsStart + $synsState +  $syncFolderItemsEnd

                $response = Invoke-WebRequest $uri -Method Post -ContentType "text/xml" -Credential $cred -Body "$syncFolderItems"
                $responseXML = [xml]$response

                $responseCode = (select-xml -Xml $responseXML -Namespace $ns -XPath "//m:ResponseCode").Node.InnerText

                if ($responseCode -eq 'NoError') {
            
                    $includesLastItemInRange = (select-xml -Xml $responseXML -Namespace $ns -XPath "//m:IncludesLastItemInRange").Node.InnerText

                    if($includesLastItemInRange -eq $false)
                    {
                        $syncStateValue = (select-xml -Xml $responseXML -Namespace $ns -XPath "//m:SyncState").Node.InnerText
                    }
                }

                $taskItems += select-xml -Xml $responseXML -Namespace $ns -XPath "//t:*[t:ItemId]"

                write-progress -Activity "Download progress"-Status ("{0} Tasks"-f $taskItems.Count) 
            }

            write-progress -Activity "Download complete" -Completed

            foreach($task in $taskItems)
            {
                $itemId = $task.Node.ItemId.Id
                $changeKey = $task.Node.ItemId.ChangeKey

                $GetItem = "<?xml version='1.0' encoding='utf-8'?>
                <soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:m='http://schemas.microsoft.com/exchange/services/2006/messages' xmlns:t='http://schemas.microsoft.com/exchange/services/2006/types' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>
                    <soap:Header>
                    <t:RequestServerVersion Version='V2017_04_14' />
                    </soap:Header>
                    <soap:Body>
                    <m:GetItem>
                        <m:ItemShape>
                        <t:BaseShape>AllProperties</t:BaseShape>
                        <t:AdditionalProperties>
                           <t:ExtendedFieldURI PropertySetId='1a417774-4779-47c1-9851-e42057495fca' PropertyName='XrmId' PropertyType='CLSID' />
                           <t:ExtendedFieldURI PropertySetId='1a417774-4779-47c1-9851-e42057495fca' PropertyName='InlineLinks' PropertyType='String' />
                           <t:ExtendedFieldURI PropertyTag='38' PropertyType='Integer' xmlns:t='http://schemas.microsoft.com/exchange/services/2006/types' />
                           <t:FieldURI FieldURI='item:TextBody' xmlns:t='http://schemas.microsoft.com/exchange/services/2006/types' />
                         </t:AdditionalProperties>
                        </m:ItemShape>
                        <m:ItemIds>
                         <t:ItemId Id='$itemId' ChangeKey='$changeKey' />
                        </m:ItemIds>
                    </m:GetItem>
                    </soap:Body>
                </soap:Envelope>"

                $response = Invoke-WebRequest $uri -Method Post -ContentType "text/xml" -Credential $cred -Body "$GetItem"
                $responseXML = [xml]$response

                $responseCode = (select-xml -Xml $responseXML -Namespace $ns -XPath "//m:ResponseCode").Node.InnerText

                if ($responseCode -eq 'NoError') {

                    $taskProperties = select-xml -Xml $responseXML -Namespace $ns -XPath "//m:*[t:Task]"

                    $obj = New-Object -TypeName psobject

                    $IsPropertyAdded = $obj.PSobject.Properties.Name -contains 'ParentType'
            
                    if($IsPropertyAdded -eq $False) {
                        $obj | Add-Member -MemberType NoteProperty -Name ParentType -Value ''
                    }

                    $IsPropertyAdded = $obj.PSobject.Properties.Name -contains 'ParentLink'
            
                    if($IsPropertyAdded -eq $False) {
                        $obj | Add-Member -MemberType NoteProperty -Name ParentLink -Value ''
                    }

                    $taskExtProperties = select-xml -Xml $responseXML -Namespace $ns -XPath "//t:ExtendedProperty"

                    foreach($taskExt in $taskExtProperties)
                    {
                        $propertyName = $taskExt.Node.ExtendedFieldURI.PropertyName

                        $propertyTag = $taskExt.Node.ExtendedFieldURI.PropertyTag

                        if($propertyTag -and $propertyTag -eq "0x26") {
                            $priority = $hashPriority[$taskExt.Node.Value]
                        }

                        if($propertyName -eq "XrmId")
                        {
                            $xrmId = $taskExt.Node.Value
                        }

                        if($propertyName -eq "InlineLinks")
                        {
                            $inlineLinks = $taskExt.Node.Value | ConvertFrom-Json

                            if($inlineLinks) {
                                foreach($link in $inlineLinks.Relationships)
                                {
                                    if($link.ItemType -eq 'IPM.AbchPerson' -or $link.ItemType -eq 'IPM.Contact')
                                    {
                                        $parentType = "Contact"
                                    }

                                    if($link.ItemType -eq 'IPM.Contact.Company')
                                    {
                                        $parentType = "Company"
                                    }

                                    if($link.ItemType -eq 'IPM.XrmProject.Deal')
                                    {
                                        $parentType = "Deal"
                                    }

                                    $parentLink = $link.ItemLinkId
                                }
                            }
                        }
                    }

                    $obj | Add-Member -MemberType NoteProperty -Name XrmID -Value $xrmId
                    $obj | Add-Member -MemberType NoteProperty -Name Subject -Value $taskProperties.Node.Task.Subject

                    if($taskProperties.Node.Task.StartDate -ne $null) {
                        $startDate = ([datetime]::Parse($taskProperties.Node.Task.StartDate)).ToString("yyyy-MM-dd")

                        if($defaultDateTime.Contains($startDate))
                        {
                            $startDate = ''
                        }
                    }
            
                    $obj | Add-Member -MemberType NoteProperty -Name StartDate -Value $startDate

                    if($taskProperties.Node.Task.DueDate -ne $null) {
                        $dueDate = ([datetime]::Parse($taskProperties.Node.Task.DueDate)).ToString("yyyy-MM-dd")

                        if($defaultDateTime.Contains($dueDate))
                        {
                            $dueDate = ''
                        }
                    }

                    $obj | Add-Member -MemberType NoteProperty -Name DueDate -Value $dueDate

                    $dateTimeCreated = $taskProperties.Node.Task.DateTimeCreated
                    if($dateTimeCreated -eq $defaultUTCDateTime)
                    {
                        $dateTimeCreated = ''
                    }

                    $obj | Add-Member -MemberType NoteProperty -Name DateTimeCreated -Value $dateTimeCreated

                    $lastModifiedTime = $taskProperties.Node.Task.LastModifiedTime
                    if($lastModifiedTime -eq $defaultUTCDateTime)
                    {
                        $lastModifiedTime = ''
                    }

                    $obj | Add-Member -MemberType NoteProperty -Name LastModifiedTime -Value $lastModifiedTime
                    $obj | Add-Member -MemberType NoteProperty -Name ReminderIsSet -Value $taskProperties.Node.Task.ReminderIsSet

                    $reminderDueBy = $taskProperties.Node.Task.ReminderDueBy
                    if($reminderDueBy -eq $defaultUTCDateTime)
                    {
                        $reminderDueBy = ''
                    }

                    $obj | Add-Member -MemberType NoteProperty -Name ReminderDueBy -Value $reminderDueBy
                    $obj | Add-Member -MemberType NoteProperty -Name Priority -Value $priority
                    $obj | Add-Member -MemberType NoteProperty -Name Status -Value $taskProperties.Node.Task.Status

                    $obj | Add-Member -MemberType NoteProperty -Name PercentComplete -Value $taskProperties.Node.Task.PercentComplete

                    $notes = $taskProperties.Node.Task.TextBody.InnerText

                    if($notes) {
                        $obj | Add-Member -MemberType NoteProperty -Name Notes -Value $notes
                    }
                    else
                    {
                        $obj | Add-Member -MemberType NoteProperty -Name Notes -Value ""
                    }
    
                    $obj | Add-Member -MemberType NoteProperty -Name  ParentType -value $parentType -Force
                    $obj | Add-Member -MemberType NoteProperty -Name  ParentLink -value $parentLink -Force
                    $obj | Add-Member -MemberType NoteProperty -Name ItemId -Value $taskProperties.Node.Task.ItemId.Id

                    $taskObjArray += $obj
                }
                else {
                    write-host -ForegroundColor Red "GetItem for task item failed with error '$responseCode'"
                }
            }

            if($taskObjArray) {
                   write-host -ForegroundColor Green "Exporting data for tasks..."
                   $filepath = ".\" + $folderName.BaseName + "\" + $hashFilepath['Task'] 
                   $taskObjArray | Export-Csv  -Path  $filepath -NoTypeInformation
             }
        }
        else {
            write-host -ForegroundColor Red "SyncFolderItems failed with error '$responseCode'"
            break
        }

        write-host -ForegroundColor Green "Exporting data for tasks complete"
    }catch {
        write-host -ForegroundColor Red "Exporting data for tasks failed -"
            $_.Exception.Message
        break
    }
}

function ParsePostsData($dataList)
{
    $postsFieldArray = @()
    $totalCount = $dataList.value.length

    foreach ($dataObj in $dataList) 
    {
        $idx = $dataList.IndexOf($dataObj)
        write-progress -Activity "Processing"-Status ("{0} of $totalCount Posts"-f $idx) -PercentComplete (($idx/$totalCount)*100)

        $obj = New-Object -TypeName psobject
        
        $obj | Add-Member -MemberType NoteProperty -Name XrmId -Value $dataObj.XrmId
        $obj | Add-Member -MemberType NoteProperty -Name PostType -Value $dataObj.Subtype
        $obj | Add-Member -MemberType NoteProperty -Name PostText -Value $dataObj.Text

        $eventDateTime = $null
        if ($dataObj.EventTime)
        {
            $eventDateTime = ([datetime]::Parse($dataObj.EventTime)).ToString($dateTimeFormat)

            if($eventDateTime -eq $defaultUTCDateTime)
            {
                $eventDateTime =''
            }
        }

        $obj | Add-Member -MemberType NoteProperty -Name EventDateTime -Value $eventDateTime

        $parentType = $null
        $parentLink = $null
        $inlineLinks = $dataObj.InlineLinks.Relationships | where LinkType -eq "ActivityPertainsTo"

        if($inlineLinks) 
        {
            ## For posts there will be a single parent
            $linkObj = $inlineLinks[$inlineLinks.Count -1]
            $parentType = $linkObj.ItemType | %{ if ($_ -eq "IPM.XRMProject.Deal") {"Deal"} else {if ($_ -eq "IPM.Contact.Company") {"Company"} else { if ($_ -eq "IPM.AbchPerson") {"Contact"} else {$_}}}}
            $parentLink = $linkObj.ItemLinkId
        }

        $obj | Add-Member -MemberType NoteProperty -Name ParentType -Value $parentType
        $obj | Add-Member -MemberType NoteProperty -Name ParentLink -Value $parentLink
        $obj | Add-Member -MemberType NoteProperty -Name ItemId -Value $dataObj.Id

        $postsFieldArray += $obj
    }

    return $postsFieldArray;
}

function CallExportPostsData 
{
    try 
    {
        $postsEntity = "/XrmActivityStreams/"
        $filterPost = "&" + "$" + "filter=ActionVerb eq 'Post'"
        $postsRequestURL = $baseURL + $postsEntity + $filterTop + $pageSize + $filterCount + $filterPost
        
        $postsData = @()

        write-host -ForegroundColor Green "Downloading data for posts..."

        $data = Invoke-WebRequest -Method Get -uri $postsRequestURL -Headers @{ 'Authorization' = "Bearer $token" } 
        $postsDataContent = $data.Content| ConvertFrom-Json | Select-Object -Property  "value"
        $postsData = $postsDataContent.value

        $totalCount = $data.Content| ConvertFrom-Json | Select -ExpandProperty  "@odata.count"
        if($totalCount -eq 0)
        {
            write-host -ForegroundColor Yellow "No data available for the posts"
            return
        }

        if ($postsData)
        {
            write-progress -Activity "Download progress"-Status ("{0} of $totalCount Posts"-f $postsData.Count) -PercentComplete (($postsData.Count/$totalCount)*100)
        }

        $noOfRequests = ([Math]::Ceiling($totalCount / 100) - 1)
        
        while($noOfRequests -gt 0)
        {
            $skipCount += $pageSize
            $postsRequestURL = $baseURL + $postsEntity + $filterTop + $pageSize + $filterSkip +$skipCount + $filterCount
            $data = Invoke-WebRequest -Method Get -uri $postsRequestURL -Headers @{ 'Authorization' = "Bearer $token" }
            $postsDataContentNext = $data.Content| ConvertFrom-Json | Select-Object -Property  "value"
            $postsData += $postsDataContentNext.value
            
            write-progress -Activity "Download progress"-Status ("{0} of $totalCount Posts"-f $postsData.Count) -PercentComplete (($postsData.Count/$totalCount)*100)

            $noOfRequests--
        }

        write-progress -Activity "Download complete" -Completed

        if ($postsData) {
            write-host -ForegroundColor Green "Exporting data for posts..."

            $parseData = @()
            $parseData = ParsePostsData $postsData
                
            if ($parseData) {
                $filepath = ".\" + $folderName.BaseName + "\" + $hashFilepath['Post'] 
                $parseData | Export-Csv  -Path  $filepath -NoTypeInformation
            }
        }
        else {
            write-host -ForegroundColor Yellow "No data available for the posts"
        }

        write-host -ForegroundColor Green "Exporting data for posts complete"

    } catch {
        write-host -ForegroundColor Green "Exporting data for Posts failed"
        $_.Exception.Message
    }
}

function ParseActivitiesData($dataList)
{
    $activitiesFieldArray = @()

    $activitiesCount = $dataList.Count

    foreach ($dataObj in $dataList) 
    {
        $idx = $dataList.IndexOf($dataObj)
        write-progress -Activity "Processing"-Status ("{0} of $activitiesCount Activities"-f $idx) -PercentComplete (($idx/$activitiesCount)*100)

        $obj = New-Object -TypeName psobject
        
        $obj | Add-Member -MemberType NoteProperty -Name SourceUserName -Value $dataObj.DisplayName
        $obj | Add-Member -MemberType NoteProperty -Name SubType -Value $dataObj.Subtype
        $obj | Add-Member -MemberType NoteProperty -Name ActionVerb -Value $dataObj.ActionVerb
        
        $eventDateTime = $null
        if ($dataObj.EventTime)
        {
            $eventDateTime = ([datetime]::Parse($dataObj.EventTime)).ToString($dateTimeFormat)

            if($eventDateTime -eq $defaultUTCDateTime)
            {
                $eventDateTime =''
            }
        }

        $obj | Add-Member -MemberType NoteProperty -Name EventTime -Value $eventDateTime
        $obj | Add-Member -MemberType NoteProperty -Name ModifiedProperties -Value $dataObj.ModifiedProperties

        $contactLinks = $null
        $dealLinks = $null
        $companyLinks = $null

        $contactLinksObjs = New-Object System.Collections.ArrayList

        $contactLinksObjs = $dataObj.InlineLinks.Relationships | where ItemType -eq "IPM.AbchPerson" | select ItemLinkId
        $dealLinksObjs = $dataObj.InlineLinks.Relationships | where ItemType -eq "IPM.XRMProject.Deal" | select ItemLinkId
        $companyLinksObjs = $dataObj.InlineLinks.Relationships | where ItemType -eq "IPM.Contact.Company" | select ItemLinkId

        if ($contactLinksObjs) {
            $contactLinks = $contactLinksObjs.ItemLinkId -join "|"
        }

        if ($dealLinksObjs) {
            $dealLinks = $dealLinksObjs.ItemLinkId -join "|"
        }

        if ($companyLinksObjs) {
            $companyLinks = $companyLinksObjs.ItemLinkId -join "|"
        }

        $obj | Add-Member -MemberType NoteProperty -Name LinkedContactIds -Value $contactLinks
        $obj | Add-Member -MemberType NoteProperty -Name LinkedDealIds -Value $dealLinks
        $obj | Add-Member -MemberType NoteProperty -Name LinkedCompanyIds -Value $companyLinks

        $activitiesFieldArray += $obj
    }

    return $activitiesFieldArray;
}

function CallExportActivitiesData 
{
    try 
    {
        $activitiesEntity = "/XrmActivityStreams/"
        $filterPost = "&" + "$" + "filter=ActionVerb ne 'Post'"
        $activitiesRequestURL = $baseURL + $activitiesEntity + $filterTop + $pageSize + $filterCount + $filterPost

        $activitiesData = @()

        write-host -ForegroundColor Green "Downloading data for activities..."

        $data = Invoke-WebRequest -Method Get -uri $activitiesRequestURL -Headers @{ 'Authorization' = "Bearer $token" } 
        $activitiesDataContent = $data.Content| ConvertFrom-Json | Select-Object -Property  "value"
        $activitiesData = $activitiesDataContent.value

        $totalCount = $data.Content| ConvertFrom-Json | Select -ExpandProperty  "@odata.count"
        if($totalCount -eq 0)
        {
            write-host -ForegroundColor Yellow "No data available for the activities"
            return
        }

        $noOfRequests = ([Math]::Ceiling($totalCount / 100) - 1)

        if ($activitiesData)
        {
            write-progress -Activity "Download progress"-Status ("{0} of $totalCount Activities"-f $activitiesData.Count) -PercentComplete (($activitiesData.Count/$totalCount)*100)
        }
        
        while($noOfRequests -gt 0)
        {
            $skipCount += $pageSize
            $activitiesRequestURL = $baseURL + $activitiesEntity + $filterTop + $pageSize + $filterSkip +$skipCount + $filterCount
            $data = Invoke-WebRequest -Method Get -uri $activitiesRequestURL -Headers @{ 'Authorization' = "Bearer $token" }
            $activitiesDataContentNext = $data.Content| ConvertFrom-Json | Select-Object -Property  "value"
            $activitiesData += $activitiesDataContentNext.value
            
            write-progress -Activity "Download progress"-Status ("{0} of $totalCount Activities"-f $activitiesData.Count) -PercentComplete (($activitiesData.Count/$totalCount)*100)
            $noOfRequests--
        }

        write-progress -Activity "Download complete" -Completed

        if ($activitiesData) {
            write-host -ForegroundColor Green "Exporting data for activities..."

            $parseData = @()
            $parseData = ParseActivitiesData $activitiesData
                
            if ($parseData) {
                $filepath = ".\" + $folderName.BaseName + "\" + $hashFilepath['Activity'] 
                $parseData | Export-Csv  -Path  $filepath -NoTypeInformation
            }
        }
        else {
            write-host -ForegroundColor Yellow "No data available for the activities"
        }

        write-host -ForegroundColor Green "Exporting data for activities complete"

    } catch {
        write-host -ForegroundColor Green "Exporting data for activities failed"
        $_.Exception.Message
    }
}

function CallPurgeAllData
{
    CallPurgeOCMPersonMetaData
    CallPurgeXrmActivityStreamData
    CallPurgeXrmInsightsData
    CallPurgeXrmDeletedItemsData
    CallPurgeXrmProjectsTemplatesData
    CallPurgeXrmProjectsDealsData
    CallPurgeCompaniesData
    CallPurgeOCMContactsData
}

function CallPurgeOCMContactsData
{
    Write-Host -ForegroundColor Green "Purging shared data for Outlook Customer Manager under contacts folder started..."
    $hashFolder = (FindFolderId "Outlook Customer Manager"  "contacts")

    if($hashFolder)
    {
        Write-Host -ForegroundColor Green "Purging" ($hashFolder["folderItemsCount"] +" items from Outlook Customer Manager folder")
        EmptyFolderByFolderId "Outlook Customer Manager" $hashFolder["folderId"]
        $hashFolder.Clear()
    }
    else
    {
        Write-Host -ForegroundColor Red "Contacts\Outlook Customer Manager folder not found."
    }
}

function CallPurgeCompaniesData
{
    Write-Host -ForegroundColor Green "Purging data for Companies started..."
    $hashFolder = (FindFolderId "Companies" "contacts")

    if($hashFolder)
    {
        $totalItemCount = [int]$hashFolder["folderItemsCount"]
        $companiesFolderId = $hashFolder["folderId"]
        $hashFolder.Clear()

        Write-Host -ForegroundColor Green "Purging" ([string]$totalItemCount +" items from Companies folder")
        EmptyFolderByFolderId "Companies" $companiesFolderId

        $hashFolder = (FindFolderIdByParentFolderId $companiesFolderId "Outlook Customer Manager")

        if($hashFolder)
        {
            $totalItemCount = [int]$hashFolder["folderItemsCount"]
            Write-Host -ForegroundColor Green "Purging" ([string]$totalItemCount +" items from Companies\" + $hashFolder["folderName"] + " folder")
            EmptyFolderByFolderId $hashFolder["folderName"] $hashFolder["folderId"]
            $hashFolder.Clear()
        }
    }
    else
    {
        Write-Host -ForegroundColor Red "Companies folder not found."
    }
}

function CallPurgeOCMPersonMetaData
{
    Write-Host -ForegroundColor Green "Purging data for Outlook Customer Manager under personmetadata folder started..."
    $hashFolder = (FindFolderId "personmetadata"  "msgfolderroot")

    if($hashFolder)
    {
        $personMetaFolderId = $hashFolder["folderId"]
        $hashFolder.Clear()

        $hashFolder = (FindFolderIdByParentFolderId $personMetaFolderId "Outlook Customer Manager")

        if($hashFolder)
        {
            Write-Host -ForegroundColor Green "Purging" ($hashFolder["folderItemsCount"] +" items from "+ $hashFolder["folderName"] +" folder")
            EmptyFolderByFolderId  $hashFolder["folderName"] $hashFolder["folderId"]
            $hashFolder.Clear()
        }
        else
        {
            Write-Host -ForegroundColor Red "Personmetadata\Outlook Customer Manager folder not found."
        }
    }
    else
    {
        Write-Host -ForegroundColor Red "personmetadata folder not found."
    }
}

function CallPurgeXrmActivityStreamData
{
    Write-Host -ForegroundColor Green "Purging data for XrmActivityStream started..."
    $hashFolder = (FindFolderId "XrmActivityStream" "root")

    if($hashFolder)
    {
        Write-Host -ForegroundColor Green "Purging" ($hashFolder["folderItemsCount"] +" items from XrmActivityStream folder")
        $xrmActivityStreamFolderId = $hashFolder["folderId"]
        EmptyFolderByFolderId "XrmActivityStream" $xrmActivityStreamFolderId
        $hashFolder.Clear()

        $hashFolder = (FindFolderIdByParentFolderId $xrmActivityStreamFolderId "Outlook Customer Manager")

        if($hashFolder)
        {
            $totalItemCount = [int]$hashFolder["folderItemsCount"]
            Write-Host -ForegroundColor Green "Purging" ([string]$totalItemCount +" items from XrmActivityStream\" + $hashFolder["folderName"] + " folder")
            EmptyFolderByFolderId $hashFolder["folderName"] $hashFolder["folderId"]
            $hashFolder.Clear()
        }
    }
    else
    {
        Write-Host -ForegroundColor Red "XrmActivityStream folder not found."
    }
}

function CallPurgeXrmInsightsData
{
    Write-Host -ForegroundColor Green "Purging data for XrmInsights started..."
    $hashFolder = (FindFolderId "XrmInsights" "root")

    if($hashFolder)
    {
        Write-Host -ForegroundColor Green "Purging" ($hashFolder["folderItemsCount"] +" items from XrmInsights folder")
        EmptyFolderByFolderId "XrmInsights" $hashFolder["folderId"]
        $hashFolder.Clear()
    }
    else
    {
        Write-Host -ForegroundColor Red "XrmInsights folder not found."
    }
}

function CallPurgeXrmDeletedItemsData
{
    Write-Host -ForegroundColor Green "Purging data for XrmDeletedItems started..."
    $hashFolder = (FindFolderId "XrmDeletedItems" "root")

    if($hashFolder)
    {
        Write-Host -ForegroundColor Green "Purging" ($hashFolder["folderItemsCount"] +" items from XrmDeletedItems folder")
        $xrmDeletedItemsFolderId = $hashFolder["folderId"]
        EmptyFolderByFolderId "XrmDeletedItems" $xrmDeletedItemsFolderId
        $hashFolder.Clear()

        $hashFolder = (FindFolderIdByParentFolderId $xrmDeletedItemsFolderId "Outlook Customer Manager")

        if($hashFolder)
        {
            $totalItemCount = [int]$hashFolder["folderItemsCount"]
            Write-Host -ForegroundColor Green "Purging" ([string]$totalItemCount +" items from XrmDeletedItems\" + $hashFolder["folderName"] + " folder")
            EmptyFolderByFolderId $hashFolder["folderName"] $hashFolder["folderId"]
            $hashFolder.Clear()
        }
    }
    else
    {
        Write-Host -ForegroundColor Red "XrmDeletedItems folder not found."
    }
}

function CallPurgeXrmProjectsTemplatesData ()
{
    Write-Host -ForegroundColor Green "Purging data for XrmProjects Templates started..."
    $hashFolder = FindFolderId "xrmprojects"  "root"

    if($hashFolder)
    {
        $xrmProjectsFolderId = $hashFolder["folderId"]
        $hashFolder.Clear()
        $hashFolder = FindFolderIdByParentFolderId $xrmProjectsFolderId "Templates"

        if($hashFolder)
        {
            Write-Host -ForegroundColor Green "Purging" ($hashFolder["folderItemsCount"] +" items from "+ $hashFolder["folderName"] +" folder")
            EmptyFolderByFolderId  $hashFolder["folderName"] $hashFolder["folderId"]
            $hashFolder.Clear()
        }
        else
        {
            Write-Host -ForegroundColor Red "Templates folder not found."
        }
    }
    else
    {
        Write-Host -ForegroundColor Red "XrmProjects folder not found."
    }
}

function CallPurgeXrmProjectsDealsData ()
{
    Write-Host -ForegroundColor Green "Purging data for Deals started..."

    $hashFolder = FindFolderId "xrmprojects" "root"

    if($hashFolder)
    {
        $xrmProjectsFolderId = $hashFolder.folderId
        $hashFolder.Clear()

        $hashFolder = (FindFolderIdByParentFolderId $xrmProjectsFolderId "Deals")

        if($hashFolder)
        {
            $totalItemCount = [int]$hashFolder["folderItemsCount"]
            $dealsFolderName = $hashFolder["folderName"]
            $dealsFolderId = $hashFolder["folderId"]

            Write-Host -ForegroundColor Green "Purging" ([string]$totalItemCount +" items from "+ $dealsFolderName +" folder")
            EmptyFolderByFolderId $dealsFolderName $dealsFolderId

            $hashFolder.Clear()

            $hashFolder = (FindFolderIdByParentFolderId $dealsFolderId "Outlook Customer Manager")

            if($hashFolder)
            {
                $totalItemCount = [int]$hashFolder["folderItemsCount"]

                Write-Host -ForegroundColor Green "Purging" ([string]$totalItemCount +" items from "+ $dealsFolderName + "\" + $hashFolder["folderName"] + " folder")
                EmptyFolderByFolderId $hashFolder["folderName"] $hashFolder["folderId"]
                $hashFolder.Clear()
            }
        }
        else
        {
            Write-Host -ForegroundColor Red "Deals folder not found."
        }
    }
    else
    {
        Write-Host -ForegroundColor Red "XrmProjects folder not found."
    }
}

function FindFolderId($folderNameToPurge , $folderParentName)
{
    try 
    {
        $fieldURI = $null
        $fieldURIValue = $null

        if($folderNameToPurge -eq "Companies") {
            $fieldURI = "folder:FolderClass"
            $fieldURIValue = "IPF.Contact.Company"
        }
        else {
            $fieldURI = "folder:DisplayName"
            $fieldURIValue = $folderNameToPurge
        }

        $findFolder = "<?xml version='1.0' encoding='utf-8'?>
        <soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:m='http://schemas.microsoft.com/exchange/services/2006/messages' xmlns:t='http://schemas.microsoft.com/exchange/services/2006/types' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>
            <soap:Header>
            <t:RequestServerVersion Version='V2017_07_11' />
            </soap:Header>
            <soap:Body>
            <m:FindFolder Traversal='Shallow'>
                <m:FolderShape>
                <t:BaseShape>Default</t:BaseShape>
                </m:FolderShape>
                <m:Restriction>
                <t:IsEqualTo>
                    <t:FieldURI FieldURI='$fieldURI' />
                    <t:FieldURIOrConstant>
                    <t:Constant Value='$fieldURIValue' />
                    </t:FieldURIOrConstant>
                </t:IsEqualTo>
                </m:Restriction>
                <m:ParentFolderIds>
                <t:DistinguishedFolderId Id='$folderParentName'>
                    <t:Mailbox>
                      <t:EmailAddress>$smtpAddress</t:EmailAddress>
                    </t:Mailbox>
                </t:DistinguishedFolderId>
                </m:ParentFolderIds>
            </m:FindFolder>
            </soap:Body>
        </soap:Envelope>"

        $response = Invoke-WebRequest $uri -Method Post -ContentType "text/xml" -Credential $cred -Body "$findFolder"
        $responseXML = [xml]$response
        $responseCode = (select-xml -Xml $responseXML -Namespace $ns -XPath "//m:ResponseCode").Node.InnerText

        $folderId = $null
        $folderItemsCount = $null

        if ($responseCode -eq 'NoError')
        {
            $folderId = (select-xml -Xml $responseXML -Namespace $ns -XPath "//t:FolderId").Node
            
            $folderItemsCount = (select-xml -Xml $responseXML -Namespace $ns -XPath "//t:TotalCount").Node.InnerText

            if($folderId -and $folderItemsCount)
            {
                $hashFolder["folderId"] = $folderId.Id
                $hashFolder["folderItemsCount"] = $folderItemsCount
            }
            else
            {
                $hashFolder = $null
            }
        }
        else
        {
            $errorMessageText = (select-xml -Xml $responseXML -Namespace $ns -XPath "//m:MessageText").Node.InnerText
            Write-Host -ForegroundColor Red "Purging data for" $folderNameToPurge "failed with error - $errorMessageText"
            $hashFolder = $null
        }
    }
    catch
    {
        Write-Host -ForegroundColor Red "Failed to get the folder id for '$folderNameToPurge'"
        $_.Exception.Message
        $hashFolder = $null
    }

    return $hashFolder
}

function EmptyFolderByFolderId($folderNameToPurge, $folderId)
{
    try
    {
        if($folderId)
        {
            $response = $null
            
            $isSubFolderDelete = $hashSubFolderDelete[$folderNameToPurge]
            #
            # EmptyFolder request
            #
            $EmptyFolderDefault = "<?xml version='1.0' encoding='utf-8'?>
            <soap:Envelope
                xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'
                xmlns:m='http://schemas.microsoft.com/exchange/services/2006/messages'
                xmlns:t='http://schemas.microsoft.com/exchange/services/2006/types'
                xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>
                <soap:Header>
                    <t:RequestServerVersion Version='V2017_04_14' />
                </soap:Header>
                <soap:Body>
                <m:EmptyFolder DeleteType='HardDelete' DeleteSubFolders='$isSubFolderDelete'>
                    <m:FolderIds>
                        <t:FolderId Id='$folderId' />
                    </m:FolderIds>
                </m:EmptyFolder>
                </soap:Body>
            </soap:Envelope>"

            $response = Invoke-WebRequest $uri -Method Post -ContentType "text/xml" -Credential $cred -Body "$EmptyFolderDefault"
            $responseXML = [xml]$response
            $responseCode = (select-xml -Xml $responseXML -Namespace $ns -XPath "//m:ResponseCode").Node.InnerText

            if ($responseCode -eq 'NoError') 
            {
               Write-Host -ForegroundColor Green "Purging data for" $folderNameToPurge "complete."
               $countRetry = 0
            }
            else
            {
                $errorMessageText = (select-xml -Xml $responseXML -Namespace $ns -XPath "//m:MessageText").Node.InnerText
                Write-Host -ForegroundColor Red "Purging data for" $folderNameToPurge "failed with error - $errorMessageText"

                $countRetry++

                CallStartXrmSessionV2

                if($countRetry -gt $maxRetryCount)
                {
                    $countRetry = 0
                    Write-Host -ForegroundColor Red "Purging data for" $folderNameToPurge "failed. Please try again later."
                }
                else
                {
                    Write-Host -ForegroundColor Yellow "Retry purging data for" $folderNameToPurge. "Retry attempt - " $countRetry
                    EmptyFolderByFolderId $folderNameToPurge $folderId
                }
            }
        }
    }
    catch
    {
        if($_.Exception.Response.StatusCode.Value__ -eq '502' -or $_.Exception.Response.StatusCode.Value__ -eq '503' -or $_.Exception.Response.StatusCode.Value__ -eq '500')
        {
            Write-Host -ForegroundColor Red "The exchange server had returned a timeout error while purging data for $folderNameToPurge."
            $countRetry++

            if($countRetry -gt $maxRetryCount)
            {
                $countRetry = 0
                Write-Host -ForegroundColor Red "Purging data for" $folderNameToPurge "failed. Please try again later."
            }
            else
            {
                Write-Host -ForegroundColor Yellow "Retry purging data for" $folderNameToPurge. "Retry attempt - " $countRetry
                EmptyFolderByFolderId $folderNameToPurge $folderId
            }
        }
        else
        {
            Write-Host -ForegroundColor Red "Purging data for" $folderNameToPurge "failed."
            $_.Exception.Message
        }
    }
}

function FindFolderIdByParentFolderId($folderId, $childFolderName)
{
    try
    {
        $findFolderIdByParentFolderId ="<?xml version='1.0' encoding='utf-8'?>
        <soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:m='http://schemas.microsoft.com/exchange/services/2006/messages' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/' xmlns:t='http://schemas.microsoft.com/exchange/services/2006/types'>
            <soap:Header>
            <RequestServerVersion Version='V2017_04_14' xmlns='http://schemas.microsoft.com/exchange/services/2006/types' />
            </soap:Header>
            <soap:Body>
            <FindFolder Traversal='Shallow' xmlns='http://schemas.microsoft.com/exchange/services/2006/messages'>
                <FolderShape>
                <t:BaseShape>Default</t:BaseShape>
                </FolderShape>
                <m:Restriction>
                <t:IsEqualTo>
                <t:FieldURI FieldURI='folder:DisplayName' />
                <t:FieldURIOrConstant>
                    <t:Constant Value='$childFolderName' />
                </t:FieldURIOrConstant>
                </t:IsEqualTo>
                </m:Restriction>
                <ParentFolderIds>
                <t:FolderId Id='$folderId' />
                </ParentFolderIds>
            </FindFolder>
            </soap:Body>
        </soap:Envelope>"

        $response = Invoke-WebRequest $uri -Method Post -ContentType "text/xml" -Credential $cred -Body "$findFolderIdByParentFolderId"
        $responseXML = [xml]$response
        $responseCode = (select-xml -Xml $responseXML -Namespace $ns -XPath "//m:ResponseCode").Node.InnerText

        $folderId = $null
        $folderName = $null
        $folderItemsCount = $null

        if ($responseCode -eq 'NoError')
        {
            $folderId = (select-xml -Xml $responseXML -Namespace $ns -XPath "//t:FolderId").Node
            $folderName = (select-xml -Xml $responseXML -Namespace $ns -XPath "//t:DisplayName").Node.InnerText
            $folderItemsCount = (select-xml -Xml $responseXML -Namespace $ns -XPath "//t:TotalCount").Node.InnerText

            if($folderId -and $folderName -and $folderItemsCount)
            {
                $hashFolder["folderName"] = $folderName
                $hashFolder["folderId"] = $folderId.Id
                $hashFolder["folderItemsCount"] = $folderItemsCount
            }
            else
            {
                $hashFolder = $null
            }
        }
        else
        {
            $errorMessageText = (select-xml -Xml $responseXML -Namespace $ns -XPath "//m:MessageText").Node.InnerText
            $hashFolder = $null
        }
    }
    catch
    {
        Write-Host -ForegroundColor Red "Failed to get the folder id for" $childFolderName"."
        $_.Exception.Message
        $hashFolder = $null
    }

     return $hashFolder
}

#
# Constants
#

$skipCount = 0
$pageSize = 100
$baseURL = "https://outlook.office365.com/api/beta/Me"
$entityTypeOrg = "/XrmOrganizations/"
$entityTypeDeal = "/XrmDeals/"
$filterTop = "?%24top="
$filterSkip = "&%24skip="
$filterCount = "&%24count=true"
$filterOrderByOrg = "&%24OrderBy=LastModifiedDateTime%20asc"
$filterOrderByDeal = "&%24OrderBy=LastModifiedTime%20asc"
$dateTimeFormat = "yyyy-MM-ddTHH:mm:ss.ffZ"
$defaultUTCDateTime = "1601-01-01T00:00:00Z"
$defaultDateTime = @("0001-01-01", "1601-01-01")
$countRetry = 0
$maxRetryCount = 3

$hashFilepath = @{ 
    Contact = 'Contacts.csv'
    ContactTemplate = 'ContactSchema.csv'
    Company = 'Companies.csv'
    CompanyTemplate = 'CompanySchema.csv'
    Deal = 'Deals.csv'
    DealTemplate = 'DealSchema.csv'
    Task = 'Tasks.csv'
    Post = 'Posts.csv'
    Activity = 'Activities.csv'
}

$hashSubFolderDelete = @{ 
    "Outlook Customer Manager" = "false"
    "Companies" = "false"
    "XrmActivityStream" = "false"
    "XrmInsights" = "false"
    "XrmDeletedItems" = "false"
    "Deals" = "false"
    "Templates" = "false"
}

$customFiedsTypeMapping = @{
    Boolean = 'Boolean'
    Choice= 'String'
    Currency = 'Double'
    Numeric = 'Double'
    Date = 'SystemTime'
    Link = 'String'
    Text = 'String'
}

$hashPriority = @{
    '1' = 'High'
    '0' = 'Normal'
   '-1' = 'Low'
}
        
$hashContactPropertyIdToText = @{
     
  '0x1000' = 'Notes'
  '0x3a06' = 'GivenName'
  '0x3a11' = 'LastName'
  '0x3a16' = 'CompanyName'
  '0x3a44' = 'MiddleName'
  '0x3a42' = 'BirthDay'
  '0x3a17' = 'JobTitle'
  '0x3a08' = 'BusinessTelephoneNumber'
  '0x3a09' = 'HomeTelephoneNumber'
  '0x3a1c' = 'MobileTelephoneNumber'
  '0x3a1f' = 'OtherTelephoneNumber'
 
  '32899' = 'Email1Address' 
  '32896' = 'Email1DisplayName' 
  '32900' = 'Email1OriginalDisplayName' 
  '32915' = 'Email2Address' 
  '32912' = 'Email2DisplayName' 
  '32916' = 'Email2OriginalDisplayName' 
  '32931' = 'Email3Address' 
  '32928' = 'Email3DisplayName' 
  '32932' = 'Email3OriginalDisplayName' 
  
  '0x3a5d' = 'HomeAddressStreet' 
  '0x3a59' = 'HomeAddressCity' 
  '0x3a5c' = 'HomeAddressStateOrProvince' 
  '0x3a5b' = 'HomeAddressPostalCode' 
  '0x3a5a' = 'HomeAddressCountry' 
 
  '32837' = 'WorkAddressStreet' 
  '32838' = 'WorkAddressCity' 
  '32839' = 'WorkAddressStateOrProvince' 
  '32840' = 'WorkAddressPostalCode' 
  '32841' = 'WorkAddressCountry' 
 
  '0x3a63' = 'OtherAddressStreet' 
  '0x3a5f' = 'OtherAddressCity' 
  '0x3a62' = 'OtherAddressStateOrProvince' 
  '0x3a61' = 'OtherAddressPostalCode' 
  '0x3a60' = 'OtherAddressCountry' 
  'CustomerBit' = 'IsBusinessContact'
  'ItemLinkId' = 'XrmId'
  'XrmSourceMailboxGuid' = 'Shared'
  'CompanyLinks' = 'CompanyLinks'
  'DealLinks' = 'DealLinks'
}

$hashStageIdToLabel = @{}
$hashFolder = @{}

# get user credentials ready
$cred = getUserCreds $user

if ($operation.ToLowerInvariant().Contains('purge-') -and !$smtpAddress)
{
    $smtpAddress = $cred.UserName
}
else
{
    $smtpAddress = $smtpAddress.ToLowerInvariant()
}

if ($operation.ToLowerInvariant().Contains('purge-')) 
{
    Read-Host "Ready to start purging data for '$smtpAddress'. Press any key to continue"
}

$folderName = GetFolderName

$uri = "https://outlook.office365.com/EWS/Exchange.asmx"
$ns = @{
    xsi = 'http://www.w3.org/2001/XMLSchema-instance';
    t = "http://schemas.microsoft.com/exchange/services/2006/types"; 
    h = "http://schemas.microsoft.com/exchange/services/2006/types"
    m = 'http://schemas.microsoft.com/exchange/services/2006/messages';
    soap = 'http://schemas.xmlsoap.org/soap/envelope/';
    s = 'http://schemas.xmlsoap.org/soap/envelope/';
}

#
# Use a SOAP call to get an access token
#
$requestXML = "<?xml version='1.0' encoding='utf-8'?>
<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:m='http://schemas.microsoft.com/exchange/services/2006/messages' xmlns:t='http://schemas.microsoft.com/exchange/services/2006/types' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>
  <soap:Header>
    <t:RequestServerVersion Version='V2017_04_14' />
  </soap:Header>
  <soap:Body>
    <m:GetClientAccessToken>
      <m:TokenRequests>
        <t:TokenRequest>
          <t:Id>8f4d1315-5cf9-4621-9872-b1b94618e70a</t:Id>
          <t:TokenType>ExtensionRestApiCallback</t:TokenType>
        </t:TokenRequest>
      </m:TokenRequests>
    </m:GetClientAccessToken>
  </soap:Body>
</soap:Envelope>"

#
# SyncfolderItems Request
#
 $syncFolderItemsStart = "<?xml version='1.0' encoding='utf-8'?>
    <soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:m='http://schemas.microsoft.com/exchange/services/2006/messages' xmlns:t='http://schemas.microsoft.com/exchange/services/2006/types' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>
      <soap:Header>
        <t:RequestServerVersion Version='V2017_04_14' />
      </soap:Header>
      <soap:Body>
        <m:SyncFolderItems>
          <m:ItemShape>
            <t:BaseShape>Default</t:BaseShape>
          </m:ItemShape>
          <m:SyncFolderId>
            <t:DistinguishedFolderId Id='tasks'/>
          </m:SyncFolderId>"
          
$syncFolderItemsEnd = "<m:MaxChangesReturned>100</m:MaxChangesReturned>
    </m:SyncFolderItems>
    </soap:Body>
</soap:Envelope>"

try {
    $response = Invoke-WebRequest $uri -Method Post -ContentType "text/xml" -Credential $cred -Body "$requestXML"
    $responseXML = [xml]$response

    $responseCode = (select-xml -Xml $responseXML -Namespace $ns -XPath "//m:ResponseCode").Node.InnerText
    if ($responseCode -eq 'NoError') {
        $token = (select-xml -Xml $responseXML -Namespace $ns -XPath "//t:TokenValue").Node.InnerText
    }
    else {
        write-host -ForegroundColor Red "Token request failed with error '$responseCode'"
    }
}catch {
    write-host -ForegroundColor Red "Token request call failed with error -"
    $_.Exception.Message
    break
}

switch ($operation.ToLowerInvariant())
{
    "export-all-data" { CallExportAllData; break }
    "export-contact-templates" { CallExportContactTemplatesData; break}
    "export-company-templates" { CallExportCompanyTemplatesData; break}
    "export-deal-templates" { CallExportDealTemplatesData; break}
    "export-contacts-data" {
        CallExportContactTemplatesData;
        CallExportContactsData; 
        break 
    }
    "export-companies-data" {
        CallExportCompanyTemplatesData;
        CallExportCompaniesData; 
        break 
    }
    "export-tasks-data" { CallExportTasksData; break }
    "export-deals-data" { 
        CallExportDealTemplatesData;
        CallExportDealsData; 
        break 
    }
    "export-posts-data" { CallExportPostsData; break }
    "export-activities-data" { CallExportActivitiesData; break }
    "purge-all-data" { CallPurgeAllData; break }
    "purge-contacts-data" { CallPurgeOCMContactsData; break }
    "purge-companies-data" { CallPurgeCompaniesData; break }
    "purge-ocm-personmeta-data" { CallPurgeOCMPersonMetaData; break }
    "purge-xrm-activity-stream-data" { CallPurgeXrmActivityStreamData; break }
    "purge-xrm-insights-data" { CallPurgeXrmInsightsData; break }
    "purge-xrm-deleted-items-data" { CallPurgeXrmDeletedItemsData; break }
    "purge-templates-data" { CallPurgeXrmProjectsTemplatesData ; break }
    "purge-deals-data" { CallPurgeXrmProjectsDealsData ; break }
  
    default {
        Write-Host -ForegroundColor Red "Unexpected operation: $operation"
    }
}