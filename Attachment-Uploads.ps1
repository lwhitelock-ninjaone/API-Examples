# API Details
# This script contains different examples for how to upload attachments to different locations via the NinjaOne Public API.

$ClientID = 'ClientID'
$Secret = 'Secret'
$Script:NinjaOneInstance = 'eu.ninjarmm.com'

# Determines the Mime Type to use based on file extension
function Get-MimeType {
    param([string]$Path)

    $ext = [IO.Path]::GetExtension($Path).ToLowerInvariant()

    switch ($ext) {
        ".jpg" { "image/jpeg" }
        ".jpeg" { "image/jpeg" }
        ".png" { "image/png" }
        ".gif" { "image/gif" }
        ".cab" { "application/vnd.ms-cab-compressed" }
        ".txt" { "text/plain" }
        ".log" { "text/plain" }
        ".pdf" { "application/pdf" }
        ".csv" { "text/csv" }
        ".mp3" { "audio/mpeg" }
        ".eml" { "message/rfc822" }
        ".dot" { "application/msword" }
        ".wbk" { "application/msword" }
        ".doc" { "application/msword" }
        ".docx" { "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }
        ".rtf" { "application/rtf" }
        ".xls" { "application/vnd.ms-excel" }
        ".xlsx" { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }
        ".ods" { "application/vnd.oasis.opendocument.spreadsheet" }
        ".ppt" { "application/vnd.ms-powerpoint" }
        ".pptx" { "application/vnd.openxmlformats-officedocument.presentationml.presentation" }
        ".pps" { "application/vnd.ms-powerpoint" }
        ".ppsx" { "application/vnd.openxmlformats-officedocument.presentationml.slideshow" }
        ".sldx" { "application/vnd.openxmlformats-officedocument.presentationml.slide" }
        ".vsd" { "application/vnd.visio" }
        ".vsdx" { "application/vnd.ms-visio.drawing.main+xml" }
        ".xml" { "application/xml" }
        ".html" { "text/html" }
        ".zip" { "application/zip" }
        ".rar" { "application/vnd.rar" }
        ".tar" { "application/x-tar" }
        default { "application/octet-stream" }
    }
}

function Invoke-UploadNinjaOneKBArticle {
    param (
        $FileName,
        $FilePath,
        $FolderPath,
        $OrganizationID,
        $FailCount = 0
    )

    try {
        $multipartContent = [System.Net.Http.MultipartFormDataContent]::new()
        # Only add fields that were provided
        if ($null -ne $OrganizationID) {
            $multipartContent.Add([System.Net.Http.StringContent]::new($OrganizationId), 'organizationId')
        }
        if ($FolderPath -ne '') {
            $multipartContent.Add([System.Net.Http.StringContent]::new($FolderPath), 'folderPath')
        }

        $MimeType = Get-MimeType $FilePath

        $FileStream = [System.IO.FileStream]::new($FilePath, [System.IO.FileMode]::Open)
        $fileHeader = [System.Net.Http.Headers.ContentDispositionHeaderValue]::new("form-data")
        $fileHeader.Name = 'files'
        $fileHeader.FileName = $FileName
        $fileContent = [System.Net.Http.StreamContent]::new($FileStream)
        $fileContent.Headers.ContentDisposition = $fileHeader
        $fileContent.Headers.ContentType = [System.Net.Http.Headers.MediaTypeHeaderValue]::Parse($MimeType)
        $multipartContent.Add($fileContent)

        $URI = "https://$($Script:NinjaOneInstance)/ws/api/v2/knowledgebase/articles/upload"
        $Result = (Invoke-WebRequest -Uri $URI -Body $multipartContent -Method 'POST' -Headers $Script:AuthHeader ).content | ConvertFrom-Json -Depth 100
        $FileStream.close()
        return $Result
    } catch {
        $FileStream.close()
        Write-Error "Failed to upload file: $_"
    }     
}



function Invoke-UploadNinjaOneFile($FileName, $FilePath, $ContentType, $EntityType) {

    try {
        $multipartContent = [System.Net.Http.MultipartFormDataContent]::new()
        $FileStream = [System.IO.FileStream]::new($FilePath, [System.IO.FileMode]::Open)
        $fileHeader = [System.Net.Http.Headers.ContentDispositionHeaderValue]::new("form-data")
        $fileHeader.Name = 'files'
        $fileHeader.FileName = $FileName
        $fileContent = [System.Net.Http.StreamContent]::new($FileStream)
        $fileContent.Headers.ContentDisposition = $fileHeader
        $fileContent.Headers.ContentType = [System.Net.Http.Headers.MediaTypeHeaderValue]::Parse($ContentType)
        $multipartContent.Add($fileContent)
        if ($EntityType) {
            $URI = "https://$($Script:NinjaOneInstance)/ws/api/v2/attachments/temp/upload?entityType=$EntityType"
        } else {
            $URI = "https://$($Script:NinjaOneInstance)/ws/api/v2/attachments/temp/upload"
        }
        $Result = (Invoke-WebRequest -Uri $URI -Body $multipartContent -Method 'POST' -Headers $Script:AuthHeader ).content | ConvertFrom-Json -Depth 100
        $FileStream.close()
        return $Result
    } catch {
        $FileStream.close()
        Throw "Failed to upload file: $_"
    }
}

function Invoke-UploadNinjaOneRelatedFile($FileName, $FilePath, $ContentType, $EntityType, $EntityID) {

    try {
        $multipartContent = [System.Net.Http.MultipartFormDataContent]::new()
        $FileStream = [System.IO.FileStream]::new($FilePath, [System.IO.FileMode]::Open)
        $fileHeader = [System.Net.Http.Headers.ContentDispositionHeaderValue]::new("form-data")
        $fileHeader.Name = 'file'
        $fileHeader.FileName = $FileName
        $fileContent = [System.Net.Http.StreamContent]::new($FileStream)
        $fileContent.Headers.ContentDisposition = $fileHeader
        $fileContent.Headers.ContentType = [System.Net.Http.Headers.MediaTypeHeaderValue]::Parse($ContentType)
        $multipartContent.Add($fileContent)
        $URI = "https://$($Script:NinjaOneInstance)/ws/api/v2/related-items/entity/$($EntityType)/$($EntityID)/attachment"
        write-host $URI
        $Result = (Invoke-WebRequest -Uri $URI -Body $multipartContent -Method 'POST' -Headers $Script:AuthHeader ).content | ConvertFrom-Json -Depth 100
        $FileStream.close()
        return $Result
    } catch {
        $FileStream.close()
        Throw "Failed to upload file: $_"
    }
}

# Supported extensions for attachment based fields
$NinjaOneSupportedUploadTypes = @(
    "jpg", "jpeg", "png", "gif", "cab", "txt", "log", "pdf",
    "csv", "mp3", "eml", "dot", "wbk", "doc", "docx", "rtf",
    "xls", "xlsx", "ods", "ppt", "pptx", "pps", "ppsx", "sldx",
    "vsd", "vsdx", "xml", "html", "zip", "rar", "tar"
)

# Supported extensions for direct upload to the Knowledge Base
$NinjaOneSupportedKBTypes = @(
    "dot", "wbk", "doc", "docx", "rtf", "xls", "xlsx", "ods",
    "ppt", "pptx", "pps", "ppsx", "sldx", "vsd", "vsdx", "pdf"
)

# Supported extensions for in line images
$NinjaOneSupportedImageTypes = @(
    "jpg", "jpeg", "png", "gif"
)

# Connect to NinjaOne and Obtain an Authorisation Token

$AuthBody = @{
    'grant_type'    = 'client_credentials'
    'client_id'     = $ClientID
    'client_secret' = $Secret
    'scope'         = 'monitoring management' 
}


$Result = Invoke-WebRequest -uri "https://$($Script:NinjaOneInstance)/ws/oauth/token" -Method POST -Body $AuthBody -ContentType 'application/x-www-form-urlencoded'

$Script:AuthHeader = @{
    'Authorization' = "Bearer $(($Result.content | ConvertFrom-Json).access_token)"
}

# Define the file to upload
$UploadFile = 'C:\Temp\ExampleFileToUpload.docx'
$UploadFileName = 'ExampleFileToUpload.docx'
$FileExtension = ([System.IO.Path]::GetExtension($UploadFile)).TrimStart('.')
$MimeType = Get-MimeType $UploadFile

######## Uploading as a KB Article ########

#### Uploading to the Knowledge Base

if ($FileExtension -in $NinjaOneSupportedKBTypes) {

    # Provide the path inside the Knowledge Base with each folder seperated with a |
    # For Global KB do not provide the OrganizationID parameter for an Organization KB provide the parameter with the ID of the organization 
    $Result = Invoke-UploadNinjaOneKBArticle -FileName $UploadFileName -FilePath $UploadFile -FolderPath 'KB|Example|Path'

} else {
    Write-Error "$FileExtension is not supported in the NinjaOne Knowledge Base"
}

######## Uploading to attachment custom fields ########

#### Custom Field Settings
$CustomFieldName = 'attachmentExample'

#### Uploading to a device attachment custom field
$DeviceID = 21
$EntityType = 'NODE'

if ($FileExtension -in $NinjaOneSupportedUploadTypes) {

    # Upload the file to the temporary endpoint using the NODE entity type for devices
    $UploadedDeviceFile = Invoke-UploadNinjaOneFile -FileName $UploadFileName -FilePath $UploadFile -ContentType $MimeType -EntityType $EntityType

    # Provide the metadata details returned from the upload as the value of the attachment custom field in an update.
    $DeviceUpdate = @{
        "$CustomFieldName" = $UploadedDeviceFile
    } | ConvertTo-Json

    # Update the devices custom field
    $Null = (Invoke-WebRequest -uri "https://$($Script:NinjaOneInstance)/v2/device/$($DeviceID)/custom-fields" -Method PATCH -Headers $Script:AuthHeader -Body $DeviceUpdate -ContentType 'application/json').Content | ConvertFrom-Json

} else {
    Write-Error "$FileExtension is not supported in the NinjaOne Knowledge Base"
}


#### Uploading to a end user attachment custom field
$EndUserID = 8
$EntityType = 'END_USER'

if ($FileExtension -in $NinjaOneSupportedUploadTypes) {

    # Upload the file to the temporary endpoint using the END_USER entity type for End Users
    $UploadedUserFile = Invoke-UploadNinjaOneFile -FileName $UploadFileName -FilePath $UploadFile -ContentType $MimeType -EntityType $EntityType

    # Provide the metadata details returned from the upload as the value of the attachment custom field in an update.
    $UserUpdate = @{
        "$CustomFieldName" = $UploadedUserFile
    } | ConvertTo-Json

    # Update the end users custom field
    $Null = (Invoke-WebRequest -uri "https://$($Script:NinjaOneInstance)/v2/user/end-user/$($EndUserID)/custom-fields" -Method PATCH -Headers $Script:AuthHeader -Body $UserUpdate -ContentType 'application/json').Content | ConvertFrom-Json

} else {
    Write-Error "$FileExtension is not supported in the NinjaOne Knowledge Base"
}

#### Uploading to a organization attachment custom field
$OrganizationID = 3
$EntityType = 'ORGANIZATION'

if ($FileExtension -in $NinjaOneSupportedUploadTypes) {

    # Upload the file to the temporary endpoint using the ORGANIZATION entity type for organizations
    $UploadedOrganizationFile = Invoke-UploadNinjaOneFile -FileName $UploadFileName -FilePath $UploadFile -ContentType $MimeType -EntityType $EntityType

    # Provide the metadata details returned from the upload as the value of the attachment custom field in an update.
    $OrgUpdate = @{
        "$CustomFieldName" = $UploadedOrganizationFile
    } | ConvertTo-Json

    # Update the organization custom field
    $Null = (Invoke-WebRequest -uri "https://$($Script:NinjaOneInstance)/v2/organization/$($OrganizationID)/custom-fields" -Method PATCH -Headers $Script:AuthHeader -Body $OrgUpdate -ContentType 'application/json').Content | ConvertFrom-Json

} else {
    Write-Error "$FileExtension is not supported in the NinjaOne Knowledge Base"
}

#### Uploading to a location attachment custom field
$OrganizationID = 3
$LocationID = 3
$EntityType = 'LOCATION'

if ($FileExtension -in $NinjaOneSupportedUploadTypes) {

    # Upload the file to the temporary endpoint using the LOCATION entity type for locations
    $UploadedLocationFile = Invoke-UploadNinjaOneFile -FileName $UploadFileName -FilePath $UploadFile -ContentType $MimeType -EntityType $EntityType

    # Provide the metadata details returned from the upload as the value of the attachment custom field in an update.
    $LocUpdate = @{
        "$CustomFieldName" = $UploadedLocationFile
    } | ConvertTo-Json

    # Update the location custom field
    $Null = (Invoke-WebRequest -uri "https://$($Script:NinjaOneInstance)/v2/organization/$($OrganizationID)/location/$($LocationID)/custom-fields" -Method PATCH -Headers $Script:AuthHeader -Body $LocUpdate -ContentType 'application/json').Content | ConvertFrom-Json

} else {
    Write-Error "$FileExtension is not supported in the NinjaOne Knowledge Base"
}

######## Related Item Uploads ########

#### Uploading as a related item attachment
$OrganizationID = 3
$EntityType = 'ORGANIZATION'

if ($FileExtension -in $NinjaOneSupportedUploadTypes) {
    Invoke-UploadNinjaOneRelatedFile -FileName $UploadFileName -FilePath $UploadFile -ContentType $MimeType -EntityType $EntityType -EntityID $OrganizationID
} else {
    Write-Error "$FileExtension is not supported in the NinjaOne Knowledge Base"
}

######## Uploading images to use inline in HTML ########

########### Image File Settings

$ImageUploadFile = 'C:\Temp\ExampleImage.png'
$ImageUploadFileName = 'ExampleImage.png'
$ImageFileExtension = ([System.IO.Path]::GetExtension($ImageUploadFile)).TrimStart('.')
$ImageMimeType = Get-MimeType $ImageUploadFile

#### Uploading as an inline image in a KB Article
$EntityType = $NULL

if ($ImageFileExtension -in $NinjaOneSupportedImageTypes) {

    # Upload the file to the temporary endpoint using no entity type for KB Articles
    $UploadedImageFile = Invoke-UploadNinjaOneFile -FileName $ImageUploadFileName -FilePath $ImageUploadFile -ContentType $ImageMimeType

    # Provide the body for creating a KB Article inserting the content ID as the CID for the image where you would like it to appear
    $KBCreate = @(@{
            name    = 'KB Article with Image'
            content = @{
                html = "<h1>This is some html with an image in it</h1> <img src=" + '"cid:' + $($UploadedImageFile.contentId) + '"></img>'
            }
        }) | ConvertTo-Json -AsArray

    # Update the location custom field
    $Null = (Invoke-WebRequest -uri "https://$($Script:NinjaOneInstance)/v2/knowledgebase/articles" -Method POST -Headers $Script:AuthHeader -Body $KBCreate -ContentType 'application/json').Content | ConvertFrom-Json

} else {
    Write-Error "$ImageFileExtension is not supported for inline images"
}


#### Uploading as an inline image in a WYSIWYG Custom Field, this example is for an Organization field, but will work for others if you set the correct Entity Type
$WYSIWYGFieldName = 'lwghtml2'
$OrganizationID = 38
$EntityType = 'ORGANIZATION'

if ($ImageFileExtension -in $NinjaOneSupportedImageTypes) {

    # Upload the file to the temporary endpoint using the ORGANIZATION entity type for organizations
    $UploadedOrganizationImageFile = Invoke-UploadNinjaOneFile -FileName $ImageUploadFileName -FilePath $ImageUploadFile -ContentType $ImageMimeType -EntityType $EntityType

    # Provide the metadata details returned from the upload as the value of the attachment custom field in an update.
    $OrgUpdate = @{
        "$WYSIWYGFieldName" = @{
            html = "<h1>This is some html with an image in it</h1> <img src=" + '"cid:' + $($UploadedOrganizationImageFile.contentId) + '"></img>'
        }
    } | ConvertTo-Json

    # Update the organization custom field
    $Null = (Invoke-WebRequest -uri "https://$($Script:NinjaOneInstance)/v2/organization/$($OrganizationID)/custom-fields" -Method PATCH -Headers $Script:AuthHeader -Body $OrgUpdate -ContentType 'application/json').Content | ConvertFrom-Json

} else {
    Write-Error "$FileExtension is not supported in the NinjaOne Knowledge Base"
}

#### Uploading to an Apps and Services Field
$WYSIWYGFieldName = 'wysiwygField'
$OrganizationID = 38
$AppsAndServiceDocID = 3
$EntityType = 'DOCUMENT'

if ($ImageFileExtension -in $NinjaOneSupportedImageTypes) {

    # Upload the file to the temporary endpoint using the ORGANIZATION entity type for organizations
    $UploadedAppsServImageFile = Invoke-UploadNinjaOneFile -FileName $ImageUploadFileName -FilePath $ImageUploadFile -ContentType $ImageMimeType -EntityType $EntityType

    # Provide the metadata details returned from the upload as the value of the attachment custom field in an update.
    $OrgUpdate = @{
        documentName = 'Test'
        fields       = @{
            "$WYSIWYGFieldName" = @{
                html = "<h1>This is some html with an image in it</h1> <img src=" + '"cid:' + $($UploadedAppsServImageFile.contentId) + '"></img>'
            }
        }
    } | ConvertTo-Json

    # Update the organization custom field
    $Null = (Invoke-WebRequest -uri "https://$($Script:NinjaOneInstance)/v2/organization/$($OrganizationID)/document/$($AppsAndServiceDocID)" -Method POST -Headers $Script:AuthHeader -Body $OrgUpdate -ContentType 'application/json').Content | ConvertFrom-Json

} else {
    Write-Error "$FileExtension is not supported in the NinjaOne Knowledge Base"
}





