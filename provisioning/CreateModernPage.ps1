#
#This is the script to create a new modern page
#
Param(
	[Parameter(Mandatory = $true)]
	[string]$TenantUrl,
	
	[Parameter(Mandatory = $false, Position = 2)]
    [PSCredential]$Credentials,
	
	[Parameter(Mandatory = $false, Position = 3)]
    [string]$SkipPowerShellInstall = $false,
	
	[Parameter(Mandatory = $true, Position = 4)]
	[string]$SiteUrl
)

# Check if PnP PowerShell is installed
if (!$SkipPowershellInstall) {
    $modules = Get-Module -Name SharePointPnPPowerShellOnline -ListAvailable
    if ($modules -eq $null) {
        # Not installed.
        Install-Module -Name SharePointPnPPowerShellOnline -Scope CurrentUser -Force
        Import-Module -Name SharePointPnPPowerShellOnline -DisableNameChecking
    }
}
if ($Credentials -eq $null) {
    $Credentials = Get-Credential -Message "Enter credentials to connect to $TenantUrl"
}
if($SiteUrl)
{	
    Connect-PnPOnline -Url $SiteUrl -Credentials $Credentials
    Write-Host "Connection not created" -ForegroundColor Red
    
	# $ArticlePageName = Read-Host 'What is name of the Page for Article Page Layout'	
	# $IsPage = Get-PnPClientSidePage -Identity $ArticlePageName -ErrorAction SilentlyContinue
	# if($IsPage -eq $null){
		# Write-Host "Page does not exist or is not modern page" -ForegroundColor Green
		# Add-PnPClientSidePage -Name $ArticlePageName
	# }
	# else{
		# Write-Host "Page: $($ArticlePageName) is present" -ForegroundColor Red		
	# }
	
	#$HomePageName = Read-Host 'What is name of the Page for Article Page Layout'	
	#$IsPage = Get-PnPClientSidePage -Identity $HomePageName -ErrorAction SilentlyContinue
	#if($IsPage -eq $null){
	#	Write-Host "Page does not exist or is not modern page" -ForegroundColor Green
	#	Add-PnPClientSidePage -Name $HomePageName -LayoutType Home
	#}
	#else{
	#	Write-Host "Page: $($ArticlePageName) is present" -ForegroundColor Red		
	#}
	
	# Write-Host "Adding default webpart on modern page" -ForegroundColor Green
	# $WebpartName = Read-Host 'What is name of the webpart for Article Page Layout'
	# Add-PnPClientSideWebpart -Page $ArticlePageName -DefaultWebPartType $WebpartName
	# Write-Host "Webpart added successfully!" -ForegroundColor Green
	
	function ConvertTo-Hashtable {
		[CmdletBinding()]
		[OutputType('hashtable')]
		param (
			[Parameter(ValueFromPipeline)]
			$InputObject
		)
	 
		process {
			if ($null -eq $InputObject) {
				return $null
			}
			if ($InputObject -is [System.Collections.IEnumerable] -and $InputObject -isnot [string]) {
				$collection = @(
					foreach ($object in $InputObject) {
						ConvertTo-Hashtable -InputObject $object
					}
				)

				Write-Output -NoEnumerate $collection
			} elseif ($InputObject -is [psobject]) {
				$hash = @{}
				foreach ($property in $InputObject.PSObject.Properties) {
					$hash[$property.Name] = ConvertTo-Hashtable -InputObject $property.Value
				}
				$hash
			} else {
				$InputObject
			}
		}
	}	
	
	$jsonObj = '{"title":"Site Contact","layout":2,"persons":[{"id":"i:0#.f|membership|heema@sharmadev1.onmicrosoft.com"}]}'
	$wp = $jsonObj | ConvertFrom-Json | ConvertTo-HashTable
			
	#Add PEOPLE webpart with default user, to column 1 in section 1
	#Add-PnPClientSideWebPart -Page "Home" -DefaultWebPartType People -Section 1 -column 1 -WebPartProperties $wp
	
	#Add TEXT webpart to page
	Add-PnPClientSideText -Page "TestingHome" -Text "This is testing the client side Text webpart"	
	
	#Finally publish the page
	Set-PnPClientSidePage -Identity "TestingHome" -Publish
	
	#Get-PnPWebPart -ServerRelativePageUrl "/sites/demomarketing/SitePages/Article.aspx"

    Write-Host "Done." -ForegroundColor Green
}
