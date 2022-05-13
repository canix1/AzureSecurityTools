 <#
.Synopsis
    Get-AzurePrivRoleGroups.ps1
     
    AUTHOR: Robin Granberg (robin.granberg@protonmail.com)
    
    THIS CODE-SAMPLE IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED 
    OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR 
    FITNESS FOR A PARTICULAR PURPOSE.
    
.DESCRIPTION
    A script that will globally search for objects in your Azure AD tenant and return the object and the object type

.EXAMPLE
    .\FindGlobalObjects.ps1 -TenantID "2e5097a7-4ead-42ae-82ef-c33d910626f6" -ObjectID "62e90394-69f5-4237-9190-012177145e10"

.OUTPUTS
    Object properties with an additional property to identify the type

.LINK
    

.NOTES
    **Version: 1.0**

    **20 January, 2022**


#>
Param
(
       # Tenant ID
    [Alias("tenant")]
    [Parameter(Mandatory=$false, 
                Position=1,
                ParameterSetName='Default')]
    [validatescript({$_ -match '(?im)^[{(]?[0-9A-F]{8}[-]?(?:[0-9A-F]{4}[-]?){3}[0-9A-F]{12}[)}]?$'})]
    [ValidateNotNull()]
    [ValidateNotNullOrEmpty()]
    [String] 
    $TenantId = "fa7e80a4-187b-4a6d-8c4a-8bbb1db67c6c",
    [string]
    $HTMLFile = "C:\temp\Doughnut.htm"


)

$VerbosePreference = "Continue"
#List of possible colors for data points
$arrColors = new-object System.Collections.ArrayList

[VOID]$arrColors.add([System.Drawing.Color]::BlueViolet)
[VOID]$arrColors.add([System.Drawing.Color]::DeepSkyBlue)
[VOID]$arrColors.add([System.Drawing.Color]::Turquoise)
[VOID]$arrColors.add([System.Drawing.Color]::DarkTurquoise)
[VOID]$arrColors.add([System.Drawing.Color]::LimeGreen)
[VOID]$arrColors.add([System.Drawing.Color]::Aquamarine)
[VOID]$arrColors.add([System.Drawing.Color]::Aqua)
[VOID]$arrColors.add([System.Drawing.Color]::SpringGreen)
[VOID]$arrColors.add([System.Drawing.Color]::SteelBlue)
[VOID]$arrColors.add([System.Drawing.Color]::Navy)
[VOID]$arrColors.add([System.Drawing.Color]::Teal)
[VOID]$arrColors.add([System.Drawing.Color]::MidnightBlue)
[VOID]$arrColors.add([System.Drawing.Color]::LightSlateGray)
[VOID]$arrColors.add([System.Drawing.Color]::LightSteelBlue)
[VOID]$arrColors.add([System.Drawing.Color]::DimGray)
[VOID]$arrColors.add([System.Drawing.Color]::DarkCyan)
[VOID]$arrColors.add([System.Drawing.Color]::Magenta)
80..300 |%{[VOID]$arrColors.Add([System.Drawing.Color]::$(([System.Drawing.Color] | gm -Static -MemberType Properties)[$_].Name))}

#==========================================================================
# Function		: New-DoughnutChartMutpleDataPoints
# Arguments     : Chart Object, Serie Name, Legend Text, CSV Data,Background color, Color 1, Color 1, Number of Chart Areas in the same Chart Object
# Returns   	: 
# Description   : Draw Doughnut Chart Object in Chart Area
#==========================================================================
Function New-DoughnutChartMutpleDataPoints
{
    param(
    $chart1,$Legend,$CSV,[string]$BackColor,$arrColors,$ChartCount)
    
    $SerieName = "Serie1"
    $Arial = new-object System.Drawing.FontFamily("Arial")
    $Font = new-object System.Drawing.Font($Arial,12 ,[System.Drawing.FontStyle]::Bold)

    $ChartCounterPosition = $chart1.ChartAreas.count
    $ChartElementPosition = new-object System.Windows.Forms.DataVisualization.Charting.ElementPosition((0+($ChartCounterPosition * ((100/$ChartCount)))),12,((100/($ChartCount*0.98))-5),100)
    $chartarea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
    $chartarea.Name = $SerieName
    $chartarea.Position = $ChartElementPosition
    $chartarea.BackColor = $BackColor
    [void]$chart1.ChartAreas.Add($chartarea)


    if($Legend -ne "")
    {
        $NewTitle = New-Object System.Windows.Forms.DataVisualization.Charting.Title
        $NewTitle.Text = $Legend
        $NewTitle.Name = "ChartTitle"+$chart1.Titles.count
        $TitlePosition = new-object System.Windows.Forms.DataVisualization.Charting.ElementPosition((0+($ChartCounterPosition * ((100/$ChartCount)))),(100-(135-(2.5*$ChartCount))),((100/($ChartCount*0.98))-5),100)
        $NewTitle.Position= $TitlePosition
        [void]$Chart1.Titles.Add($NewTitle);
        $chart1.Titles[($chart1.Titles.count-1)].Font = $Font 
        $chart1.Titles[($chart1.Titles.count-1)].ForeColor = [System.Drawing.Color]::White
        $Chart1.Titles[($chart1.Titles.count-1)].DockedToChartArea = $chartarea.Name
    }

        

    [void]$chart1.Series.Add($SerieName)
    $chart1.Series[$SerieName].ChartType = "Doughnut"
    $chart1.Series[$SerieName].SetCustomProperty("DoughnutRadius","50")
    $chart1.Series[$SerieName].IsVisibleInLegend = $true
    $chart1.Series[$SerieName].chartarea = $chartarea.Name
 
    $chart1.Series[$SerieName].LabelForeColor = [System.Drawing.Color]::White
    $chart1.Series[$SerieName].BorderColor = $BackColor
    $chart1.Series[$SerieName].BorderWidth = 5
    $chart1.Series[$SerieName].Font = $Font
    $chart1.Series[$SerieName].IsValueShownAsLabel = $true
    $chart1.Series[$SerieName].IsXValueIndexed = $true

    $colHeaders = ( $CSV | Get-member -MemberType 'NoteProperty' | Select-Object -ExpandProperty 'Name')

    $i = 0    
    $Total = 0
    #For all headers create a data point
    Foreach ($ColumnName in $colHeaders )
    {
        $CSV | ForEach-Object {
            $Point = $chart1.Series[$SerieName].Points.addxy( "$ColumnName" , ($_.$ColumnName)) 
            $Total = $Total + $_.$ColumnName
        }
        $chart1.Series[$SerieName].Points[$Point].Color = $arrColors[$i]
        if($i -ge $($arrColors.count -1))
        {
            $i = 0
        }
        else {
            $i++
        }
        
    }    


    $TextAnno = New-Object System.Windows.Forms.DataVisualization.Charting.TextAnnotation
    $TextAnno.Text = $Total
    $TextAnno.Width = ((100/($ChartCount*0.98))-5)
    $TextAnno.X = (0+($ChartCounterPosition * ((100/$ChartCount))))
    $TextAnno.Y = 53
    $TextAnno.Font = "Segoe UI Black,20pt"
    $TextAnno.ForeColor = [System.Drawing.Color]::White
    $TextAnno.BringToFront()
    [void]$chart1.Annotations.Add($TextAnno)
    

}
#==========================================================================
# Function		: Create-ChartDoughnutMutpleDataPoints
# Arguments     : Chart Title, picuture file, CSV data , Background Color, Doughnut Color
# Returns   	: 
# Description   : Create Chart Object and save to png file/filestream.
# Requires      : Function New-DoughnutChartMutpleDataPoints
#==========================================================================
Function Create-ChartDoughnutMutpleDataPoints {
    param([string]$ChartTitle,$TitlelinGraph,$picturefile,$CSV,[string]$Backcolor,$arrColors)

    [void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")

    #Get all headers in CSV
    $ChartCount = 2

    ## Chart Object 
    $chart1 = New-object System.Windows.Forms.DataVisualization.Charting.Chart
    $chart1.Width = 200 *$ChartCount
    $chart1.Height = (240 + (5 * $ChartCount))
    $chart1.BackColor = $Backcolor
    ## Title
    [void]$chart1.Titles.Add($ChartTitle)
    $chart1.Titles[0].Font = "Arial,20pt"
    $chart1.Titles[0].Alignment = "topLeft"
    $chart1.Titles[0].ForeColor = [System.Drawing.Color]::White


    ## Legend 
    $Legend = New-Object system.Windows.Forms.DataVisualization.Charting.Legend
    $Legend.name = "Legend1"
    $Legend.Font = "Arial,12pt"
    $Legend.ForeColor = [System.Drawing.Color]::White
    $Legend.BackColor = $Backcolor
    $Legend.MaximumAutoSize = 100
    $Legend.IsDockedInsideChartArea = $false
    $Legend.TextWrapThreshold = 30
    $Legend.Alignment = [System.Drawing.StringAlignment]::Near

    $chart1.Legends.Add($Legend)
    ## Data Series

    $i = 0
    New-DoughnutChartMutpleDataPoints $chart1 $TitlelinGraph $CSV $Backcolor $arrColors 2
       
    # Save Chart
    $chart1.SaveImage($picturefile,"png")

}

#==========================================================================
# Function		: Add-DoughnutGraph
# Arguments     : Adds doughnut graph as a picuture file from CSV data 
# Returns   	: 
# Description   : Create Chart Object and save to png file/filestream input the data in a HTML string
# Requires      : Create-ChartDoughnutMutpleDataPoints,New-DoughnutChartMutpleDataPoints
#==========================================================================
Function Add-DoughnutGraph
{
    param($Data,$GraphTitle,$DoughnutTitle,$arrColors)
$PNGFileName = New-Object System.IO.MemoryStream


Create-ChartDoughnutMutpleDataPoints $GraphTitle $DoughnutTitle $PNGFileName $Data "#131313" $arrColors

$Stat = [convert]::ToBase64String($PNGFileName.ToArray())


$strHTMLGraph = "<img src=""data:image/png;base64,$Stat "" />"

return $strHTMLGraph

}

#==========================================================================
# Function		: New-DoughnutChartMutpleDataPoints
# Arguments     : Chart Object, Serie Name, Legend Text, CSV Data,Background color, Color 1, Color 1, Number of Chart Areas in the same Chart Object
# Returns   	: 
# Description   : Draw Doughnut Chart Object in Chart Area
#==========================================================================
Function New-BigDoughnutChartMutpleDataPoints
{
    param(
    $chart1,$Legend,$CSV,[string]$BackColor,$arrColors,$ChartCount)
    
    $SerieName = "Serie1"
    $Arial = new-object System.Drawing.FontFamily("Arial")
    $Font = new-object System.Drawing.Font($Arial,12 ,[System.Drawing.FontStyle]::Bold)

    $ChartCounterPosition = $chart1.ChartAreas.count
    $ChartElementPosition = new-object System.Windows.Forms.DataVisualization.Charting.ElementPosition((0+($ChartCounterPosition * ((100/$ChartCount)))),12,((100/($ChartCount*0.98))-5),100)
    $chartarea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
    $chartarea.Name = $SerieName
    $chartarea.Position = $ChartElementPosition
    $chartarea.BackColor = $BackColor
    [void]$chart1.ChartAreas.Add($chartarea)


    if($Legend -ne "")
    {
        $NewTitle = New-Object System.Windows.Forms.DataVisualization.Charting.Title
        $NewTitle.Text = $Legend
        $NewTitle.Name = "ChartTitle"+$chart1.Titles.count
        $TitlePosition = new-object System.Windows.Forms.DataVisualization.Charting.ElementPosition((0+($ChartCounterPosition * ((100/$ChartCount)))),(100-(135-(2.5*$ChartCount))),((100/($ChartCount*0.98))-5),100)
        $NewTitle.Position= $TitlePosition
        [void]$Chart1.Titles.Add($NewTitle);
        $chart1.Titles[($chart1.Titles.count-1)].Font = $Font 
        $chart1.Titles[($chart1.Titles.count-1)].ForeColor = [System.Drawing.Color]::White
        $Chart1.Titles[($chart1.Titles.count-1)].DockedToChartArea = $chartarea.Name
    }

        

    [void]$chart1.Series.Add($SerieName)
    $chart1.Series[$SerieName].ChartType = "Doughnut"
    $chart1.Series[$SerieName].SetCustomProperty("DoughnutRadius","50")
    $chart1.Series[$SerieName].IsVisibleInLegend = $true
    $chart1.Series[$SerieName].chartarea = $chartarea.Name
 
    $chart1.Series[$SerieName].LabelForeColor = [System.Drawing.Color]::White
    $chart1.Series[$SerieName].BorderColor = $BackColor
    $chart1.Series[$SerieName].BorderWidth = 5
    $chart1.Series[$SerieName].Font = $Font
    $chart1.Series[$SerieName].IsValueShownAsLabel = $true
    $chart1.Series[$SerieName].IsXValueIndexed = $true

    $colHeaders = ( $CSV | Get-member -MemberType 'NoteProperty' | Select-Object -ExpandProperty 'Name')

    $i = 0    
    $Total = 0
    #For all headers create a data point
    Foreach ($ColumnName in $colHeaders )
    {
        $CSV | ForEach-Object {
            $Point = $chart1.Series[$SerieName].Points.addxy( "$ColumnName" , ($_.$ColumnName)) 
            $Total = $Total + $_.$ColumnName
        }
        $chart1.Series[$SerieName].Points[$Point].Color = $arrColors[$i]
        if($i -ge $($arrColors.count -1))
        {
            $i = 0
        }
        else {
            $i++
        }
        
    }    


    $TextAnno = New-Object System.Windows.Forms.DataVisualization.Charting.TextAnnotation
    $TextAnno.Text = $Total
    $TextAnno.Width = ((100/($ChartCount*0.98))-5)
    $TextAnno.X = (0+($ChartCounterPosition * ((100/$ChartCount))))
    $TextAnno.Y = 58
    $TextAnno.Font = "Segoe UI Black,20pt"
    $TextAnno.ForeColor = [System.Drawing.Color]::White
    $TextAnno.BringToFront()
    [void]$chart1.Annotations.Add($TextAnno)
    

}
#==========================================================================
# Function		: Create-ChartDoughnutMutpleDataPoints
# Arguments     : Chart Title, picuture file, CSV data , Background Color, Doughnut Color
# Returns   	: 
# Description   : Create Chart Object and save to png file/filestream.
# Requires      : Function New-DoughnutChartMutpleDataPoints
#==========================================================================
Function Create-BigChartDoughnutMutpleDataPoints {
    param([string]$ChartTitle,$TitlelinGraph,$picturefile,$CSV,[string]$Backcolor,$arrColors)

    [void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")

    #Get all headers in CSV
    $ChartCount = 2

    ## Chart Object 
    $chart1 = New-object System.Windows.Forms.DataVisualization.Charting.Chart
    $chart1.Width = 800
    $chart1.Height = 600
    $chart1.BackColor = $Backcolor
    ## Title
    [void]$chart1.Titles.Add($ChartTitle)
    $chart1.Titles[0].Font = "Arial,20pt"
    $chart1.Titles[0].Alignment = "topLeft"
    $chart1.Titles[0].ForeColor = [System.Drawing.Color]::White


    ## Legend 
    $Legend = New-Object system.Windows.Forms.DataVisualization.Charting.Legend
    $Legend.name = "Legend1"
    $Legend.Font = "Arial,12pt"
    $Legend.ForeColor = [System.Drawing.Color]::White
    $Legend.BackColor = $Backcolor
    $Legend.MaximumAutoSize = 100
    $Legend.IsDockedInsideChartArea = $false
    $Legend.TextWrapThreshold = 30
    $Legend.Alignment = [System.Drawing.StringAlignment]::Near

    $chart1.Legends.Add($Legend)
    ## Data Series

    $i = 0
    New-BigDoughnutChartMutpleDataPoints $chart1 $TitlelinGraph $CSV $Backcolor $arrColors 2
       
    # Save Chart
    $chart1.SaveImage($picturefile,"png")

}

#==========================================================================
# Function		: Add-DoughnutGraph
# Arguments     : Adds doughnut graph as a picuture file from CSV data 
# Returns   	: 
# Description   : Create Chart Object and save to png file/filestream input the data in a HTML string
# Requires      : Create-ChartDoughnutMutpleDataPoints,New-DoughnutChartMutpleDataPoints
#==========================================================================
Function Add-BigDoughnutGraph
{
    param($Data,$GraphTitle,$DoughnutTitle,$arrColors)
$PNGFileName = New-Object System.IO.MemoryStream


Create-BigChartDoughnutMutpleDataPoints $GraphTitle $DoughnutTitle $PNGFileName $Data "#131313" $arrColors

$Stat = [convert]::ToBase64String($PNGFileName.ToArray())


$strHTMLGraph = "<img src=""data:image/png;base64,$Stat "" />"

return $strHTMLGraph

}


function Get-AzureADIRApiToken {

    ############################################################################

    <#
    .SYNOPSIS

        Get an access token for use with the API cmdlets.


    .DESCRIPTION

        Uses MSAL.ps to obtain an access token. Has an option to refresh an existing token.


    .EXAMPLE

        Get-AzureADIRApiToken -TenantId b446a536-cb76-4360-a8bb-6593cf4d9c7f

        Gets or refreshes an access token for making API calls for the tenant ID
        b446a536-cb76-4360-a8bb-6593cf4d9c7f.


    .EXAMPLE

        Get-AzureADIRApiToken -TenantId b446a536-cb76-4360-a8bb-6593cf4d9c7f -ForceRefresh

        Refreshes an access token for making API calls for the tenant ID
        b446a536-cb76-4360-a8bb-6593cf4d9c7f.


    .EXAMPLE

        Get-AzureADIRApiToken -TenantId b446a536-cb76-4360-a8bb-6593cf4d9c7f -LoginHint Bob@Contoso.com

        Gets or refreshes an access token for making API calls for the tenant ID
        b446a536-cb76-4360-a8bb-6593cf4d9c7f and user Bob@Contoso.com.


    .EXAMPLE

        Get-AzureADIRApiToken -TenantId b446a536-cb76-4360-a8bb-6593cf4d9c7f -InterActive

        Gets or refreshes an access token for making API calls for the tenant ID
        b446a536-cb76-4360-a8bb-6593cf4d9c7f. Ensures a pop-up box appears.

    #>

    ############################################################################

    [CmdletBinding(DefaultParameterSetName="InterActive")]
    param(

        #The tenant ID
        [Parameter(Mandatory,Position=0)]
        [guid]$TenantId,

        #Force a token refresh
        [Parameter(Position=1,ParameterSetName="ForceRefresh")]
        [switch]$ForceRefresh,

        #The user's upn used for the login hint
        [Parameter(Position=2,ParameterSetName="InterActive")]
        [string]$LoginHint,

        #Force a pop-up box
        [Parameter(Position=3,ParameterSetName="InterActive")]
        [switch]$InterActive,

        #get an Azure AD Graph token
        [Parameter(Position=4)]
        [switch]$AadGraph

    )


    ############################################################################


    #Get an access token using the PowerShell client ID
    $ClientId = "1b730954-1685-4b74-9bfd-dac224a7b894" 
    #$RedirectUri = "urn:ietf:wg:oauth:2.0:oob"
    $Authority = "https://login.microsoftonline.com/$TenantId"

    if ($AadGraph) {

        $Scopes = "https://graph.windows.net/.default"

    }
    else {
    
        $Scopes = "https://graph.microsoft.com/.default"

    }
    

    if ($ForceRefresh) {

        Write-Verbose -Message "$(Get-Date -f T) - Attempting to refresh an existing access token"

        #Attempt to refresh access token
        try {

            $Response = Get-MsalToken -ClientId $ClientId -RedirectUri $RedirectUri -Authority $Authority -Scopes $Scopes -ForceRefresh
        }
        catch {}

        #Error handling for token acquisition
        if ($Response) {

            Write-Verbose -Message "$(Get-Date -f T) - API Access Token refreshed - new expiry: $(($Response).ExpiresOn.UtcDateTime)"

            return $Response

        }
        else {
            
            Write-Warning -Message "$(Get-Date -f T) - Failed to refresh Access Token - try re-running the cmdlet again"

        }

    }
    elseif ($LoginHint) {

        Write-Verbose -Message "$(Get-Date -f T) - Checking token cache with -LoginHint for $LoginHint"

        #Run this to obtain an access token - should prompt on first run to select the account used for future operations
        try {

            if ($InterActive) {

                $Response = Get-MsalToken -ClientId $ClientId -RedirectUri $RedirectUri -Authority $Authority -LoginHint $LoginHint -Scopes $Scopes -Interactive

            } 
            else {

                $Response = Get-MsalToken -ClientId $ClientId -RedirectUri $RedirectUri -Authority $Authority -LoginHint $LoginHint -Scopes $Scopes 

            }
        }
        catch {}

        #Error handling for token acquisition
        if ($Response) {

            Write-Verbose -Message "$(Get-Date -f T) - API Access Token obtained for: $(($Response).Account.Username) ($(($Response).Account.HomeAccountId.ObjectId))"
            #Write-Verbose -Message "$(Get-Date -f T) - API Access Token scopes: $(($Response).Scopes)"

            return $Response

        }
        else {

            Write-Warning -Message "$(Get-Date -f T) - Failed to obtain an Access Token - try re-running the cmdlet again"
            Write-Warning -Message "$(Get-Date -f T) - If the problem persists, use `$Error[0] for more detail on the error or start a new PowerShell session"

        }

    }
    else {

        Write-Verbose -Message "$(Get-Date -f T) - Checking token cache with -Prompt"

        #Run this to obtain an access token - should prompt on first run to select the account used for future operations
        try {

            if ($InterActive) {

                $Response = Get-MsalToken -ClientId $ClientId -RedirectUri $RedirectUri -Authority $Authority -Prompt SelectAccount -Interactive -Scopes $Scopes 

            }
            else {

                $Response = Get-MsalToken -ClientId $ClientId -RedirectUri $RedirectUri -Authority $Authority -Prompt SelectAccount -Scopes $Scopes 

            }

        }
        catch {}

        #Error handling for token acquisition
        if ($Response) {

            Write-Verbose -Message "$(Get-Date -f T) - API Access Token obtained for: $(($Response).Account.Username) ($(($Response).Account.HomeAccountId.ObjectId))"
            #Write-Verbose -Message "$(Get-Date -f T) - API Access Token scopes: $(($Response).Scopes)"

            return $Response

        }
        else {

            Write-Warning -Message "$(Get-Date -f T) - Failed to obtain an Access Token - try re-running the cmdlet again"
            Write-Warning -Message "$(Get-Date -f T) - If the problem persists, run Connect-AzureADIR with the -UserUpn parameter"

        }

    }


}   

function Get-AzureADIRHeader {

    ############################################################################

    <#
    .SYNOPSIS

        Uses a supplied Access Token to construct a header for a an API call.


    .DESCRIPTION

        Uses a supplied Access Token to construct a header for a an API call with 
        Invoke-WebRequest.

        Can supply the ConsistencyLevel = Eventual parameter for performing Count
        activities.


    .EXAMPLE

        Get-AzureADIRHeader -Token $Token

        Constructs a header with an obtained token for using with Invoke-WebRequest.


    .EXAMPLE

        Get-AzureADIRHeader -Token $Token -ConsistencyLevelEventual

        Constructs a header with an obtained token for using with Invoke-WebRequest.

        Uses the optional -ConsistencyLevelEventual switch for use in conjunction with
        the count call.

    #>

    ############################################################################
    
    [CmdletBinding()]
    param(

        #The tenant ID
        [Parameter(Mandatory,Position=0)]
        [string]$Token,

        #Switch to include ConsistencyLevel = Eventual for $count operations
        [Parameter(Position=1)]
        [switch]$ConsistencyLevelEventual

        )

    ############################################################################

    if ($ConsistencyLevelEventual) {

        return @{

            "Authorization" = ("Bearer {0}" -f $Token);
            "Content-Type" = "application/json";
                "ConsistencyLevel" = "eventual";

        }

    }
    else {

        return @{

            "Authorization" = ("Bearer {0}" -f $Token);
            "Content-Type" = "application/json";

        }

    }

}   #end function

function Invoke-AzureADIRWebRequest {

    ############################################################################

    <#
    .SYNOPSIS

        Perform Invoke-WebRequest with additional error handling.


    .DESCRIPTION

        Perform Invoke-WebRequest with additional error handling for supplied
        query URL and authentication header.

        Has retry logic.

    .EXAMPLE

        Invoke-AzureADIRWebRequest -Header $Header -Url $Url

        Calls Invoke-Webrequest with the supplied authentication header and query
        URL with error checking and retry logic.


    #>

    ############################################################################

    [CmdletBinding()]
    param(

        #The header for the API call
        [Parameter(Mandatory,Position=0)]
        $Header,

        #the query Url 
        [Parameter(Mandatory,Position=1)]
        [string]$Url

        )

    ############################################################################
    
    $RetryCount = 0


    ##################################
    #Do our stuff with error handling
    try {

        #Invoke the web request
        $MyReport = (Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $Url -Verbose:$false)

    }
    catch [System.Net.WebException] {
        
        $StatusCode = [int]$_.Exception.Response.StatusCode
        Write-Warning -Message "$(Get-Date -f T) - $($_.Exception.Message)"

        #Check what's gone wrong
        if (($StatusCode -eq 401) -and ($OneSuccessfulFetch)) {

            #Token might have expired; renew token and try again
            $Token = (Get-AzureADIRApiToken -TenantId $TenantId -InterActive).AccessToken
            $Header = Get-AzureADIRHeader -Token $Token
            $OneSuccessfulFetch = $False

        }
        elseif (($StatusCode -eq 429) -or ($StatusCode -eq 504) -or ($StatusCode -eq 503)) {

            #Throttled request or a temporary issue, wait for a few seconds and retry
            Start-Sleep -Seconds 5

        }
        elseif (($StatusCode -eq 403) -or ($StatusCode -eq 401)) {

            Write-Warning -Message "$(Get-Date -f T) - Please check the permissions of the user"
            break

        }
        elseif ($StatusCode -eq 400) {

            Write-Warning -Message "$(Get-Date -f T) - Please check the query used"
            break

        }
        else {
            
            #Retry up to 5 times
            if ($RetryCount -lt 5) {
                
                write-output "Retrying..."
                $RetryCount++

            }
            else {
                
                #Write to host and exit loop
                Write-Warning -Message "$(Get-Date -f T) - Download request failed. Please try again in the future"
                break

            }

        }

    }
    catch {

        #Write error details to host
        Write-Warning -Message "$(Get-Date -f T) - $($_.Exception)"


        #Retry up to 5 times    
        if ($RetryCount -lt 5) {

            write-output "Retrying..."
            $RetryCount++

        }
        else {

            #Write to host and exit loop
            Write-Warning -Message "$(Get-Date -f T) - Download request failed - please try again in the future"
            break

        }

    } # end try / catch


    return $MyReport


}   #end function

function Get-RoleDefinitions {
     ############################################################################

    <#
    .SYNOPSIS

        Search Azure AD for objects using ObjectID,Id,DeviceID,TemplateID 


    .DESCRIPTION

        Function to find user,group,application,device,application and administrativeunit

    .EXAMPLE

        FindGlobalObject $Token $Header $TenantID $ObjectID


    #>

    ############################################################################

     [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        $Token,

        [Parameter(Mandatory)]
        $Header,

        [Parameter(Mandatory)]
        $TenantID

    )
    
    

    $ResponseData = $null

    Write-Verbose -Message "$(Get-Date -f T) - Looking up role definitions for all roles"

    #All Users 

    $Url = "https://graph.microsoft.com/beta/privilegedAccess/aadroles/resources/$TenantId/roleDefinitions?`&`$Select=id,displayName,Type"

    
    try {

        # Convert the content in the response from Json and expand all values
       $ResponseData = (Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $Url -Verbose:$true ).Content | ConvertFrom-Json | Select-Object -ExpandProperty Value

    }    
    catch {}

    Return $ResponseData 


} # End Function


function Get-RoleAssignments{
    ############################################################################

   <#
   .SYNOPSIS

       Search Azure AD for objects using ObjectID,Id,DeviceID,TemplateID 


   .DESCRIPTION

       Function to find user,group,application,device,application and administrativeunit

   .EXAMPLE

       FindGlobalObject $Token $Header $TenantID $ObjectID


   #>

   ############################################################################

    [CmdletBinding()]
   param (
       [Parameter(Mandatory)]
       $Token,

       [Parameter(Mandatory)]
       $Header,

       [Parameter(Mandatory)]
       $TenantID

   )
   
   

   $ResponseData = $null

   Write-Verbose -Message "$(Get-Date -f T) - Looking up role definitions for all roles"

   #All Users 

   $Url = "https://graph.microsoft.com/beta/privilegedAccess/aadroles/resources/$TenantId/roleAssignments"
   
   try {

       # Convert the content in the response from Json and expand all values
      $ResponseData = (Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $Url -Verbose:$true ).Content | ConvertFrom-Json | Select-Object -ExpandProperty Value

   }    
   catch {}

   Return $ResponseData 


} # End Function

function FindGlobalObject {
    ############################################################################

   <#
   .SYNOPSIS

       Search Azure AD for objects using ObjectID,Id,DeviceID,TemplateID 


   .DESCRIPTION

       Function to find user,group,application,device,application and administrativeunit

   .EXAMPLE

       FindGlobalObject $Token $Header $TenantID $ObjectID


   #>

   ############################################################################

    [CmdletBinding()]
   param (
       [Parameter(Mandatory)]
       $Token,

       [Parameter(Mandatory)]
       $Header,

       [Parameter(Mandatory)]
       $TenantID,

       [Parameter()]
       [string]
       $ObjectID,

       [Parameter()]
       [string]
       $Properties
   )
   
   

    $ResponseData = $null
   
    #All Users 
    if($Properties)
    {
        $Url = "https://graph.microsoft.com/beta/users?`$filter=startsWith(displayName,'$ObjectID') OR startsWith(givenName,'$ObjectID') OR startsWith(surName,'$ObjectID') OR startsWith(mail,'$ObjectID') OR startsWith(userPrincipalName,'$ObjectID') OR id eq '$ObjectID'&`$Select='$Properties'"
    }
    else {
        $Url = "https://graph.microsoft.com/beta/users?`$filter=startsWith(displayName,'$ObjectID') OR startsWith(givenName,'$ObjectID') OR startsWith(surName,'$ObjectID') OR startsWith(mail,'$ObjectID') OR startsWith(userPrincipalName,'$ObjectID') OR id eq '$ObjectID'"     
    }
   
    try {

       # Convert the content in the response from Json and expand all values
       $ResponseData = (Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $Url -Verbose:$false).Content | ConvertFrom-Json | Select-Object -ExpandProperty Value

   }    
   catch {}

   
   if ($ResponseData) {
       add-member -InputObject $ResponseData -MemberType NoteProperty -Name "type" -Value "user"
       Return $ResponseData 
   }

   #All Groups
   if($Properties)
   {
        $Url = "https://graph.microsoft.com/beta/groups?`$filter=startsWith(displayName,'$ObjectID') OR startsWith(mail,'$ObjectID')&`$Select='$Properties'"
   }
   else {   
        $Url = "https://graph.microsoft.com/beta/groups?`$filter=startsWith(displayName,'$ObjectID') OR startsWith(mail,'$ObjectID')"
   }
   try {

       # Convert the content in the response from Json and expand all values
       $ResponseData = (Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $Url -Verbose:$false).Content | ConvertFrom-Json | Select-Object -ExpandProperty Value

   }    
   catch {}

   if ($ResponseData) {
       add-member -InputObject $ResponseData -MemberType NoteProperty -Name "type" -Value "group"
       Return $ResponseData 
   }


   #All Groups ID
   if($Properties)
   {
        $Url = "https://graph.microsoft.com/beta/groups?`$filter=id eq '$ObjectID'&`$Select='$Properties'"
   }
   else {      
        $Url = "https://graph.microsoft.com/beta/groups?`$filter=id eq '$ObjectID'"
   }
   try {

       # Convert the content in the response from Json and expand all values
       $ResponseData = (Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $Url -Verbose:$false).Content | ConvertFrom-Json | Select-Object -ExpandProperty Value

   }    
   catch {}


   if ($ResponseData) {
       add-member -InputObject $ResponseData -MemberType NoteProperty -Name "type" -Value "group"        
       
        Return $ResponseData 
   }


    #All Applications
    if($Properties)
    {
        $Url = "https://graph.microsoft.com/beta/applications?`$filter=startsWith(displayName,'$ObjectID') OR id eq '$ObjectID' OR appId eq '$ObjectID'&`$Select='$Properties'"
    }
    else {    
        $Url = "https://graph.microsoft.com/beta/applications?`$filter=startsWith(displayName,'$ObjectID') OR id eq '$ObjectID' OR appId eq '$ObjectID'"
    }
    try {

       # Convert the content in the response from Json and expand all values
       $ResponseData = (Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $Url -Verbose:$false).Content | ConvertFrom-Json | Select-Object -ExpandProperty Value

    }    
    catch {}

   if ($ResponseData) {
       add-member -InputObject $ResponseData -MemberType NoteProperty -Name "type" -Value "application"
       Return $ResponseData 
   }

    #All Devices
    if($Properties)
    {
        $Url = "https://graph.microsoft.com/beta/devices?`$filter=startswith(displayName,'$ObjectID')&`$Select='$Properties'"
    }
    else {     
        $Url = "https://graph.microsoft.com/beta/devices?`$filter=startswith(displayName,'$ObjectID')"
    }
   try {

       # Convert the content in the response from Json and expand all values
       $ResponseData = (Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $Url -Verbose:$false).Content | ConvertFrom-Json | Select-Object -ExpandProperty Value

   }    
   catch {}

   if ($ResponseData) {
       add-member -InputObject $ResponseData -MemberType NoteProperty -Name "type" -Value "device"
       Return $ResponseData 
   }

   #Device DeviceID
    if($Properties)
    {
        $Url = "https://graph.microsoft.com/beta/devices?`$filter=deviceId eq '$ObjectID'&`$Select='$Properties'"
    }
    else {    
        $Url = "https://graph.microsoft.com/beta/devices?`$filter=deviceId eq '$ObjectID'"
    }
    try {

       # Convert the content in the response from Json and expand all values
       $ResponseData = (Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $Url -Verbose:$false).Content | ConvertFrom-Json | Select-Object -ExpandProperty Value

    }    
   catch {}

   if ($ResponseData) {
       add-member -InputObject $ResponseData -MemberType NoteProperty -Name "type" -Value "device"
       Return $ResponseData 
   }

   #Device id
    if($Properties)
    {
        $Url = "https://graph.microsoft.com/beta/devices?`$filter=id eq '$ObjectID'&`$Select='$Properties'"
    }
    else {       
        $Url = "https://graph.microsoft.com/beta/devices?`$filter=id eq '$ObjectID'"
    }
   try {

       # Convert the content in the response from Json and expand all values
       $ResponseData = (Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $Url -Verbose:$false).Content | ConvertFrom-Json | Select-Object -ExpandProperty Value

   }    
   catch {}

   if ($ResponseData) {
       add-member -InputObject $ResponseData -MemberType NoteProperty -Name "type" -Value "device"
       Return $ResponseData 
   }
   
   #Role TemplateId
   if($Properties)
   {
    $Url = "https://graph.microsoft.com/beta/privilegedAccess/aadRoles/resources/$TenantID/roleDefinitions?"+"$"+"filter=templateID+eq+'$ObjectID'&`$Select='$Properties'"
   }
   else {     
        $Url = "https://graph.microsoft.com/beta/privilegedAccess/aadRoles/resources/$TenantID/roleDefinitions?"+"$"+"filter=templateID+eq+'$ObjectID'"
   }
   try {

       # Convert the content in the response from Json and expand all values
       $ResponseData = (Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $Url -Verbose:$false).Content | ConvertFrom-Json | Select-Object -ExpandProperty Value

   }    
   catch {}

   if ($ResponseData) {
       Return $ResponseData 
   }
   
   ##Search  all ServicePrincipals
   if($Properties)
   {
        $Url = "https://graph.microsoft.com/beta/servicePrincipals?"+"$"+"filter=startswith(displayName,'$ObjectID') OR id eq '$ObjectID' OR appId eq '$ObjectID'&`$Select='$Properties'"
   }
   else {     
    $Url = "https://graph.microsoft.com/beta/servicePrincipals?"+"$"+"filter=startswith(displayName,'$ObjectID') OR id eq '$ObjectID' OR appId eq '$ObjectID'"
   }
   try {

       # Convert the content in the response from Json and expand all values
       $ResponseData = (Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $Url -Verbose:$false).Content | ConvertFrom-Json | Select-Object -ExpandProperty Value

   }    
   catch {}

   if ($ResponseData) {
       add-member -InputObject $ResponseData -MemberType NoteProperty -Name "type" -Value "application"
       Return $ResponseData 
   }

   #Search Administrative Units
   if($Properties)
   {
        $Url = "https://graph.microsoft.com/beta/administrativeUnits?"+"$"+"filter=startswith(displayName,'$ObjectID') OR id eq '$ObjectID'&`$Select='$Properties'"
   }
   else {    
        $Url = "https://graph.microsoft.com/beta/administrativeUnits?"+"$"+"filter=startswith(displayName,'$ObjectID') OR id eq '$ObjectID'"
   }
   try {

       # Convert the content in the response from Json and expand all values
       $ResponseData = (Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $Url -Verbose:$false).Content | ConvertFrom-Json | Select-Object -ExpandProperty Value

   }    
   catch {}

   if ($ResponseData) {
       add-member -InputObject $ResponseData -MemberType NoteProperty -Name "type" -Value "administrativeunit"
       Return $ResponseData 
   }    
} # End Function


function Get-RoleMembers {
    ############################################################################

   <#
   .SYNOPSIS

       Search Azure AD for objects using ObjectID,Id,DeviceID,TemplateID 


   .DESCRIPTION

       Function to find user,group,application,device,application and administrativeunit

   .EXAMPLE

       FindGlobalObject $Token $Header $TenantID $ObjectID


   #>

   ############################################################################

    [CmdletBinding()]
   param (
       [Parameter(Mandatory)]
       $Token,

       [Parameter(Mandatory)]
       $Header,

       [Parameter(Mandatory)]
       $TenantID,

       [Parameter(Mandatory)]
       $RoleID

   )
   
   

   $ResponseData = $null

   Write-Verbose -Message "$(Get-Date -f T) - Looking up role assignments for role details - $(($RoleID))"

   #All Users 
   $Url = "https://graph.microsoft.com/beta/privilegedAccess/aadRoles/resources/$TenantID/roleAssignments?"+"$"+"filter=RoleDefinitionId+eq+'$RoleID'"

   try {

       # Convert the content in the response from Json and expand all values
       $ResponseData = (Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $Url -Verbose:$false).Content | ConvertFrom-Json | Select-Object -ExpandProperty Value

   }    
   catch {}

   
   if ($ResponseData) {
       add-member -InputObject $ResponseData -MemberType NoteProperty -Name "type" -Value "user"
       Return $ResponseData 
   }

} # End Function
function Get-GroupOwner {
    ############################################################################

   <#
   .SYNOPSIS

       Search Azure AD for objects using ObjectID,Id,DeviceID,TemplateID 


   .DESCRIPTION

       Function to find user,group,application,device,application and administrativeunit

   .EXAMPLE

       FindGlobalObject $Token $Header $TenantID $ObjectID


   #>

   ############################################################################

    [CmdletBinding()]
   param (
       [Parameter(Mandatory)]
       $Token,

       [Parameter(Mandatory)]
       $Header,

       [Parameter(Mandatory)]
       $TenantID,

       [Parameter(Mandatory)]
       $ObjectID

   )
   
   
   $ResponseData = $null

   #Get Owner
    $Url = "https://graph.microsoft.com/beta/groups/$ObjectID/owners?" + "$"+" orderby=displayName asc&"+"$"+"count=true"
    try {

        # Convert the content in the response from Json and expand all values
        $ResponseData = (Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $Url -Verbose:$false).Content | ConvertFrom-Json | Select-Object -ExpandProperty Value

    }    
    catch {}
    ####################### - Owner
   
   if ($ResponseData) {
       Return $ResponseData 
   }

} # End Function

function Get-GroupMembers {
    ############################################################################

   <#
   .SYNOPSIS

       Search Azure AD for objects using ObjectID,Id,DeviceID,TemplateID 


   .DESCRIPTION

       Function to find user,group,application,device,application and administrativeunit

   .EXAMPLE

       FindGlobalObject $Token $Header $TenantID $ObjectID


   #>

   ############################################################################

    [CmdletBinding()]
   param (
       [Parameter(Mandatory)]
       $Token,

       [Parameter(Mandatory)]
       $Header,

       [Parameter(Mandatory)]
       $TenantID,

       [Parameter(Mandatory)]
       $ObjectID

   )
   
   
   $ResponseData = $null

   #Get Owner
    $Url = "https://graph.microsoft.com/beta/groups/$ObjectID/transitiveMembers?" +"$"+" orderby=displayName asc&"+"$"+"count=true"
    try {

        # Convert the content in the response from Json and expand all values
        $ResponseData = (Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $Url -Verbose:$false).Content | ConvertFrom-Json | Select-Object -ExpandProperty Value

    }    
    catch {}
    ####################### - Owner
   
   if ($ResponseData) {
       Return $ResponseData 
   }

} # End Function

function Get-PAGGroupMembers {
  ############################################################################

 <#
 .SYNOPSIS

     Search Azure AD for objects using ObjectID,Id,DeviceID,TemplateID 


 .DESCRIPTION

     Function to find user,group,application,device,application and administrativeunit

 .EXAMPLE

     FindGlobalObject $Token $Header $TenantID $ObjectID


 #>

 ############################################################################

  [CmdletBinding()]
 param (
     [Parameter(Mandatory)]
     $Token,

     [Parameter(Mandatory)]
     $Header,

     [Parameter(Mandatory)]
     $TenantID,

     [Parameter(Mandatory)]
     $ObjectID

 )
 
 
 $ResponseData = $null

 #Get Owner
  

  $Url = "https://graph.microsoft.com/beta/privilegedAccess/aadGroups/resources/$ObjectID/roleAssignments?"
  try {

      # Convert the content in the response from Json and expand all values
      $ResponseData = (Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $Url -Verbose:$false).Content | ConvertFrom-Json | Select-Object -ExpandProperty Value

  }    
  catch {}
  ####################### - Owner
 
 if ($ResponseData) {
     Return $ResponseData 
 }

} # End Function


#==========================================================================
# Function		: ConvertTo-ObjectArrayListFromPsCustomObject  
# Arguments     : Defined Object
# Returns   	: Custom Object List
# Description   : Convert a defined object to a custom, this will help you  if you got a read-only object 
# 
#==========================================================================
function ConvertTo-ObjectArrayListFromPsCustomObject 
{ 
     param ( 
         [Parameter(  
             Position = 0,   
             Mandatory = $true,   
             ValueFromPipeline = $true,  
             ValueFromPipelineByPropertyName = $true  
         )] $psCustomObject
     ); 
     
     process {
 
        $myCustomArray = New-Object System.Collections.ArrayList
     
         foreach ($myPsObject in $psCustomObject) { 
             $hashTable = @{}; 
             $myPsObject | Get-Member -MemberType *Property | ForEach-Object { 
                 $hashTable.($_.name) = $myPsObject.($_.name); 
             } 
             $Newobject = new-object psobject -Property  $hashTable
             [void]$myCustomArray.add($Newobject)
         } 
         return $myCustomArray
     } 
 }# End function
 Write-Verbose -Message "$(Get-Date -f T) - Authenticate"
Function Get-TenantInformation
{
    ############################################################################

   <#
   .SYNOPSIS

       Search Azure AD for objects using ObjectID,Id,DeviceID,TemplateID 


   .DESCRIPTION

       Function to find user,group,application,device,application and administrativeunit

   .EXAMPLE

       FindGlobalObject $Token $Header $TenantID $ObjectID


   #>

   ############################################################################

    [CmdletBinding()]
   param (
       [Parameter(Mandatory)]
       $Token,

       [Parameter(Mandatory)]
       $Header

   )
   
   

   $ResponseData = $null

   Write-Verbose -Message "$(Get-Date -f T) - Looking up Tenant Information"

   #Get Domains 
   $Url = "https://graph.microsoft.com/beta/organization"

   try {

       # Convert the content in the response from Json and expand all values
       $ResponseData = (Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $Url -Verbose:$false).Content | ConvertFrom-Json | Select-Object -ExpandProperty Value

   }    
   catch {}

   
   if ($ResponseData) {
       Return $ResponseData 
   }

} # End Function
$LoginToken = Get-AzureADIRApiToken -TenantId $TenantId -LoginHint $UserUpn
$Token =  ($LoginToken).AccessToken
$Header = Get-AzureADIRHeader -Token $Token

##Role
$Roles = Get-RoleDefinitions $Token $Header $TenantID
$RolesAssignments = Get-RoleAssignments $Token $Header $TenantID
#$RolesAssignments
if($true)
{
$RoleMembers = New-Object System.Collections.ArrayList
foreach( $role in $Roles)
{
    $RoleName =  $role.displayName 
    $Members = $RolesAssignments | Where-Object{$_.RoleDefinitionId -eq $role.id}
    $Members| ForEach-Object{
        Write-Verbose -Message "$(Get-Date -f T) - Fetching attributes on $($_.subjectId)"
        $MemberObject = $(ConvertTo-ObjectArrayListFromPsCustomObject  $_)
        if($RoleMembers.subjectId -contains $MemberObject.subjectId)
        {
            $MemberProperties = ($RoleMembers | Where-object {$_.subjectId -eq $MemberObject.subjectId}  | select-object -Property displayName,type,userPrincipalName,isAssignableToRole)[0]
        }
        else {
            $MemberProperties = (FindGlobalObject $Token $Header $TenantID $_.subjectId) | select-object -Property displayName,type,userPrincipalName,isAssignableToRole
        }
        
        Add-Member -InputObject $MemberObject -MemberType NoteProperty -Name Role $RoleName
        Add-Member -InputObject $MemberObject -MemberType NoteProperty -Name displayName $MemberProperties.displayName
        Add-Member -InputObject $MemberObject -MemberType NoteProperty -Name userPrincipalName $MemberProperties.userPrincipalName
        Add-Member -InputObject $MemberObject -MemberType NoteProperty -Name type $MemberProperties.type        
        Add-Member -InputObject $MemberObject -MemberType NoteProperty -Name isAssignableToRole $MemberProperties.isAssignableToRole   
        [VOID]$RoleMembers.Add($MemberObject)
    } 

    
}


if ($RoleMembers | Where-Object{$_.Type -eq "group"})
{
    $MemberTypeGroup = $RoleMembers | Where-Object{$_.Type -eq "group"}
    foreach ($Group in $MemberTypeGroup)
    {
        $Owners = @(Get-GroupOwner $Token $Header $TenantID $Group.subjectId)
        if($Owners)
        {
            foreach($Owner in $Owners)
            {
                $OwnerObject = $(ConvertTo-ObjectArrayListFromPsCustomObject  $Group)
                $OwnerProperties = (FindGlobalObject $Token $Header $TenantID $Owner.id) | select-object -Property displayName,type,userPrincipalName,isAssignableToRole
                Add-Member -InputObject $OwnerObject -MemberType NoteProperty -Name NestedGroupID -value $Group.subjectId
                Add-Member -InputObject $OwnerObject -MemberType NoteProperty -Name NestedGroupdisplayName -value $Group.displayName 
                if($(($OwnerObject | get-member -MemberType NoteProperty ).name.contains("userPrincipalName")))
                {  
                                         
                    $OwnerObject.UserPrincipalName = $OwnerProperties.UserPrincipalName                         
                }
                else
                {
                    Add-Member -InputObject $OwnerObject -MemberType NoteProperty -Name UserPrincipalName $OwnerProperties.userPrincipalName
                }                               
                $OwnerObject.userPrincipalName = $OwnerProperties.userPrincipalName
                $OwnerObject.subjectId = $Owner.id
                $OwnerObject.displayName = $Owner.displayName
                $OwnerObject.startDateTime = $null
                $OwnerObject.endDateTime = $null                
                $OwnerObject.Type = $Owner.'@odata.type'.split(".")[-1]                 
                $OwnerObject.Status = $null                   
                $OwnerObject.memberType = "Owner"            
                $OwnerObject.isAssignableToRole = $OwnerProperties.isAssignableToRole                                       
                [VOID]$RoleMembers.Add($OwnerObject)
            }
        }
    }
}



if ($RoleMembers | Where-Object{$_.Type -eq "group"})
{
    $MemberTypeGroup = $RoleMembers | Where-Object{$_.Type -eq "group"}
    foreach ($Group in $MemberTypeGroup)
    {

        $GroupID = $Group.subjectid
        if($Group.isAssignableToRole -eq "true")
        {
          $GroupMembers = Get-PAGGroupMembers $Token $Header $TenantID $GroupID
        }
        else {
          $GroupMembers = Get-GroupMembers $Token $Header $TenantID $GroupID
        }
        
        if($GroupMembers)
        {
            foreach($GroupMember in $GroupMembers)
            {
                if($Group.isAssignableToRole -eq "true")
                {
                  $GroupMemberid = $GroupMember.subjectId
                }
                else {
                  $GroupMemberid = $GroupMember.id
                }
                $GroupMemberObject = $(ConvertTo-ObjectArrayListFromPsCustomObject  $Group)
                $GroupMemberProperties = (FindGlobalObject $Token $Header $TenantID $GroupMemberid) | select-object -Property displayName,type,userPrincipalName,isAssignableToRole
                Add-Member -InputObject $GroupMemberObject -MemberType NoteProperty -Name NestedGroupID -value $Group.subjectId
                Add-Member -InputObject $GroupMemberObject -MemberType NoteProperty -Name NestedGroupdisplayName -value $Group.displayName     
                if($(($GroupMemberObject | get-member -MemberType NoteProperty ).name.contains("userPrincipalName")))
                {  
                                         
                    $GroupMemberObject.UserPrincipalName = $GroupMemberProperties.UserPrincipalName                         
                }
                else
                {
                    Add-Member -InputObject $OwnerObject -MemberType NoteProperty -Name UserPrincipalName $GroupMemberProperties.userPrincipalName
                }
                $GroupMemberObject.subjectId = $GroupMemberid
                $GroupMemberObject.displayName = $GroupMemberProperties.displayName
                $GroupMemberObject.startDateTime = $null
                $GroupMemberObject.endDateTime = $null                
                if($GroupMember.'@odata.type' )
                {
                  $GroupMemberObject.Type = $GroupMember.'@odata.type'.split(".")[-1]                                  
                }
                $GroupMemberObject.Status = $null                   
                $GroupMemberObject.memberType = "Nested"
                if($GroupMemberProperties.isAssignableToRole )
                {
                $GroupMemberObject.isAssignableToRole = $GroupMemberProperties.isAssignableToRole                                
                }
                else {
                  $GroupMemberObject.isAssignableToRole = $null
                }
                [VOID]$RoleMembers.Add($GroupMemberObject)
            }
        }
    }
}


#$RoleMembers | sort-object -Property Role |Select-Object -property Role,userPrincipalName,type,status,membertype,NestedGroupdisplayName,subjectId | ft -AutoSize
}



$strFontColor = "#ffffff"

$strHTMLTextCurrent = @"
<!DOCTYPE html>
<html>
<head>
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
body {background-color:#131313;font-family: Arial;color: #ffffff  }

#TopTable table {
  width: 100%;
border-style: none;
}
#TopTable th {
  font-family: Arial;
  color: #00BFFF;
  text-align: center;
  border-style: none;
}
#TopTable td {
  border-style: none;
  text-align: center;
  vertical-align: top
}
#TopTable tr:hover {
background-color: transparent;
}
#TopTable tr:nth-child(even) {
background-color: #000000;
}

table {
    width: 100%;
  }
  
th {
    font-family: Arial;
    color: #00BFFF;
    border-top: 1px solid #ffffff;
    border-left: 1px solid #ffffff;
    border-right: 1px solid #ffffff;
    border-bottom: 1px solid #ffffff;
}
td {
    border-top: 1px solid #ffffff;
    border-left: 1px solid #ffffff;
    border-right: 1px solid #ffffff;
    border-bottom: 1px solid #ffffff;
}
  tr:hover {background-color: coral;}

/* Style the tab */
.tab {
  overflow: hidden;
  border: 1px solid #ccc;
  background-color: #f1f1f1;
}

/* Style the buttons inside the tab */
.tab button {
  background-color: inherit;
  float: left;
  border: none;
  outline: none;
  cursor: pointer;
  padding: 14px 16px;
  transition: 0.3s;
  font-size: 17px;
}

/* Change background color of buttons on hover */
.tab button:hover {
  background-color: #ddd;
}

/* Create an active/current tablink class */
.tab button.active {
  background-color: #ccc;
}

/* Style the tab content */
.tabcontent {
  display: none;
  padding: 6px 12px;
  border: 1px solid #ccc;
  border-top: none;
}
</style>
</head>
<body>
<script>
function openRole(evt, RoleName) {
  var i, tabcontent, tablinks;
  tabcontent = document.getElementsByClassName("tabcontent");
  for (i = 0; i < tabcontent.length; i++) {
    tabcontent[i].style.display = "none";
  }
  tablinks = document.getElementsByClassName("tablinks");
  for (i = 0; i < tablinks.length; i++) {
    tablinks[i].className = tablinks[i].className.replace(" active", "");
  }
  document.getElementById(RoleName).style.display = "block";
  evt.currentTarget.className += " active";
}

/* W3.JS 1.04 April 2019 by w3schools.com */
"use strict";
var w3 = {};
w3.hide = function (sel) {
  w3.hideElements(w3.getElements(sel));
};
w3.hideElements = function (elements) {
  var i, l = elements.length;
  for (i = 0; i < l; i++) {
    w3.hideElement(elements[i]);
  }
};
w3.hideElement = function (element) {
  w3.styleElement(element, "display", "none");
};
w3.show = function (sel, a) {
  var elements = w3.getElements(sel);
  if (a) {w3.hideElements(elements);}
  w3.showElements(elements);
};
w3.showElements = function (elements) {
  var i, l = elements.length;
  for (i = 0; i < l; i++) {
    w3.showElement(elements[i]);
  }
};
w3.showElement = function (element) {
  w3.styleElement(element, "display", "block");
};
w3.addStyle = function (sel, prop, val) {
  w3.styleElements(w3.getElements(sel), prop, val);
};
w3.styleElements = function (elements, prop, val) {
  var i, l = elements.length;
  for (i = 0; i < l; i++) {    
    w3.styleElement(elements[i], prop, val);
  }
};
w3.styleElement = function (element, prop, val) {
  element.style.setProperty(prop, val);
};
w3.toggleShow = function (sel) {
  var i, x = w3.getElements(sel), l = x.length;
  for (i = 0; i < l; i++) {    
    if (x[i].style.display == "none") {
      w3.styleElement(x[i], "display", "block");
    } else {
      w3.styleElement(x[i], "display", "none");
    }
  }
};
w3.addClass = function (sel, name) {
  w3.addClassElements(w3.getElements(sel), name);
};
w3.addClassElements = function (elements, name) {
  var i, l = elements.length;
  for (i = 0; i < l; i++) {
    w3.addClassElement(elements[i], name);
  }
};
w3.addClassElement = function (element, name) {
  var i, arr1, arr2;
  arr1 = element.className.split(" ");
  arr2 = name.split(" ");
  for (i = 0; i < arr2.length; i++) {
    if (arr1.indexOf(arr2[i]) == -1) {element.className += " " + arr2[i];}
  }
};
w3.removeClass = function (sel, name) {
  w3.removeClassElements(w3.getElements(sel), name);
};
w3.removeClassElements = function (elements, name) {
  var i, l = elements.length, arr1, arr2, j;
  for (i = 0; i < l; i++) {
    w3.removeClassElement(elements[i], name);
  }
};
w3.removeClassElement = function (element, name) {
  var i, arr1, arr2;
  arr1 = element.className.split(" ");
  arr2 = name.split(" ");
  for (i = 0; i < arr2.length; i++) {
    while (arr1.indexOf(arr2[i]) > -1) {
      arr1.splice(arr1.indexOf(arr2[i]), 1);     
    }
  }
  element.className = arr1.join(" ");
};
w3.toggleClass = function (sel, c1, c2) {
  w3.toggleClassElements(w3.getElements(sel), c1, c2);
};
w3.toggleClassElements = function (elements, c1, c2) {
  var i, l = elements.length;
  for (i = 0; i < l; i++) {    
    w3.toggleClassElement(elements[i], c1, c2);
  }
};
w3.toggleClassElement = function (element, c1, c2) {
  var t1, t2, t1Arr, t2Arr, j, arr, allPresent;
  t1 = (c1 || "");
  t2 = (c2 || "");
  t1Arr = t1.split(" ");
  t2Arr = t2.split(" ");
  arr = element.className.split(" ");
  if (t2Arr.length == 0) {
    allPresent = true;
    for (j = 0; j < t1Arr.length; j++) {
      if (arr.indexOf(t1Arr[j]) == -1) {allPresent = false;}
    }
    if (allPresent) {
      w3.removeClassElement(element, t1);
    } else {
      w3.addClassElement(element, t1);
    }
  } else {
    allPresent = true;
    for (j = 0; j < t1Arr.length; j++) {
      if (arr.indexOf(t1Arr[j]) == -1) {allPresent = false;}
    }
    if (allPresent) {
      w3.removeClassElement(element, t1);
      w3.addClassElement(element, t2);          
    } else {
      w3.removeClassElement(element, t2);        
      w3.addClassElement(element, t1);
    }
  }
};
w3.getElements = function (id) {
  if (typeof id == "object") {
    return [id];
  } else {
    return document.querySelectorAll(id);
  }
};
w3.filterHTML = function(id, sel, filter) {
  var a, b, c, i, ii, iii, hit;
  a = w3.getElements(id);
  for (i = 0; i < a.length; i++) {
    b = a[i].querySelectorAll(sel);
    for (ii = 0; ii < b.length; ii++) {
      hit = 0;
      if (b[ii].innerText.toUpperCase().indexOf(filter.toUpperCase()) > -1) {
        hit = 1;
      }
      c = b[ii].getElementsByTagName("*");
      for (iii = 0; iii < c.length; iii++) {
        if (c[iii].innerText.toUpperCase().indexOf(filter.toUpperCase()) > -1) {
          hit = 1;
        }
      }
      if (hit == 1) {
        b[ii].style.display = "";
      } else {
        b[ii].style.display = "none";
      }
    }
  }
};
w3.sortHTML = function(id, sel, sortvalue) {
  var a, b, i, ii, y, bytt, v1, v2, cc, j;
  a = w3.getElements(id);
  for (i = 0; i < a.length; i++) {
    for (j = 0; j < 2; j++) {
      cc = 0;
      y = 1;
      while (y == 1) {
        y = 0;
        b = a[i].querySelectorAll(sel);
        for (ii = 0; ii < (b.length - 1); ii++) {
          bytt = 0;
          if (sortvalue) {
            v1 = b[ii].querySelector(sortvalue).innerText;
            v2 = b[ii + 1].querySelector(sortvalue).innerText;
          } else {
            v1 = b[ii].innerText;
            v2 = b[ii + 1].innerText;
          }
          v1 = v1.toLowerCase();
          v2 = v2.toLowerCase();
          if ((j == 0 && (v1 > v2)) || (j == 1 && (v1 < v2))) {
            bytt = 1;
            break;
          }
        }
        if (bytt == 1) {
          b[ii].parentNode.insertBefore(b[ii + 1], b[ii]);
          y = 1;
          cc++;
        }
      }
      if (cc > 0) {break;}
    }
  }
};
w3.slideshow = function (sel, ms, func) {
  var i, ss, x = w3.getElements(sel), l = x.length;
  ss = {};
  ss.current = 1;
  ss.x = x;
  ss.ondisplaychange = func;
  if (!isNaN(ms) || ms == 0) {
    ss.milliseconds = ms;
  } else {
    ss.milliseconds = 1000;
  }
  ss.start = function() {
    ss.display(ss.current)
    if (ss.ondisplaychange) {ss.ondisplaychange();}
    if (ss.milliseconds > 0) {
      window.clearTimeout(ss.timeout);
      ss.timeout = window.setTimeout(ss.next, ss.milliseconds);
    }
  };
  ss.next = function() {
    ss.current += 1;
    if (ss.current > ss.x.length) {ss.current = 1;}
    ss.start();
  };
  ss.previous = function() {
    ss.current -= 1;
    if (ss.current < 1) {ss.current = ss.x.length;}
    ss.start();
  };
  ss.display = function (n) {
    w3.styleElements(ss.x, "display", "none");
    w3.styleElement(ss.x[n - 1], "display", "block");
  }
  ss.start();
  return ss;
};
w3.includeHTML = function(cb) {
  var z, i, elmnt, file, xhttp;
  z = document.getElementsByTagName("*");
  for (i = 0; i < z.length; i++) {
    elmnt = z[i];
    file = elmnt.getAttribute("w3-include-html");
    if (file) {
      xhttp = new XMLHttpRequest();
      xhttp.onreadystatechange = function() {
        if (this.readyState == 4) {
          if (this.status == 200) {elmnt.innerHTML = this.responseText;}
          if (this.status == 404) {elmnt.innerHTML = "Page not found.";}
          elmnt.removeAttribute("w3-include-html");
          w3.includeHTML(cb);
        }
      }      
      xhttp.open("GET", file, true);
      xhttp.send();
      return;
    }
  }
  if (cb) cb();
};
w3.getHttpData = function (file, func) {
  w3.http(file, function () {
    if (this.readyState == 4 && this.status == 200) {
      func(this.responseText);
    }
  });
};
w3.getHttpObject = function (file, func) {
  w3.http(file, function () {
    if (this.readyState == 4 && this.status == 200) {
      func(JSON.parse(this.responseText));
    }
  });
};
w3.displayHttp = function (id, file) {
  w3.http(file, function () {
    if (this.readyState == 4 && this.status == 200) {
      w3.displayObject(id, JSON.parse(this.responseText));
    }
  });
};
w3.http = function (target, readyfunc, xml, method) {
  var httpObj;
  if (!method) {method = "GET"; }
  if (window.XMLHttpRequest) {
    httpObj = new XMLHttpRequest();
  } else if (window.ActiveXObject) {
    httpObj = new ActiveXObject("Microsoft.XMLHTTP");
  }
  if (httpObj) {
    if (readyfunc) {httpObj.onreadystatechange = readyfunc;}
    httpObj.open(method, target, true);
    httpObj.send(xml);
  }
};
w3.getElementsByAttribute = function (x, att) {
  var arr = [], arrCount = -1, i, l, y = x.getElementsByTagName("*"), z = att.toUpperCase();
  l = y.length;
  for (i = -1; i < l; i += 1) {
    if (i == -1) {y[i] = x;}
    if (y[i].getAttribute(z) !== null) {arrCount += 1; arr[arrCount] = y[i];}
  }
  return arr;
};  
w3.dataObject = {},
w3.displayObject = function (id, data) {
  var htmlObj, htmlTemplate, html, arr = [], a, l, rowClone, x, j, i, ii, cc, repeat, repeatObj, repeatX = "";
  htmlObj = document.getElementById(id);
  htmlTemplate = init_template(id, htmlObj);
  html = htmlTemplate.cloneNode(true);
  arr = w3.getElementsByAttribute(html, "w3-repeat");
  l = arr.length;
  for (j = (l - 1); j >= 0; j -= 1) {
    cc = arr[j].getAttribute("w3-repeat").split(" ");
    if (cc.length == 1) {
      repeat = cc[0];
    } else {
      repeatX = cc[0];
      repeat = cc[2];
    }
    arr[j].removeAttribute("w3-repeat");
    repeatObj = data[repeat];
    if (repeatObj && typeof repeatObj == "object" && repeatObj.length != "undefined") {
      i = 0;
      for (x in repeatObj) {
        i += 1;
        rowClone = arr[j];
        rowClone = w3_replace_curly(rowClone, "element", repeatX, repeatObj[x]);
        a = rowClone.attributes;
        for (ii = 0; ii < a.length; ii += 1) {
          a[ii].value = w3_replace_curly(a[ii], "attribute", repeatX, repeatObj[x]).value;
        }
        (i === repeatObj.length) ? arr[j].parentNode.replaceChild(rowClone, arr[j]) : arr[j].parentNode.insertBefore(rowClone, arr[j]);
      }
    } else {
      console.log("w3-repeat must be an array. " + repeat + " is not an array.");
      continue;
    }
  }
  html = w3_replace_curly(html, "element");
  htmlObj.parentNode.replaceChild(html, htmlObj);
  function init_template(id, obj) {
    var template;
    template = obj.cloneNode(true);
    if (w3.dataObject.hasOwnProperty(id)) {return w3.dataObject[id];}
    w3.dataObject[id] = template;
    return template;
  }
  function w3_replace_curly(elmnt, typ, repeatX, x) {
    var value, rowClone, pos1, pos2, originalHTML, lookFor, lookForARR = [], i, cc, r;
    rowClone = elmnt.cloneNode(true);
    pos1 = 0;
    while (pos1 > -1) {
      originalHTML = (typ == "attribute") ? rowClone.value : rowClone.innerHTML;
      pos1 = originalHTML.indexOf("{{", pos1);
      if (pos1 === -1) {break;}
      pos2 = originalHTML.indexOf("}}", pos1 + 1);
      lookFor = originalHTML.substring(pos1 + 2, pos2);
      lookForARR = lookFor.split("||");
      value = undefined;
      for (i = 0; i < lookForARR.length; i += 1) {
        lookForARR[i] = lookForARR[i].replace(/^\s+|\s+$/gm, ''); //trim
        if (x) {value = x[lookForARR[i]];}
        if (value == undefined && data) {value = data[lookForARR[i]];}
        if (value == undefined) {
          cc = lookForARR[i].split(".");
          if (cc[0] == repeatX) {value = x[cc[1]]; }
        }
        if (value == undefined) {
          if (lookForARR[i] == repeatX) {value = x;}
        }
        if (value == undefined) {
          if (lookForARR[i].substr(0, 1) == '"') {
            value = lookForARR[i].replace(/"/g, "");
          } else if (lookForARR[i].substr(0,1) == "'") {
            value = lookForARR[i].replace(/'/g, "");
          }
        }
        if (value != undefined) {break;}
      }
      if (value != undefined) {
        r = "{{" + lookFor + "}}";
        if (typ == "attribute") {
          rowClone.value = rowClone.value.replace(r, value);
        } else {
          w3_replace_html(rowClone, r, value);
        }
      }
      pos1 = pos1 + 1;
    }
    return rowClone;
  }
  function w3_replace_html(a, r, result) {
    var b, l, i, a, x, j;
    if (a.hasAttributes()) {
      b = a.attributes;
      l = b.length;
      for (i = 0; i < l; i += 1) {
        if (b[i].value.indexOf(r) > -1) {b[i].value = b[i].value.replace(r, result);}
      }
    }
    x = a.getElementsByTagName("*");
    l = x.length;
    a.innerHTML = a.innerHTML.replace(r, result);
  }
};
  </script>

<h2>Azure AD Role Assessment</h2>
"@

$strHTMLTextCurrent = $strHTMLTextCurrent + '<table id="TopTable"><tr id="TopTable"><td id="TopTable">'


$strHTMLTextCurrent = $strHTMLTextCurrent + '<h3>Tenant Information</h3>'
#Table with Tenant Information
$TenantData = New-Object PSCustomObject
Add-Member -InputObject $TenantData -MemberType NoteProperty -Name "Tenant ID" -Value $TenantId
$OrganizationData = Get-TenantInformation $Token $Header 
Add-Member -InputObject $TenantData -MemberType NoteProperty -Name "Creation Date" -Value $(($OrganizationData).createdDateTime)
$IntialDomainName = (($OrganizationData).verifiedDomains | Where-Object{$_.isInitial}).Name
Add-Member -InputObject $TenantData -MemberType NoteProperty -Name "Initial Doamin" -Value $IntialDomainName
$DefaultDomainName = (($OrganizationData).verifiedDomains | Where-Object{$_.isDefault}).Name
Add-Member -InputObject $TenantData -MemberType NoteProperty -Name "Default Doamin" -Value $DefaultDomainName
$TenantTable = ($TenantData | ConvertTo-Html -Fragment -As List)
$strHTMLTextCurrent = $strHTMLTextCurrent + $TenantTable + "`n"

$strHTMLTextCurrent = $strHTMLTextCurrent + '</td><td id="TopTable">' + "`n"

### Graph of all role memberships
$RoleSummaryData = New-Object PSCustomObject
$RoleNames = $(($RoleMembers | Select-Object -Property Role -Unique).Role)
Foreach ($RoleName in $RoleNames)
{
    Add-Member -InputObject $RoleSummaryData -MemberType NoteProperty -Name $RoleName -Value $(($RoleMembers | Where-object{$_.Role -eq $RoleName}).count)
}
$strHTMLTextWithGraph = Add-BigDoughnutGraph -Data $RoleSummaryData -GraphTitle "Role Population" -DoughnutTitle "" -arrColors $arrColors 
$strHTMLTextCurrent = $strHTMLTextCurrent + $strHTMLTextWithGraph + "`n"
$strHTMLTextCurrent = $strHTMLTextCurrent + '</td><td id="TopTable">' + "`n"

### Table of all role memberships
$TableAllRole = ($RoleMembers |  Group-Object -Property role | Select-Object -Property @{N="Members";E={$_.count}},@{N="Role";E={$_.name}} | ConvertTo-Html -Fragment)
$TableAllRole = $TableAllRole -replace $TableAllRole[1], ""
$strHTMLTextCurrent = $strHTMLTextCurrent + $TableAllRole + "`n"

$strHTMLTextCurrent = $strHTMLTextCurrent + "</td></tr></table>" + "`n"

$strHTMLTextCurrent = $strHTMLTextCurrent + "<p>Click on the buttons inside the tabbed menu:</p>" + "`n"
$strTab = "<div class='tab'>" + "`n"
Foreach ($RoleName in $(($RoleMembers | Select-Object -Property Role -Unique).Role))
{
$RoleTab = @"
"<button class="tablinks" onclick="openRole(event, '$RoleName')">$RoleName</button>"
"@
    $strTab = $strTab + $RoleTab + "`n"
}
$strTab = $strTab + "</div>" + "`n"

$strHTMLTextCurrent = $strHTMLTextCurrent + $strTab

$strHTMLTextCurrent = $strHTMLTextCurrent + $strHTMLContent
   
$iCount = 0
Foreach ($RoleName in $RoleNames)
{
    #Add Role Name in Header
    $strHTMLGraphs = "<h1><font color='$strFontColor'>$RoleName</font></h1>`n"

    $AssignmentData = New-Object PSCustomObject
    Add-Member -InputObject $AssignmentData -MemberType NoteProperty -Name Eligible -Value $(($RoleMembers | Where-object{($_.Role -eq $RoleName) -and ($_.assignmentState -eq "Eligible")}).count)
    Add-Member -InputObject $AssignmentData -MemberType NoteProperty -Name Active -Value $(($RoleMembers | Where-object{($_.Role -eq $RoleName) -and ($_.assignmentState -eq "Active")}).count)

    $strHTMLTextWithGraph = Add-DoughnutGraph -Data $AssignmentData -GraphTitle "Assignments" -DoughnutTitle "" -arrColors $arrColors 
    $strHTMLGraphs = $strHTMLGraphs + $strHTMLTextWithGraph


    $MemberTypeData = New-Object PSCustomObject
    $MemberTypes = (($RoleMembers | Where-object{($_.Role -eq $RoleName)}) | Select-Object -property membertype)
    $MemberTypesNames = ($MemberTypes | Select-Object -property membertype -Unique).membertype
    foreach($MemberType in $MemberTypesNames)
    {
        Add-Member -InputObject $MemberTypeData -MemberType NoteProperty -Name $MemberType -Value $(($MemberTypes | Where-object{$_.membertype -eq $MemberType}).count)
    }

    $strHTMLTextWithGraph = Add-DoughnutGraph -Data $MemberTypeData -GraphTitle "Memberships" -DoughnutTitle "" -arrColors $arrColors 
    $strHTMLGraphs = $strHTMLGraphs + $strHTMLTextWithGraph

    $TypeData = New-Object PSCustomObject
    $Types = (($RoleMembers | Where-object{($_.Role -eq $RoleName)}) | Select-Object -property type)
    $TypesNames = ($Types | Select-Object -property type -Unique).type
    foreach($Type in $TypesNames)
    {
        Add-Member -InputObject $TypeData -MemberType NoteProperty -Name $Type -Value $(($Types | Where-object{$_.type -eq $Type}).count)
    }

    $strHTMLTextWithGraph = Add-DoughnutGraph -Data $TypeData -GraphTitle "Member object types" -DoughnutTitle "" -arrColors $arrColors 
    $strHTMLGraphs = $strHTMLGraphs + $strHTMLTextWithGraph + "`n"
    
    $UniqueMemberData = New-Object PSCustomObject
    Add-Member -InputObject $UniqueMemberData -MemberType NoteProperty -Name "Unique" -Value $((($RoleMembers | Where-object{($_.Role -eq $RoleName)}) | Select-Object -property subjectid -Unique).count)
    $strHTMLTextWithGraph = Add-DoughnutGraph -Data $UniqueMemberData -GraphTitle "Unique assignments" -DoughnutTitle "" -arrColors $arrColors 
    $strHTMLGraphs = $strHTMLGraphs + $strHTMLTextWithGraph + "`n"

    $strHTMLRoleTable = ""
    $strHTMLRoleTable = ($RoleMembers | Where-object{$_.Role -eq $RoleName} | Select-Object @{Name = "Display Name"; E = {$_.displayName}},@{Name = "UserPrincipalName"; E = {$_.userPrincipalName}},@{Name = "Subject ID"; E = {$_.subjectId}},@{Name = "Type"; E = {$_.type}},@{Name = "Assignment State"; E = {$_.assignmentState}},@{Name = "Member Type"; E = {$_.memberType}},@{Name = "Nested Group"; E = {$_.NestedGroupdisplayName}},@{Name = "Start Time"; E = {$_.startDateTime}},@{Name = "End Time"; E = {if($_.endDateTime){$_.endDateTime}else{"Permanent"}}} | ConvertTo-Html -Fragment).replace("<table>",'<table id="myTable' + $iCount + '">')
    $strHTMLRoleTable = $strHTMLRoleTable.replace("<tr><td>",'<tr class="item"><td>')
    $tableHeaders = ($strHTMLRoleTable | select-string -Pattern "<th>").tostring().split("/")
    $NewtableHeaders = "" 
    $i = 1
    ForEach($tbHead in $tableHeaders)
    {
$NewHeader = @"
<th onclick="w3.sortHTML('#myTable$iCount', '.item', 'td:nth-child($i)')" style="cursor:pointer">
"@        
        $NewtableHeaders = $NewtableHeaders + $tbHead.replace("<th>",$NewHeader) +"/"
        $i++
    }
    $NewtableHeaders = $NewtableHeaders + "</tr>"
    $strHTMLRoleTable = $strHTMLRoleTable -replace  "^<tr><th.+", $NewtableHeaders
    $strHTMLRoleTable = $strHTMLRoleTable + "`n"
$strHTMLRoleContent = @"
<div id="$RoleName" class="tabcontent">
    $strHTMLGraphs
    $strHTMLRoleTable 
</div>
"@    
$strHTMLTextCurrent = $strHTMLTextCurrent + $strHTMLRoleContent + "`n"
$iCount++

}
$strHTMLTextCurrent | Out-File -FilePath $("$HTMLFile") -Force