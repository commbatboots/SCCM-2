# PowerShell 5.1
# Flashplayer
$Global:PathRoot = "C:\Downloads\"

function Get-CurrentFlashPlayer ()
{
    $DownloadFolder = $Global:Pathroot
    $Download_ActiveX = $true   # IE
    $Download_PPAPI   = $true   # Chrome Pepper
    $Download_NPAPI  = $true   # Firefox
    # Example of link format:
    # https://fpdownload.macromedia.com/pub/flashplayer/pdc/24.0.0.194/install_flash_player_24_active_x.msi
    # https://fpdownload.macromedia.com/pub/flashplayer/pdc/24.0.0.194/install_flash_player_24_plugin.msi
    # https://fpdownload.macromedia.com/pub/flashplayer/pdc/24.0.0.194/install_flash_player_24_ppapi.msi


    $fplayer_mainlink = 'https://fpdownload.macromedia.com/pub/flashplayer/pdc/'


    # Code Barrowed from https://www.reddit.com/r/PowerShell/comments/3tgr2m/get_current_versions_of_adobe_products/
    $Uri = "http://fpdownload2.macromedia.com/pub/flashplayer/update/current/sau"	
    $FlashMajorVersion = Invoke-WebRequest -Uri "$Uri/currentmajor.xml"
    # $FlashMajorVersion = Invoke-WebRequest -Uri "$($Uri)/currentmajor.xml"
    $MajorVersion = ([xml]$FlashMajorVersion.content).version.player.major  # Corrected bug on this line to fully parase the XML
    # $CurrentFlashVersion = Invoke-WebRequest -Uri "$uri/$MajorVersion/xml/version.xml"
    [xml]$CurrentFlashVersion = Invoke-WebRequest -Uri "$($Uri)/$($MajorVersion)/xml/version.xml"

    # Active X Version - This Auto Updates on Windows 10 with Windows Update
    $ActiveXVersion = ($CurrentFlashVersion.version.ActiveX.major) + "." + ($CurrentFlashVersion.version.ActiveX.minor) + "." + ($CurrentFlashVersion.version.ActiveX.buildMajor) + "." + ($CurrentFlashVersion.version.ActiveX.buildMinor)

    # Plugin Version (NPAPI)
    $PluginVersion = ($CurrentFlashVersion.version.Plugin.major) + "." + ($CurrentFlashVersion.version.Plugin.minor) + "." + ($CurrentFlashVersion.version.Plugin.buildMajor) + "." + ($CurrentFlashVersion.version.Plugin.buildMinor)

    # Chrome (PPAPI) - This Auto Updates via Chrome
    $PepperVersion = ($CurrentFlashVersion.version.Pepper.major) + "." + ($CurrentFlashVersion.version.Pepper.minor) + "." + ($CurrentFlashVersion.version.Pepper.buildMajor) + "." + ($CurrentFlashVersion.version.Pepper.buildMinor)

    # Download Versions
    # ActiveX
    if ($Download_ActiveX)
    {
        $ActiveXMSI = "install_flash_player_" + $MajorVersion + "_active_x.msi"

        # URL to download Active X
        $ActiveX = $fplayer_mainlink + $ActiveXVersion + "/" + $ActiveXMSI

        # Download Path \ Version \ Active X Plugin
        $ActiveXDownloadFolder = $DownloadFolder + "Adobe\" + "Flash\" + "Active X Plugin\" + $ActiveXVersion + "\x86\" + "Files\"

        # Full Download Path
        $ActiveXDownloadPath = $ActiveXDownloadFolder + "\" + $ActiveXMSI
    
        # Create the full folder path if it doesnt exist
        if (!(Test-Path -Path $ActiveXDownloadFolder))
        {
            New-Item -ItemType Directory -Path $ActiveXDownloadFolder
        }

        # Download the Active X Plugin
        Invoke-WebRequest -Uri $ActiveX -Outfile $ActiveXDownloadPath
    }

    # NPAPI / FireFox
    if ($Download_NPAPI)
    {
        $NPAPIMSI = "install_flash_player_" + $MajorVersion + "_plugin.msi"

        # URL to download Active X
        $NPAPI = $fplayer_mainlink + $PluginVersion + "/" + $NPAPIMSI
        # Download Path \ Version \ Active X Plugin
        # $PPAPIDownloadFolder = $DownloadFolder + $PPAPIVersion + "\NPAPI Plugin"
        $NPAPIDownloadFolder = $DownloadFolder + "Adobe\" + "Flash\" + "NPAPI Plugin\" + $PluginVersion + "\x86\" + "Files\"
        
        # Full Download Path
        $NPAPIDownloadPath = $NPAPIDownloadFolder + "\" + $NPAPIMSI
    
        # Create the full folder path if it doesnt exist
        if (!(Test-Path -Path $NPAPIDownloadFolder))
        {
            New-Item -ItemType Directory -Path $NPAPIDownloadFolder
        }
    
        Invoke-WebRequest -Uri $NPAPI -Outfile $NPAPIDownloadPath
    }

    # Pepper / Chrome
    if ($Download_PPAPI)
    {
        $PPAPIMSI = "install_flash_player_" + $MajorVersion + "_ppapi.msi"

        # URL to download Active X
        $PPAPI = $fplayer_mainlink + $PepperVersion + "/" + $PPAPIMSI
        # Download Path \ Version \ Active X Plugin
        # $PPAPIDownloadFolder = $DownloadFolder + $PPAPIVersion + "\NPAPI Plugin"
        $PPAPIDownloadFolder = $DownloadFolder + "Adobe\" + "Flash\" + "Pepper Plugin\" + $PepperVersion + "\x86\" + "Files\"
        
        # Full Download Path
        $PPAPIDownloadPath = $PPAPIDownloadFolder + "\" + $PPAPIMSI
    
        # Create the full folder path if it doesnt exist
        if (!(Test-Path -Path $PPAPIDownloadFolder))
        {
            New-Item -ItemType Directory -Path $PPAPIDownloadFolder
        }
    
        Invoke-WebRequest -Uri $PPAPI -Outfile $PPAPIDownloadPath
    }

}

# Note - Setup to only download and prepare Active X - BWT
Get-CurrentFlashPlayer