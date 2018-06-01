<#
Installs, updates, and logs info for ruby versions

$rubies_install and $rubies_update variables define versions processed

Updates: RubyGems, bundler, minitest, rake, test-unit
  
Initial code by FeodorFitsner and MSP-Greg
#>

#———————————————————————————————————————————————————————— Install & Version Info

# Path prefix for rubies
$ruby_pre = 'C:\AV\'
# make sure it exists - omit at root of drive...
if( !(Test-Path -Path $ruby_pre) ) { New-Item -Path $ruby_pre -ItemType Directory 1> $null }
$ruby_pre += 'Ruby'

# url for RubyInstaller  (vers 2.3 and lower)
$ri_url  = "https://dl.bintray.com/oneclick/rubyinstaller/"

# url for RubyInstaller2 (vers 2.4 and higher)
$ri2_url = "https://github.com/oneclick/rubyinstaller2/releases/download/"

$r19 = @(
    @{  "version" = "Ruby 1.9.3-p551"
        "suffix"  = "193"
        "download_path" = "ruby-1.9.3-p551-i386-mingw32.7z"
        "devkit_path"   = "DevKit-tdm-32-4.5.2-20111229-1559-sfx.exe"
        "devkit_sufs"   = @("193")
        "install_psych" = "true"
    }
)    

$r20 = @(
    @{  "version" = "Ruby 2.0.0-p648"
        "suffix"  = "200"
        "download_path" = "ruby-2.0.0-p648-i386-mingw32.7z"
        "install_psych" = "true"
    }
    @{  "version" = "Ruby 2.0.0-p648 (x64)"
        "suffix"  = "200-x64"
        "download_path" = "ruby-2.0.0-p648-x64-mingw32.7z"
        "install_psych" = "true"
    }
)

$r21 = @(
    @{  "version" = "Ruby 2.1.9"
        "suffix"  = "21"
        "download_path" = "ruby-2.1.9-i386-mingw32.7z"
    }
    @{  "version" = "Ruby 2.1.9 (x64)"
        "suffix"  = "21-x64"
        "download_path" = "ruby-2.1.9-x64-mingw32.7z"
    }
)

$r22 = @(
    @{
        "version" = "Ruby 2.2.6"
        "suffix"  = "22"
        "download_path" = "ruby-2.2.6-i386-mingw32.7z"
    }
    @{  "version" = "Ruby 2.2.6 (x64)"
        "suffix"  = "22-x64"
        "download_path" = "ruby-2.2.6-x64-mingw32.7z"
    }
)

$r23 = @(
    @{  "version" = "Ruby 2.3.3"
        "suffix"  = "23"
        "download_path" = "ruby-2.3.3-i386-mingw32.7z"
        "devkit_path"   = "DevKit-mingw64-32-4.7.2-20130224-1151-sfx.exe"
        "devkit_sufs"   = @('200', '21', '22', '23')
    }
    @{  "version" = "Ruby 2.3.3 (x64)"
        "suffix"  = "23-x64"
        "download_path" = "ruby-2.3.3-x64-mingw32.7z"
        "devkit_path"   = "DevKit-mingw64-64-4.7.2-20130224-1432-sfx.exe"
        "devkit_sufs"   = @('200-x64', '21-x64', '22-x64', '23-x64')
    }
)

$r24 = @(
    @{  "version"  = "Ruby 2.4.4-1"
        "suffix" = "24"
        "download_path" = "rubyinstaller-2.4.4-1/rubyinstaller-2.4.4-1-x86.7z"
        "devkit_file" = ""
        "devkit_sufs" = @()
    }
    @{  "version" = "Ruby 2.4.4-1 (x64)"
        "suffix"  = "24-x64"
        "download_path" = "rubyinstaller-2.4.4-1/rubyinstaller-2.4.4-1-x64.7z"
        "devkit_file"   = ""
        "devkit_sufs"   = @()
    }
)

$r25 = @(
    @{  "version" = "Ruby 2.5.1-1"
        "suffix"  = "25"
        "download_path" = "rubyinstaller-2.5.1-1/rubyinstaller-2.5.1-1-x86.7z"
        "devkit_file"   = ""
        "devkit_sufs"   = @()
    }
    @{  "version" = "Ruby 2.5.1-1 (x64)"
        "suffix"  = "25-x64"
        "download_path" = "rubyinstaller-2.5.1-1/rubyinstaller-2.5.1-1-x64.7z"
        "devkit_file"   = ""
        "devkit_sufs"   = @()
    }
)

$rubies_install = $r19 + $r20 + $r21 + $r22 + $r23 + $r24 + $r25
$rubies_update  = $r19 + $r20 + $r21 + $r22 + $r23 + $r24 + $r25

#———————————————————————————————————————————————————————————————————————————————

# dash character
$dash = "$([char]0x2015)"

if ($env:SSL_CERT_FILE) {
  $SSL_CERT_FILE_EXISTS = $true
  $SSL_CERT_FILE = $env:SSL_CERT_FILE
} else {
  $SSL_CERT_FILE_EXISTS = $false
}

# download SSL certificates
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
(New-Object Net.WebClient).DownloadFile('https://curl.haxx.se/ca/cacert.pem', "$env:temp\cacert.pem")
$env:SSL_CERT_FILE = "$env:temp\cacert.pem"

# reset at finish
$orig_path = $env:path
# hopefully, a ruby free path
$no_ruby_path = $env:path -ireplace "[^;]+?ruby[^;]+?bin;", ''

function GetUninstallString($productName) {
    $x64items = @(Get-ChildItem "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
    ($x64items + @(Get-ChildItem "HKLM:SOFTWARE\wow6432node\Microsoft\Windows\CurrentVersion\Uninstall") `
        | ForEach-object { Get-ItemProperty Microsoft.PowerShell.Core\Registry::$_ } `
        | Where-Object { $_.DisplayName -and $_.DisplayName -eq $productName } `
        | Select UninstallString).UninstallString
}

function Get-FileNameFromUrl($url) { return $url.Trim('/').split('/')[-1] }

# Install 7z file to $install_path
function Install-7z($ruby, [string]$install_path, [string]$base_url) {
    # create temp directory for all downloads
    $tempPath = Join-Path ([IO.Path]::GetTempPath()) ([IO.Path]::GetRandomFileName())
    New-Item $tempPath -ItemType Directory | Out-Null

    $url = $base_url  + $ruby.download_path
    $distFileName = Get-FileNameFromUrl $ruby.download_path
    $distName = [IO.Path]::GetFileNameWithoutExtension($distFileName)
    $distLocalFileName = (Join-Path $tempPath $distFileName)

    # download archive to a temp
    Write-Host "Downloading $($ruby.version)  ($($ruby.download_path)) from`n            $base_url" -ForegroundColor Gray
    (New-Object Net.WebClient).DownloadFile($url, $distLocalFileName)

    # extract archive to C:\
    Write-Host "Extracting Ruby files..." -ForegroundColor Gray
    7z.exe x $distLocalFileName -y -oC:\ | Out-Null

    # move
    Move-Item -Path "C:\$distName" -Destination $install_path
}

# installer for RubyInstaller based rubies  (vers 2.3 and lower)
function Install-RI($ruby, $install_path) {
    Install-7z $ruby $install_path $ri_url

    # download DevKit
    if ($ruby.devkit_path) {
        $tempPath = Join-Path ([IO.Path]::GetTempPath()) ([IO.Path]::GetRandomFileName())
        New-Item $tempPath -ItemType Directory | Out-Null

        $url = $ri_url + $ruby.devkit_path
        Write-Host "Downloading DevKit ($($ruby.devkit_path)) from`n            $ri_url" -ForegroundColor Gray
        $devKitFileName = Get-FileNameFromUrl $ruby.devkit_path
        $devKitLocalFileName = (Join-Path $tempPath $devKitFileName)
        (New-Object Net.WebClient).DownloadFile($url, $devKitLocalFileName)

        # extract DevKit
        $devKitPath = (Join-Path $install_path 'DevKit')
        Write-Host "Extracting DevKit to $devKitPath..." -ForegroundColor Gray
        7z.exe x $devKitLocalFileName -y -o"$devKitPath" | Out-Null

        # create config.yml
        $configYamlPath = (Join-Path $devKitPath 'config.yml')
        New-Item $configYamlPath -ItemType File | Out-Null

        # write yaml file
        $dk_path = $ruby_pre.replace('\', '/')
        $yaml = "---`n- $dk_path"
        $yaml += ($ruby.devkit_sufs -join "`n- $dk_path" | Out-String )
        Add-Content -Path $configYamlPath -Value $yaml`n

        # install DevKit
        Write-Host "Installing DevKit..." -ForegroundColor Gray
        Push-Location $devKitPath
        ruby.exe dk.rb install
        Pop-Location
    }
    # delete temp path
    if ($tempPath) {
        Write-Host "Cleaning up..." -ForegroundColor Gray
        Remove-Item $tempPath -Force -Recurse
    }

}

# installer for RubyInstaller2 based rubies (vers 2.4 and higher)
function Install-RI2($ruby, $install_path) {
    # uninstall existing
    $rubyUninstallPath = "$install_path\unins000.exe"
    if([IO.File]::Exists($rubyUninstallPath)) {
        Write-Host "Uninstalling previous Ruby 2.4..." -ForegroundColor Gray
        "`"$rubyUninstallPath`" /silent" | out-file "$env:temp\uninstall-ruby.cmd" -Encoding ASCII
        & "$env:temp\uninstall-ruby.cmd"
        del "$env:temp\uninstall-ruby.cmd"
        Start-Sleep -s 5
    }

    Install-7z $ruby $install_path $ri2_url

    # delete html docs
    if (Test-Path $install_path\share\doc\ruby\html) {
        Write-Host "Deleting docs at $install_path\share\doc\ruby\html..." -ForegroundColor Gray
        Remove-Item $install_path\share\doc\ruby\html -Force -Recurse
    }
}

function Update-Ruby($ruby, $install_path) {
    $gem_vers = $(gem --version)
    Write-Host Current RubyGems version is $gem_vers
    
    # RubyGems 3 will not support Ruby < 2.2
    # RubyGems < 2 takes different arguments
    if ($ruby.suffix -lt "22") {
      if ($gem_vers -lt '2.0') {
        Write-Host "gem install rubygems-update --no-rdoc --no-ri --version '~> 2.7'" -ForegroundColor Gray
        gem install rubygems-update --no-rdoc --no-ri --version '~> 2.7'
      } else {
        Write-Host "gem install rubygems-update -N --version '~> 2.7'" -ForegroundColor Gray
        gem install rubygems-update -N --version '~> 2.7'  
      }
      Push-Location $install_path\bin
      ruby.exe update_rubygems 1> $null
      Pop-Location
    } else {
      Write-Host "gem update --system -N" -ForegroundColor Gray
      gem update --system -N -q 1> $null
    }
    Write-Host Done Updating to RubyGems (gem --version) -ForegroundColor Gray

    if ($ruby.install_psych) {
        Write-Host "gem install psych -v 2.2.4 -N" -ForegroundColor Gray
        gem install psych -v 2.2.4 -N
    } elseif ($ruby.update_psych) {
        Write-Host "gem update psych -N" -ForegroundColor Gray
        gem update psych -N
    }

    Write-Host "`ngem update minitest rake test-unit -N" -ForegroundColor Gray
    gem update minitest rake test-unit -N -f
    
    # cleanup old gems
    Write-Host "`ngem cleanup" -ForegroundColor Gray
    gem uninstall rubygems-update -x
    gem cleanup

    # fix "bundler" executable
    Write-Host "add bin\bundler & bin\bundler.bat"
    Copy-Item -Path "$install_path\bin\bundle"     -Destination "$install_path\bin\bundler"     -Force
    Copy-Item -Path "$install_path\bin\bundle.bat" -Destination "$install_path\bin\bundler.bat" -Force

    Write-Host "Done!" -ForegroundColor Green
}

#———————————————————————————————————————————————————————————————————— Main Loops

# install rubies & devkits
foreach ($ruby in $rubies_install) { 
    Write-Host "`n$($dash * 60) Installing $($ruby.version)" -ForegroundColor Cyan
    $install_path = $ruby_pre + $ruby.suffix

    # delete if exists
    if (Test-Path $install_path) {
        Write-Host "Deleting $install_path..." -ForegroundColor Gray
        Remove-Item $install_path -Force -Recurse
    }

    $env:path = "$install_path\bin;$no_ruby_path"
    if ($ruby.suffix -ge '24') { Install-RI2 -ruby $ruby -install_path $install_path
                        } else { Install-RI  -ruby $ruby -install_path $install_path
    }
    Write-Host "Done!" -ForegroundColor Green
}

# update rubies
foreach ($ruby in $rubies_update) { 
    Write-Host "`n$($dash * 60) Updating $($ruby.version)" -ForegroundColor Cyan
    $install_path = $ruby_pre + $ruby.suffix
    $env:path = "$install_path\bin;$no_ruby_path"
    Update-Ruby -ruby $ruby -install_path $install_path
}

# reset SSL_CERT_FILE
if ($SSL_CERT_FILE_EXISTS) { $env:SSL_CERT_FILE = $SSL_CERT_FILE 
                    } else { Remove-Item env:SSL_CERT_FILE }

# install info
$enc = [Console]::OutputEncoding.HeaderName
Push-Location $PSScriptRoot
foreach ($ruby in $rubies_update) {
    Write-Host "`n$($dash * 60) $($ruby.version) Info" -ForegroundColor Cyan
    $install_path = $ruby_pre + $ruby.suffix
    $env:Path = "$install_path\bin;$no_ruby_path"
    ruby.exe install_ruby_info.rb $enc
}
Pop-Location

Write-Host "`n$($dash * 8) Encoding $($dash * 8)" -ForegroundColor Cyan
Write-Host "PS Console  $enc"
iex "ruby.exe -e `"['external','filesystem','internal','locale'].each { |e| puts e.ljust(12) + Encoding.find(e).to_s }`""

$env:path = $orig_path
Write-Host "`n$($dash * 80)`n" -ForegroundColor Cyan
# Add-Path 'C:\Ruby193\bin'
