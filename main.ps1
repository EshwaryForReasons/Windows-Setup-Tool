<#
    Author: Eshwary Mishra
#>

#FUNCTIONS

function insert_blank_line {
    Write-Host ""
}

class menu_info {
    [string]$menu_name
    [program_info[]]$menu_options

    menu_info([string]$_menu_name, [program_info[]]$_menu_options) {
        $this.menu_name = $_menu_name
        $this.menu_options = $_menu_options
    }
}

class program_info {
    [string]$program_id
    [string]$program_name
    [string]$install_overrides = ""
    [bool]$custom = $false
    [bool]$force_machine_installation

    program_info([string]$_program_id, [string]$_program_name) {
        $this.program_id = $_program_id
        $this.program_name = $_program_name
    }

    program_info([string]$_program_id, [string]$_program_name, [string]$_install_overrides) {
        $this.program_id = $_program_id
        $this.program_name = $_program_name
        $this.install_overrides = $_install_overrides
    }

    program_info([string]$_program_id, [string]$_program_name, [bool]$_force_machine_installation, [bool]$_custom) {
        $this.program_id = $_program_id
        $this.program_name = $_program_name
        $this.custom = $_custom
        $this.force_machine_installation = $_force_machine_installation
    }
}

#Confirmation dialog
function ask_confirmation {
    param (
        [string]$message
    )

    Write-Host $message -ForegroundColor Yellow -NoNewline
    $CONFIRMATION_RECEIVED = [Console]::ReadKey($true)
    Write-Host "" -ForegroundColor White

    if ($CONFIRMATION_RECEIVED.Key -eq "y") {
        return $true
    } elseif ($CONFIRMATION_RECEIVED.Key -eq "n") {
        return $false
    } else {
        return ask_confirmation $message
    }
}

#Display programs to install
function display_programs_to_install {
    param (
        [program_info[]]$options,
        [string[]]$selected_options_ids
    )

    insert_blank_line
    Write-Host "Programs to be installed:"
    insert_blank_line
    foreach ($program in $options) {
        if ($selected_options_ids -contains $program.program_id) {
            $program.program_name
        }
    }
    insert_blank_line
}

#Exit program
function exit_program {
    param (
        [string]$message = "Exiting.",
        [bool]$condition = $true
    )

    if ($condition) {
        Write-Host $message
        exit
    }
}

function select_menu {
    param (
        [menu_info[]]$options
    )

    $hovered_index = 0

    while ($true) {
        Clear-Host

        Write-Host "Select a menu to start choosing programs in:"
        insert_blank_line

        for($i = 0; $i -lt $options.Count; $i++) {
            if($hovered_index -eq $i) {
                Write-Host ("{0}. {1} Menu" -f ($i + 1), $options[$i].menu_name) -ForegroundColor Green
            } else {
                Write-Host ("{0}. {1} Menu" -f ($i + 1), $options[$i].menu_name)
            }
        }

        insert_blank_line
        Write-Host "Press [Esc] to exit."
        Write-Host "Press [x] to begin installation."

        $key_pressed = [Console]::ReadKey($true)

        if ($key_pressed.Key -eq "UpArrow") {
            if($hovered_index -gt 0) {
                $hovered_index--
            }
        } elseif ($key_pressed.Key -eq "DownArrow") {
            if($hovered_index -lt ($options.Count - 1)) {
                $hovered_index++
            }
        } elseif ($key_pressed.Key -eq "Enter") {
            return $options[$hovered_index]
        } elseif ($key_pressed.Key -eq "Escape") {
            exit
        } elseif ($key_pressed.Key -eq "x") {
            return "begin_install"
        }
    }
}

#Retrieve options
function retrieve_options {
    param (
        [string]$message,
        [program_info[]]$options,
        [string[]]$previous_selections
    )


    $selected_option_ids = New-Object System.Collections.ArrayList
    $selected_option_ids.AddRange($previous_selections)
    $hovered_index = 0

    while ($true) {
        Clear-Host

        Write-Host $message
        insert_blank_line

        for($i = 0; $i -lt $options.Count; $i++) {
            $marker = if($selected_option_ids.Contains($options[$i].program_id)) {"[*]"} else {"[ ]"}

            if($hovered_index -eq $i) {
                Write-Host ("{0}. {1} {2}" -f ($i + 1), $marker, $options[$i].program_name) -ForegroundColor Green
            } else {
                Write-Host ("{0}. {1} {2}" -f ($i + 1), $marker, $options[$i].program_name)
            }
        }

        insert_blank_line
        Write-Host "Press [a] to select all."
        Write-Host "Press [r] to remove all."
        Write-Host "Press [Esc] to go back without saving."
        Write-Host "Press [x] to confirm selections and go back."

        $key_pressed = [Console]::ReadKey($true)

        if ($key_pressed.Key -eq "UpArrow") {
            if($hovered_index -gt 0) {
                $hovered_index--
            }
        } elseif ($key_pressed.Key -eq "DownArrow") {
            if($hovered_index -lt ($options.Count - 1)) {
                $hovered_index++
            }
        } elseif ($key_pressed.Key -eq "Enter") {
            if($selected_option_ids.Contains($options[$hovered_index].program_id)) {
                $selected_option_ids.Remove($options[$hovered_index].program_id)
            } else {
                $selected_option_ids.Add($options[$hovered_index].program_id)
            }
        } elseif ($key_pressed.Key -eq "a") {
            for($i = 0; $i -lt $options.Count; $i++) {
                if(-not $selected_option_ids.Contains($options[$i].program_id)) {
                    $selected_option_ids.Add($options[$i].program_id)
                }
            }
        } elseif ($key_pressed.Key -eq "r") {
            for($i = 0; $i -lt $options.Count; $i++) {
                if($selected_option_ids.Contains($options[$i].program_id)) {
                    $selected_option_ids.Remove($options[$i].program_id)
                }
            }
        } elseif ($key_pressed.Key -eq "Escape") {
            return [PSCustomObject]@{
                ProgramIDs = $previous_selections
            }
        } elseif ($key_pressed.Key -eq "x") {
            return [PSCustomObject]@{
                ProgramIDs = $selected_option_ids
            }
        }
    }
}

#Custom install functions
function download_file {
    param (
        [string]$item_name,
        [string]$download_url
    )

    Invoke-WebRequest -URI $download_url -OutFile $item_name
}

function install_exe {
    param (
        [string]$installer_name,
        [string]$custom_arguments
    )

    Start-Process -Wait $installer_name `"$custom_arguments`"
}

function install_msi {
    param (
        [string]$installer_name
    )

    $installer_path = Join-Path $PSScriptRoot $installer_name
    Start-Process msiexec.exe -Wait "/I `"$installer_path`" /quiet"
}

function cleanup_installer {
    param (
        [string]$installer_name
    )

    Remove-Item $installer_name
}

function install_program {
    param (
        [switch]$exe,
        [switch]$msi,
        [string]$item_name,
        [string]$download_url,
        [string]$custom_arguments = "/S"
    )

    $installer_name = if ($exe) {$item_name + "Installer.exe"} elseif ($msi) {$item_name + "Installer.msi"}

    download_file $installer_name $download_url
    if ($exe) {
        install_exe $installer_name $custom_arguments
    } elseif ($msi) {
        install_msi $installer_name
    }
    cleanup_installer $installer_name
}

#Winget install
function install_winget_program {
    param (
        [program_info]$program
    )

    $winget_install_command = 'winget install --id $program.program_id -e --source winget '
    if (-not ($program.install_overrides -eq "")) {
        $winget_install_command += '--override "' + $program.install_overrides + '"'
    }
    if ($program.force_machine_installation) {
        $winget_install_command += '--scope machine '
    }
    Invoke-Expression $winget_install_command
}

#MAIN

$GENERAL_OPTIONS = @(
    [program_info]::new("7zip.7zip", "7zip"),
    [program_info]::new("Bitwarden.Bitwarden", "Bitwarden"),
    [program_info]::new("Bitwarden.CLI", "Bitwarden CLI"),
    [program_info]::new("Brave.Brave", "Brave"),
    [program_info]::new("Parsec.Parsec", "Parsec"),
    [program_info]::new("DebaucheeOpenSourceGroup.Barrier", "Barrier"),
    [program_info]::new("Spotify.Spotify", "Spotify")
)

$GAMES_OPTIONS = @(
    [program_info]::new("Valve.Steam", "Steam"),
    [program_info]::new("EpicGames.EpicGamesLauncher", "Epic Games Launcher"),
    [program_info]::new("Mojang.MinecraftLauncher", "Minecraft Launcher")
)

$DEVELOPMENT_OPTIONS = @(
    [program_info]::new("Git.Git", "git"),
    [program_info]::new("OpenJS.NodeJS.LTS", "node.js"),
    [program_info]::new("Terraform", "Terraform", $false, $true),
    [program_info]::new("Microsoft.VisualStudioCode", "Visual Studio Code", $true, $false),
    [program_info]::new("Microsoft.VisualStudio.2022.Professional", "Visual Studio", "
        --add Microsoft.VisualStudio.Workload.NativeGame;includeRecommended 
        --add Microsoft.VisualStudio.Workload.NativeDesktop 
        --add Microsoft.VisualStudio.Component.Debugger.JustInTime 
        --add Microsoft.VisualStudio.Component.SecurityIssueAnalysis 
        --add Microsoft.VisualStudio.Component.VC.DiagnosticTools 
        --add Microsoft.VisualStudio.Component.VC.CMake.Project 
        --remove Microsoft.VisualStudio.Component.IntelliCode --passive --norestart")
)

$WINDOWS_OPTIONS = @(
    [program_info]::new("ChangeWallpaper", "Change windows wallpaper.", $false, $true)
)

$MENUS = @(
    [menu_info]::new("General", $GENERAL_OPTIONS),
    [menu_info]::new("Games", $GAMES_OPTIONS),
    [menu_info]::new("Development", $DEVELOPMENT_OPTIONS),
    [menu_info]::new("Windows", $WINDOWS_OPTIONS)
)

$ALL_OPTIONS = $GENERAL_OPTIONS + $GAMES_OPTIONS + $DEVELOPMENT_OPTIONS + $WINDOWS_OPTIONS
$SELECTED_PROGRAMS_IDS = @()

#Check if winget is installed
try {
    Invoke-Expression "winget --version"
    Write-Host "Winget is installed. Continuing."
} catch {
    Write-Host "Winget is not installed. Retriving winget-cli from github and installing."
    insert_blank_line
    
    download_file "winget.msixbundle" "https://github.com/microsoft/winget-cli/releases/download/v1.4.11071/Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle"
    Start-Job -Name InstallWingetJob -ScriptBlock {Add-AppxPackage "winget.msixbundle"}
    Wait-Job -Name InstallWingetJob
    cleanup_installer "winget.msixbundle"

    Write-Host "Winget-cli successfully isntalled. Continuing."
}

#Place into loop so if user does not confirm selections, we can just redo the whole thing without losing any data
while ($true) {
    while ($true) {
        $selected_menu = select_menu $MENUS
        Write-Host "Selected Menu: " $selected_menu.menu_name

        if ($selected_menu -eq "begin_install") {
            break;
        }

        $SELECTED_PROGRAMS_IDS = (retrieve_options -message ($selected_menu.menu_name + " programs to install:") -options $selected_menu.menu_options -previous_selections $SELECTED_PROGRAMS_IDS).ProgramIDs
    }

    exit_program "No programs to install. Press enter to quit." ($SELECTED_PROGRAMS_IDS.Count -eq 0)
    display_programs_to_install $ALL_OPTIONS $SELECTED_PROGRAMS_IDS
    if (ask_confirmation "Continue with install? (y/n)") {
        break
    }
}

foreach ($program in $ALL_OPTIONS) {
    if ($SELECTED_PROGRAMS_IDS -contains $program.program_id -and -not $program.custom) {
        install_winget_program $program
    } elseif ($SELECTED_PROGRAMS_IDS -contains $program.program_id -and $program.custom) {
        #Handle installing terraform
        if ($program.program_id -eq "Terraform") {
            download_file "Terraform.zip" "https://releases.hashicorp.com/terraform/1.4.6/terraform_1.4.6_windows_amd64.zip"
            Expand-Archive "./Terraform.zip" "C:/Program Files/Terraform" -Force
            #Add Terraform to path
            $existing_path = [Environment]::GetEnvironmentVariable("Path", [EnvironmentVariableTarget]::Machine)
            $new_path = $existing_path + ";C:/Program Files/Terraform"
            [Environment]::SetEnvironmentVariable("Path", $new_path, [EnvironmentVariableTarget]::Machine)
            #Remove Terraform.zip
            cleanup_installer "Terraform.zip"
        }
        #Handle changing windows themes
        if ($program.program_id -eq "ChangeWallpaper") {
            #Download all required files
            download_file "img1.jpg" "https://github.com/EshwaryForReasons/Windows-Setup-Tool/blob/main/img1.jpg?raw=true"
            download_file "img2.jpg" "https://github.com/EshwaryForReasons/Windows-Setup-Tool/blob/main/img2.jpg?raw=true"
            download_file "img3.jpg" "https://github.com/EshwaryForReasons/Windows-Setup-Tool/blob/main/img3.jpg?raw=true"
            download_file "img4.jpg" "https://github.com/EshwaryForReasons/Windows-Setup-Tool/blob/main/img4.jpg?raw=true"
            download_file "windows.theme" "https://raw.githubusercontent.com/EshwaryForReasons/Windows-Setup-Tool/main/windows.theme"

            New-Item ("C:\Users\" + $Env:UserName + "\AppData\Local\Theme") -ItemType Directory
            Move-Item "./img1.jpg" ("C:\Users\" + $Env:UserName + "\AppData\Local\Theme")
            Move-Item "./img2.jpg" ("C:\Users\" + $Env:UserName + "\AppData\Local\Theme")
            Move-Item "./img3.jpg" ("C:\Users\" + $Env:UserName + "\AppData\Local\Theme")
            Move-Item "./img4.jpg" ("C:\Users\" + $Env:UserName + "\AppData\Local\Theme")
            Move-Item "./windows.theme" ("C:\Users\" + $Env:UserName + "\AppData\Local\Theme")

            Start ("C:\Users\" + $Env:UserName + "\AppData\Local\Theme\windows.theme")
        }
    }
}

exit_program "All operations completed. Press enter to quit."