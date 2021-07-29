#Region Find Windows Start Menu folder location
function Get-ProgDataBasePath {
  Write-Output $env:ALLUSERSPROFILE
}
function Get-AppDataBasePath {
  Write-Output $env:APPDATA
}

function Get-StartMenuDir {
  [CmdletBinding()]
  param (
    [Parameter(ValueFromPipeline)]
    [String]
    $RootDir
  )
    
  $path, $target = $null, $null
    
  try {
    $path = "$RootDir\Microsoft\Windows\Start Menu\Programs"
    $target = Get-Item $path -ErrorAction Stop
    if ($target.PSIsContainer) {
      Write-Host "Found Start Menu folder: $target" -ForegroundColor "green"
      Write-Output $path
    }
    else { Write-Error "Destination is not a folder." }
  }
  catch [System.Management.Automation.ItemNotFoundException] {
    Write-Error "Start Menu folder not found in default location: $path"
    Exit
  }
  catch {
    Write-Error "Unexpected error ocurred type: " $_.Exception.GetType().FullName
    Exit
  }
}
#EndRegion

#Region Reduce parallel arrays
function Join-Lists {
  [CmdletBinding()]
  param (
    [Parameter(ValueFromPipeline)]
    [pscustomobject]
    $Array1,
    [Parameter(ValueFromPipeline)]
    [pscustomobject]
    $Array2
  )

  [pscustomobject[]]$outlist = @()

  ForEach ($item1 in $Array1) {
    ForEach ($item2 in $Array2) {

      if ($item1.BaseName -eq $item2.BaseName) {

        $outlist += [pscustomobject]@{
          BaseName = $item1.BaseName;
          Count1   = $item1.Count;
          Count2   = $item2.Count;
          Total    = $item1.Count + $item2.Count;
          Path1    = $item1.FullPath;
          Path2    = $item2.FullPath;
        }

        Break
      }
    }
  }

  $outlist | Sort-Object -Property BaseName  | Write-Output 
}
#EndRegion

#Region Cleanup helpers
function Remove-Urls {
  [Object[]]$linksToDelete = @()
  foreach ($dir in $args ) {
    foreach ($child in Get-ChildItem $dir -Recurse) {
      if ($child.Extension -eq '.url') {
        $linksToDelete += $child
      }
    }
  }
  if ($linksToDelete.Count) {
    Write-Host "Removing $($linksToDelete.Count) URL$($linksToDelete.Count -gt 1 ? 's' : ''):" -ForegroundColor "yellow"
    ForEach-Object -InputObject $linksToDelete {
      Remove-item $linksToDelete
      Write-Host "`t- $($_.Basename -join "`n`t- ")" -ForegroundColor "green" 
    }
  }
}
function Remove-EmtpyDirs {

  [Object[]]$emptyToDelete = @()
  foreach ($child in $args ) {
    $emptyToDelete += Get-ChildItem $child -Directory -Recurse | Where-Object { $_.GetFileSystemInfos().Count -eq 0 }
  }

  if ($emptyToDelete.Count) {
    Write-Host "Cleanup: Removing $($emptyToDelete.Count) empty director$($emptyToDelete.Count -eq 1 ? "y" : "ies"):" -ForegroundColor "yellow"
    ForEach-Object -InputObject $emptyToDelete { 
      Remove-Item $_;
      Write-Host "`t- $($_.BaseName -join "`n`t- ")" -ForegroundColor "green" 
    }
  }
  else {
    Write-Host "No empty directories found in $child" -ForegroundColor "green"
  }
}
#EndRegion

#Region File moving functions
function Move-SingleItemFolders {
  [CmdletBinding()]
  param (
    [Parameter(Mandatory, ValueFromPipeline)]
    [Object[]]
    $targets,
    [Parameter(Mandatory)]
    [String]
    $path
  )

  $path =

  $targets | Get-ChildItem | Where-Object { $_.Extension -eq ".lnk" } -OutVariable "moved" |
  Move-Item -Destination $path
  
  # Output
  if ($moved) {
    Write-Host "Moved the contents of $($moved.Count) single-item folder$($moved.Count -gt 1 ? 's' : '') to $path :" -ForegroundColor "yellow"
    Write-Host "`t- $($moved.BaseName -join ".lnk,`n`t- ").lnk" -ForegroundColor "green"
  }
  else {
    Write-Host "No single-item folders were found" -ForegroundColor "green"
  }
}

function Move-MultipleItemFolders {
  [CmdletBinding()]
  param (
    [Parameter(Mandatory, ValueFromPipeline)]
    [Object[]]
    $multiItemDirs,
    [Parameter(Mandatory)]
    [String]
    $path
  )

  Write-Host "`nFound $($multiItemDirs.Count) folders with multiple items in`n$($multiItemDirs[0].Parent) :" -ForegroundColor "yellow" 
  $multiItemDirs | ForEach-Object { [PSCustomObject]@{
      Folder = $_.Basename
      Items  = $_.GetFileSystemInfos().Count
    } 
  } | Format-Table
  
  ForEach-Object -InputObject $multiItemDirs {

    $counter = 0
    ForEach ($childDir in $_) {
      $counter++
      $linksInDir = $childDir | Get-ChildItem -Recurse | Where-Object { $_.Extension -eq ".lnk" }

      $title = "Continue? ($($counter)/$($multiItemDirs.Count))"
      $question = "Move $($linksInDir.Count) links in $($childDir.BaseName) ($($linksInDir.BaseName -join ', ')) up to Start Menu?"
      $choices = '&Skip', '&Yes', '&Folder name prefix', 'Custom &prefix'
      
      #-------------------------- Rename ---------------------------
      $prefix = $null
      $confirmedPrefix = $false

      $decision = $linksInDir.Count ? $Host.UI.PromptForChoice($title, $question, $choices, 2) : 0 
      
      if ($decision -eq 0) {
        continue
      }
      if ($decision -eq 2) {
        $prefix = $childDir.BaseName
      }
      if ($decision -eq 3) {
        $prefix = Read-Host "Enter a custom prefix"
      }
      if ($prefix) {

        do {
          $title = "Is this prefix correct?"
          $question = "$($linksInDir | ForEach-Object {
            "`t - " + (($_.BaseName -notmatch "^"+$prefix) ? `
             $_.BaseName + " -> " + $($prefix + " " + $_.BaseName)  + "`n" : `
             "No change: " + $_.BaseName + "`n")
          })"
          $choices = '&No prefix', '&Yes', '&Change', '&Skip folder'
          
          $confirmedPrefix = $Host.UI.PromptForChoice($title, $question, $choices, 1)

          if ($confirmedPrefix -eq 2)
          { $prefix = Read-Host "Enter a different prefix" }

        } while ($confirmedPrefix -eq 2)

        if ($confirmedPrefix -eq 1) {
          $skipped = 0
          $linksInDir | ForEach-Object {
            if ($_.BaseName -notmatch "^" + $prefix) {
              $_ | Rename-Item -NewName $($prefix + " " + $_.BaseName + $_.Extension)
            }
            else {
              $skipped++
            }
          }
          Write-Host "$($linksInDir.Count - $skipped) files renamed" -ForegroundColor "green"
          $linksInDir = $childDir | Get-ChildItem | Where-Object { $_.Extension -eq ".lnk" }
        }
      }
      if ($confirmedPrefix -eq 3)
      { continue }

      #-------------------------- Move files ---------------------------
      if ($linksInDir -and ($decision -eq 1 -or $prefix)) {
        Move-Item -Path $linksInDir -Destination $path
        Write-Host "$($linksInDir.Count) files moved" -ForegroundColor "green"
      }
    }
  }
}
#EndRegion

# Region Warning
$title = 'Run?'
$question = "This script changes files on your computer in your Start Menu folders for all users without making a backup.`nAre you sure you want to proceed?"
$choices = '&Yes', '&No'

$decision = $Host.UI.PromptForChoice($title, $question, $choices, 1)
if ($decision -eq 1) {
  exit
}
# EndRegion


#Region Integration
function Move-FolderContents {
  $list1 = get-ChildItem $args[0] -Directory
  $list2 = get-ChildItem $args[1] -Directory
  $jointLists = Join-Lists $list1 $list2

  ForEach ($path in $args) {
    $uniqueDirs = $path | get-ChildItem -Directory | Where-Object {
      $jointLists.Basename -notcontains $_.BaseName 
    }
   
    if ($uniqueDirs) {
      #-------------------------- Single-item folders ---------------------------
      , ($uniqueDirs | Where-Object { $_.GetFileSystemInfos().Count -eq 1 }) | Move-SingleItemFolders -path $path
      
      #--------------------------- Multi-item folders ---------------------------
      , ($uniqueDirs | Where-Object { $_.GetFileSystemInfos().Count -gt 1 }) | Move-MultipleItemFolders -path $path
    }
  }
}

$ProgDataPath = Get-ProgDataBasePath | Get-StartMenuDir
$AppDataPath = Get-AppDataBasePath | Get-StartMenuDir

Remove-Urls $AppDataPath $ProgDataPath
Move-FolderContents  $AppDataPath $ProgDataPath
Remove-EmtpyDirs $AppDataPath $ProgDataPath
#EndRegion
