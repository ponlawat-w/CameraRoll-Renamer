param ([string]$d = '.')

$localeName = (Get-ItemProperty -Path 'HKCU:\Control Panel\International\' -Name 'LocaleName').LocaleName
$shortDateFormat = (Get-ItemProperty -Path 'HKCU:\Control Panel\International\' -Name 'sShortDate').sShortDate
$shortTimeFormat = (Get-ItemProperty -Path 'HKCU:\Control Panel\International\' -Name 'sShortTime').sShortTime
$dateFormat = "$($shortDateFormat) $($shortTimeFormat)"
$culture = New-Object System.Globalization.CultureInfo($localeName)

function GetFileType { param ([string]$path)
    
    $ext = [System.IO.Path]::GetExtension($path).ToLower()
    
    if ($ext -eq '.jpg' -or $ext -eq '.png') {
        return 'pic'
    } elseif ($ext -eq '.mp4' -or $ext -eq '.mov') {
        return 'vid'
    }
    return $null
}

function GetAttribute { param ($file, [string]$attr)

    $attr = $attr.ToLower()

    0..266 | ForEach-Object {
        $attrName = $folder.GetDetailsOf($folder.Items, $_)
        if ($attrName.ToLower() -eq $attr) {
            $date = $folder.GetDetailsOf($file, $_) -replace [char]8206 -replace [char]8207
            if ($date -eq '' -or $null -eq $date) {
                return ''
            }

            $parsedDate = [DateTime]::ParseExact($date, $dateFormat, $culture)

            return $parsedDate.ToString('yyyy-MM-dd_HH-mm', $culture)
        }
    }
    
    return ''
}

function CreateFilePath { param($date, $number, $format, $extension)
    $numberStr = $number.ToString($format)
    return "$($path.Path)\$($date)_$($numberStr)$($extension)" -replace ' '
}

function Rename { param($file, $date)
    $i = -1
    $destinationPath = ''
    $ext = [System.IO.Path]::GetExtension($file.Path).ToLower()
    do {
        $i++
        $destinationPath = CreateFilePath $date $i '00' $ext
    } while (Test-Path -Path $destinationPath)

    Write-Output "$($file.Path) -> $($destinationPath)"
    Move-Item -Path $file.Path -Destination $destinationPath
}

function RenameNodate { param($file)
    $i = 0
    $destinationPath = ''
    $ext = [System.IO.Path]::GetExtension($file.Path).ToLower()
    do {
        $i = Get-Random -Minimum 0 -Maximum 99999
        $destinationPath = CreateFilePath "9999-99-99_99-99" $i '000000' $ext
    } while (Test-Path -Path $destinationPath)

    Write-Output "$($file.Path) -> $($destinationPath)"
    Move-Item -Path $file.Path -Destination $destinationPath
}

$shell = New-Object -ComObject Shell.Application

$path = Join-Path $pwd $d | Resolve-Path

$folder = $shell.NameSpace($path.Path)

if ($null -eq $folder) {
    Write-Error "Path $d is not found."
    exit
}

foreach ($file in $folder.items()) {
    $fileType = GetFileType $file.Path
    
    if ($fileType -eq 'pic') {
        $date = GetAttribute $file 'Date taken'
    } elseif ($fileType -eq 'vid') {
        $date = GetAttribute $file 'Media created'
    } else {
        continue
    }

    if ($date -eq '') {
        $n = (Get-Random -Minimum 0 -Maximum 99999).ToString('000000')
        $date = "C$(GetAttribute $file 'Date created')_$($n)"
    }

    if ($date -eq '') {
        RenameNodate $file
    } else {
        Rename $file $date
    }
}
