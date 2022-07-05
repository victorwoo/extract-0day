#Requires –Version 6

# prerequests
#
# On macOS:
# crew install p7zip
# https://www.rarlab.com/download.htm
#
# On debian:
# apt-get install genisoimage
# https://stackoverflow.com/questions/286419/how-to-build-a-dmg-mac-os-x-file-on-a-non-mac-platform/16902409
# http://www.tuxarena.com/static/tut_iso_cli.php
# (rar)
# https://www.rarlab.com/download.htm
# apt-get install hfsprogs （未使用）
# https://askubuntu.com/questions/1117461/how-do-i-create-a-dmg-file-on-linux-ubuntu-for-macos

param(
    # 是否需要人工确认每项 0day 软件的处理方式
    [switch]$NeedConfirm = $false
)

# 调试参数
$ErrorActionPreference = 'Inquire'
$DebugPreference = 'Continue'
#$DebugPreference = 'SilentlyContinue'
$VerbosePreference = 'Continue'
#Set-Location '/Users/vichamp/0day'

# 设置路径
$inputPath = Get-Item input
$outputPath = Get-Item output
$tempPath = Join-Path (Get-Item .) 'temp'
$configFile = 'config.csv'

# 压缩文件密码清单
$passwords = @('0daydown')

# 垃圾文件黑名单
$ignoredFiles = @'
__MACOSX
.DS_Store
.AppleDouble
.LSOverride
._*
Thumbs.db
ehthumbs.db
ehthumbs_vista.db
desktop.ini
Desktop.ini
'@.Split("`n") | ForEach-Object { $PSItem.Trim() }

# functions

function GuessConfig([string]$name) {
    Write-Debug $name
    if ($name -cmatch '\W(?i:macos(x?))$|\W(?i:mac(x?))$|\W(?i:mas(x?))$|\W(?i:for mac(?:os)?)') { return 'DMG' }
    if ($name -cmatch 'Topaz') { return 'ISO'}
    if ($name -cmatch 'Antidote') { return 'ISO'}
    if ($name -cmatch 'Red Giant') { return 'ISO'}
    if ($name -cmatch 'Luminar') { return 'ISO'}
    if ($name -cmatch 'JetBrains') { return 'ISO'}
    if ($name -cmatch 'Capture One') { return 'ISO'}
    if ($name -cmatch 'SAPIEN') { return 'ISO'}
    if ($name -match '\W(Multilingual)|(Multilanguage)|(x64)|(x86)|(win)|(Build)|(V?(\d+)(\.\d+)+)\W') {
        if ($name -cmatch '^Adobe ') {
            return 'ISO'
        } else {
            return 'ZIP'
        }
    }
    return 'FLD'
}

# 清除临时目录
function Clear-TempPath {
    if (Test-Path $tempPath) {
        Get-ChildItem -LiteralPath $tempPath | ForEach-Object {
            Remove-Item -Recurse -Force -LiteralPath $PSItem
        }
    } else {
        $null = mkdir $tempPath
    }
}

function Expand-Zip($source, $target) {
    if ($IsMacOS) {
        $allargs = @(
            $source,
            '-d',
            $target
        )
        if ($VerbosePreference -eq 'SilentlyContinue') {
            $allargs += '-q' # 静默
        }
        $InformationPreference = 'Continue'
        &'unzip' $allargs
        $exitCodes = @{
            0='normal; no errors or warnings detected.'
            2='unexpected end of zip file.'
            3='a generic error in the zipfile format was detected.  Pro-cessing may have completed successfully anyway; some bro-ken zipfiles created by other archivers have simple work-arounds.'
            4='zip was unable to allocate memory for one or more buffersduring program initialization.'
            5='a severe error in the zipfile format was detected.   Pro-cessing probably failed immediately.'
            6='entry  too  large  to  be  processed (such as input fileslarger than 2 GB when not using Zip64 or trying  to  readan existing archive that is too large) or entry too largeto be split with zipsplit'
            7='invalid comment format'
            8='zip -T failed or out of memory'
            9='the user aborted zip prematurely with control-C (or simi-lar)'
            10='zip encountered an error while using a temp file'
            11='read or seek error'
            12='zip has nothing to do'
            13='missing or empty zip file'
            14='error writing to a file'
            15='zip was unable to create a file to write to'
            16='bad command line parameters'
            18='zip could not open a specified file to read'
            19='zip  was compiled with options not supported on this sys-tem'
        }
        
        switch ($LASTEXITCODE) {
            0 {
                Write-Output "解压成功"
                $success = $true
            }
            Default {
                if ($exitCodes.ContainsKey($LASTEXITCODE)) {
                    Write-Warning $exitCodes.$LASTEXITCODE
                } else {
                    Write-Warning "未知错误-$LASTEXITCODE"
                }
            }
        } # of switch
    } else {
        $result = Expand-Archive -LiteralPath $source -DestinationPath $target -Force -PassThru -ErrorAction SilentlyContinue
        if ($result) {
            Write-Output '解压成功'
            $success = $true
        } else {
            Write-Warning '解压失败'
        }
    } # of if OS..
}

function Remove-RedundantDir($target) {
    $innerDir = $target
    # 钻取最深的独立目录
    while ((Get-ChildItem $innerDir -Directory).Length -eq 1 -and (Get-ChildItem $innerDir -File).Length -eq 0 -and (!(Get-ChildItem $innerDir)[0].Name.EndsWith('.app'))) {
        $innerDir = (Get-ChildItem $innerDir)[0]
    }
    Write-Debug $innerDir
    if ($innerDir -ne $target) {
        $firstLevelSubdir = (Get-ChildItem $target)[0]
        Write-Debug "降低目录层级 $innerDir -> $target"
        Get-ChildItem -LiteralPath $innerDir.FullName | ForEach-Object {
            Move-Item -LiteralPath $PSItem $target
        }
        Remove-Item -Force -Recurse -LiteralPath $firstLevelSubdir
    }
}

# main script

# 清除临时目录

Clear-TempPath

# 删除百度云生成的临时文件。这些文件会影响解压。正常下载完的目录里不该有这些文件。
Get-ChildItem $inputPath -Include *.baiduyun.downloading -Recurse | Remove-Item
Get-ChildItem $inputPath -Include *.baiduyun.downloading.cfg -Recurse | Remove-Item

# 读取 input 目录
$subdirs = Get-ChildItem $inputPath -Directory

# 交互确定重打包策略
if (Test-Path $configFile) {
    $loadedConfigs = Import-Csv $configFile
} else {
    $loadedConfigs = [PSCustomObject]@{}
}

$configs = @{}
foreach ($subdir in $subdirs) {
    $config =  $loadedConfigs.($subdir.Name)
    if ($config) {
        # 配置文件中有配置项
        $configs[$subdir.Name] = $config
    } else {
        $configs[$subdir.Name] = GuessConfig $subdir.Name
    }
}

$configs | ForEach-Object{ [PSCustomObject]$_ } | Export-CSV -Path $configFile

if ($NeedConfirm) {
    $confirmedConfig = $false
    do {
        $title = 'Is it OK?'
        $message = $configs.GetEnumerator() | Sort-Object -Property Name | Format-Table -AutoSize | Out-String
    
        $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Yes"
        $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "No"
    
        $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
        $result = $host.ui.PromptForChoice($title, $message, $options, 0)
    
        if ($result -eq 0) {
            # 选择了 Yes
            $confirmedConfig = $true
        } else {
            # 选择了 No
            foreach ($subdir in $subdirs) {
                $title = 'How to process'
                $message = $subdir.Name
    
                $folder = New-Object System.Management.Automation.Host.ChoiceDescription "&Folder", "Folder"
                $zip = New-Object System.Management.Automation.Host.ChoiceDescription "&ZIP", "ZIP Archive"
                $dmg = New-Object System.Management.Automation.Host.ChoiceDescription "&DMG", "DMG Archive"
                $iso = New-Object System.Management.Automation.Host.ChoiceDescription "&ISO", "ISO Archive"
    
                $options = [System.Management.Automation.Host.ChoiceDescription[]]($folder, $zip, $dmg, $iso)
    
                $methods = @('FLD', 'ZIP', 'DMG', 'ISO')
                $defaultIndex = $methods.IndexOf($configs[$subdir.Name])
                $result = $host.ui.PromptForChoice($title, $message, $options, $defaultIndex)
                $configs[$subdir.Name] = $methods[$result]
            }
        }
    } while (-not $confirmedConfig)
}

# 写入 csv 文件
$configs | ForEach-Object{ [PSCustomObject]$_ } | Export-CSV -Path $configFile

# 循环处理所有子目录

foreach ($subdir in $subdirs) {
    Write-Output ">> 正在处理 $($subdir.Name)"
    # 分析子目录压缩包格式
    $files = @(Get-ChildItem -File $subdir)

    <#
    # 修正某些 _zip 格式的文件，为 .zip 文件。
    if ($files.Length -eq 1 -and ([string]$files[0]).EndsWith('_zip')) {
        $zipTarget = ([System.IO.FileInfo]$files[0]).FullName.Replace('_zip', '.zip')
        Move-Item $files[0] $zipTarget
    }
    #>
    $extensions = @($files | Group-Object -Property Extension -NoElement | Select-Object -ExpandProperty Name)
    $target = Join-Path $tempPath $subdir.Name '/'
    if (Test-Path $target) {
        Remove-Item -LiteralPath $target -Recurse
    }
    if ($extensions.Count -eq 1) {
        switch ($extensions[0]) {
            '.rar' {
                # 解压 rar 格式子目录
                if ($files.Count -eq 1) {
                    # 只有一个 RAR 文件
                    $source = $files[0]
                } else {
                    # 多个 RAR 文件
                    if (($files | `
                        Where-Object { $_ -cnotmatch '\.part(\d)+\.rar$' } | `
                        Measure-Object).Count) {
                            # 如果有文件不符合 .partX.rar 格式
                            Write-Warning "不支持的格式 - $subdir"
                        } else {
                            # 全部都是分包文件
                            $source = ($files | Sort-Object Name)[0]
                        }
                }

                # rar x -p0daydown "/Users/vichamp/0day/input/Adobe RoboHelp 2019.0.7 Multilingual/RoboHelp.2019.0.7.part1.rar" "/Users/vichamp/0day/output/Adobe RoboHelp 2019.0.7 Multilingual/RoboHelp.2019.0.7/"
                $success = $false
                foreach ($password in $passwords) {
                    $allargs = @(
                        'x',
                        '-c-', # Disable comments show
                        '-o+', # Set the overwrite mode
                        "-p$password",
                        #'-y',
                        $source.FullName,
                        $target
                    )
                    if ($DebugPreference -eq 'SilentlyContinue') {
                        $allargs += '-inul' # Disable all messages
                    }
                    Write-Output "尝试用密码 $password 解压"
                    & 'rar' $allargs

                    $exitCodes = @{
                        0='Successful operation.';
                        1='Warning. Non fatal error(s) occurred.';
                        2='A fatal error occurred.';
                        3='Invalid checksum. Data is damaged.';
                        4='Attempt to modify a locked archive.';
                        5='Write error.';
                        6='File open error.';
                        7='Wrong command line option.';
                        8='Not enough memory.';
                        9='File create error.';
                        10='No files matching the specified mask and options were found.';
                        11='Wrong password.';
                        255='User break.';
                    }
                    switch ($LASTEXITCODE) {
                        0 {
                            # 解压成功
                            Write-Output "解压成功"
                            $success = $true
                        }
                        11 {
                            # 密码错
                            Write-Output '密码错'
                        }
                        Default {
                            if ($exitCodes.ContainsKey($LASTEXITCODE)) {
                                Write-Warning $exitCodes.$LASTEXITCODE
                            } else {
                                Write-Warning "未知错误-$LASTEXITCODE"
                            }
                        }
                    }
                    if ($success) { 
                        # 修正某些 _zip 格式的文件，为 .zip 文件。
                        $files = Get-ChildItem -File -Recurse temp/
                        if ($files.Count -eq 1 -and ([string]$files[0]).EndsWith('_zip')) {
                            Write-Output "解开 $($files[0])"
                            $zipTarget = ([System.IO.FileInfo]$files[0]).FullName.Replace('_zip', '.zip')
                            Move-Item -LiteralPath $files[0] $zipTarget
                            
                            # 去掉 ".zip" 结尾
                            $zipFolder = $zipTarget.Substring(0, $zipTarget.Length - 4) + '/'
                            
                            Write-Output "准备解压 $zipTarget"
                            if (Expand-Zip $zipTarget $zipFolder) {
                                # 解压成功，删除临时的 ZIP 文件。
                                Write-Output "准备移动 $zipTarget"
                                Remove-Item -LiteralPath $zipTarget
                                Write-Output "移动完毕 $zipTarget"
                            }
                        }
                        break
                    }
                }
             }
            '.zip' {
                # 解压 zip 格式子目录
                if ($files.Count -eq 1) {
                    # 只有一个 ZIP 文件
                    $source = $files[0]
                    $success = Expand-Zip $source.FullName $target
                } else {
                    Write-Warning "不支持多个 ZIP 文件 - $($subdir)"
                    $success = $false
                }
             }
             { '.dmg', '.iso' -contains $PSItem } {
                # 只含 DMG/ISO 镜像，直接移动文件夹
                Write-Output "只包含镜像文件 $($extensions[0])，直接移动文件夹"
                $source = $subdir
                Move-Item -LiteralPath $source.FullName $target
                $success = $true
             }
            Default {
                # 只包含未知类型文件，直接移动文件夹
                Write-Debug "只包含未知类型文件 $($extensions[0])，直接移动文件夹"
                $source = $subdir
                Move-Item -LiteralPath $source.FullName $target
                $success = $true
            }
        }
    } else {
        # 不是一种扩展名，直接移动内容
        if ($extensions.Count -eq 0) {
            if (Get-ChildItem $subdir) {
                # 有内容（只有子文件夹）
                Write-Debug "只有子目录，没有文件，直接移动文件夹"
                Move-Item -LiteralPath $subdir.FullName $target
                $success = $true
            } else {
                # 无内容（空文件夹）
                Write-Warning "空文件夹"
            }
        } else {
            Write-Debug "有多种扩展名：$($extensions)，直接移动文件夹"
            Move-Item -LiteralPath $subdir.FullName $target
            $success = $true
        }
    }

    if (-not $success) {
        Continue
    }

    # 清除黑名单文件
    if (Test-Path $target) {
        foreach ($ignoredFile in $ignoredFiles){
            #$path = Join-Path $target '*'
            #Remove-Item $path -Recurse -Force -Include $ignoredFile
            Get-ChildItem $target -Recurse | Where-Object { $PSItem.Name -like $ignoredFile } | ForEach-Object {
                Write-Debug "删除 $($PSItem.FullName)"
                Remove-Item -LiteralPath $PSItem.FullName -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }

    # 降低目录层级
    if (Test-Path $target) {
        Remove-RedundantDir $target

        # 根据重打包策略重打包
        switch ($configs[$subdir.Name]) {
            'ZIP' {
                $zipPath = (Join-Path $outputPath $subdir.Name) + '.zip'
                if ((@(Get-ChildItem $target)).Count -eq 1 -and (Get-ChildItem $target).Extension -eq '.zip') {
                    $singleFile = Get-ChildItem $target
                    Move-Item -LiteralPath $singleFile $zipPath
                    Write-Output "移动单个 zip 文件到 $zipPath"
                }  else {
                    if ($IsMacOS) {
                        $allargs = @(
                            '-r', # recurse into directories
                            '-m', # move into zipfile (delete OS files)
                            $zipPath,
                            '*'
                        )
                        if ($VerbosePreference -eq 'SilentlyContinue') {
                            $allargs += '-q' # 静默
                        }
                        Push-Location
                        Set-Location $target
                        & 'zip' $allargs
                        Pop-Location
                        $exitCodes = @{
                            0='normal; no errors or warnings detected.'
                            2='unexpected end of zip file.'
                            3='a generic error in the zipfile format was detected.  Pro-cessing may have completed successfully anyway; some bro-ken zipfiles created by other archivers have simple work-arounds.'
                            4='zip was unable to allocate memory for one or more buffersduring program initialization.'
                            5='a severe error in the zipfile format was detected.   Pro-cessing probably failed immediately.'
                            6='entry  too  large  to  be  processed (such as input fileslarger than 2 GB when not using Zip64 or trying  to  readan existing archive that is too large) or entry too largeto be split with zipsplit'
                            7='invalid comment format'
                            8='zip -T failed or out of memory'
                            9='the user aborted zip prematurely with control-C (or simi-lar)'
                            10='zip encountered an error while using a temp file'
                            11='read or seek error'
                            12='zip has nothing to do'
                            13='missing or empty zip file'
                            14='error writing to a file'
                            15='zip was unable to create a file to write to'
                            16='bad command line parameters'
                            18='zip could not open a specified file to read'
                            19='zip  was compiled with options not supported on this sys-tem'
                        }
                        
                        switch ($LASTEXITCODE) {
                            0 {
                                Write-Output "压缩成功"
                                $success = $true
                            }
                            Default {
                                if ($exitCodes.ContainsKey($LASTEXITCODE)) {
                                    Write-Warning $exitCodes.$LASTEXITCODE
                                } else {
                                    Write-Warning "未知错误-$LASTEXITCODE"
                                }
                            }
                        }
                    } else {
                        $result = Compress-Archive -PassThru -Force -LiteralPath $target -DestinationPath $zipPath
                        if ($result) {
                            Write-Output "生成 $zipPath 成功"
                            $success = $true
                        } else {
                            Write-Warning "生成 $zipPath 失败"
                            $success = $false
                        }
                    }
                } # of if 单个 .zip
            }
            'DMG' { 
                $dmgPath = (Join-Path $outputPath $subdir.Name) + '.dmg'
                if ((@(Get-ChildItem $target)).Count -eq 1 -and (Get-ChildItem $target).Extension -eq '.dmg') {
                    $singleFile = Get-ChildItem $target
                    Move-Item -LiteralPath $singleFile $dmgPath
                    Write-Output "移动单个 dmg 文件到 $dmgPath"
                }  else {
                    if ($IsMacOS) {
                        # hdiutil create -format UDZO -srcfolder "/Users/vichamp/0day/.temp/multi-levels 1.2" "/Users/vichamp/0day/.temp/multi-levels 1.2.dmg"
                        $volname = $subdir.Name
                        if ($volname.Length -gt 11) {
                            $volname = $volname.Substring(0, 10) + '_'
                        }
                        $allargs = @(
                            'create',
                            '-volname',
                            $volname,
                            '-format',
                            'UDZO',
                            '-ov' # 覆盖
                            '-srcfolder'
                            $target,
                            $dmgPath
                        )
                        if ($DebugPreference -eq 'SilentlyContinue') {
                            $allargs += '-quiet' # 静默
                        }
                        & 'hdiutil' $allargs
                    }

                    if ($IsLinux) {
                        # genisoimage -V progname -D -R -apple -no-pad -o progname.dmg dmgdir
                        $allargs = @(
                            '-V',
                            $subdir.Name,
                            '-D',
                            '-R',
                            '-apple',
                            '-no-pad',
                            '-o',
                            $dmgPath,
                            $target
                        )
                        & 'genisoimage' $allargs
                    }

                    if ($LASTEXITCODE) {
                        Write-Warning "生成 $dmgPath 失败"
                        $success = $false
                    } else {
                        Write-Output "生成 $dmgPath 成功"
                        $success = $true
                    }
                }
            }
            'ISO' {
                $isoPath = (Join-Path $outputPath $subdir.Name) + '.iso'
                if ((@(Get-ChildItem $target)).Count -eq 1 -and (Get-ChildItem $target).Extension -eq '.iso') {
                    $singleFile = Get-ChildItem $target
                    Move-Item -LiteralPath $singleFile $isoPath
                    Write-Output "移动单个 iso 文件到 $isoPath"
                }  else {

                    $isoPath = (Join-Path $outputPath $subdir.Name) + '.iso'

                    # mac:
                    if ($IsMacOS) {
                        # hdiutil makehybrid -iso -joliet -o git-local-2017.iso git-local-2017
                        $allargs = @(
                            'makehybrid',
                            '-iso',
                            '-joliet',
                            '-ov' # 覆盖
                            '-o'
                            $isoPath,
                            $target
                        )
                        if ($DebugPreference -eq 'SilentlyContinue') {
                            $allargs += '-quiet' # 静默
                        }
                        & 'hdiutil' $allargs
                    }

                    # Linux (debian):
                    if ($IsLinux) {
                        # genisoimage -o output_image.iso directory_name
                        $allargs = @(
                            '-o',
                            $isoPath,
                            $target
                        )
                        & 'genisoimage' $allargs
                    }

                    if ($LASTEXITCODE) {
                        Write-Warning "生成 $isoPath 失败"
                        $success = $false
                    } else {
                        Write-Output "生成 $isoPath 成功"
                        $success = $true
                    }
                }
            }
            'FLD' { 
                $finalPath = Join-Path $outputPath $subdir.Name
                if (Test-Path $finalPath) {
                    Remove-Item -Recurse -Force $finalPath
                }
                Move-Item -LiteralPath $target $outputPath
                Write-Output "生成文件夹 $finalPath 成功"
                $success = $true
            }
            Default {
                Write-Error "未知的配置 - $($configs[$subdir.Name])"
            }
        }

        if ($success -and (Test-Path $subdir)) {
            Get-ChildItem -LiteralPath $subdir -Recurse | Remove-Item -Recurse -Force
            Remove-Item -Recurse -Force -LiteralPath $subdir
        }

        if ($success -and (Test-Path $target)) {
            Get-ChildItem -LiteralPath $target -Recurse | Remove-Item -Recurse -Force
            Remove-Item -Recurse -Force -LiteralPath $target
        }
    }
}

# 清理现场
# Remove-Item $tempPath
# Remove-Item $configFile