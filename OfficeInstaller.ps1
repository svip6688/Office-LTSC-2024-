# =========================================
# Office LTSC 2024 自动下载/安装助手 (PowerShell)
# 完整版：自动识别路径 + 下载 + 安装 + 日志
# =========================================

# ------------------------
# 自动识别脚本所在目录
# ------------------------
$BaseDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
Set-Location $BaseDir

# ------------------------
# 自动搜索 setup.exe 和 configuration.xml（递归搜索）
# ------------------------
$Setup = Get-ChildItem -Path $BaseDir -Recurse -Filter "setup.exe" | Select-Object -First 1
$Config = Get-ChildItem -Path $BaseDir -Recurse -Filter "configuration.xml" | Select-Object -First 1

if (-not $Setup) {
    Write-Host "找不到 setup.exe"
    Read-Host "按 Enter 键退出"
    exit
}
if (-not $Config) {
    Write-Host "找不到 configuration.xml"
    Read-Host "按 Enter 键退出"
    exit
}

$SetupPath = $Setup.FullName
$ConfigPath = $Config.FullName

Write-Host "找到 setup.exe： $SetupPath"
Write-Host "找到 configuration.xml： $ConfigPath"

# ------------------------
# 切换到 setup.exe 所在目录（关键，模拟手动 cd）
# ------------------------
$SetupDir = Split-Path -Parent $SetupPath
Set-Location $SetupDir

# ------------------------
# 日志文件
# ------------------------
function Write-Log {
    param([string]$Message)

    # 如果日志不存在，创建并加 BOM 让记事本显示中文
    if (-not (Test-Path $Log)) {
        "`uFEFF" | Out-File -FilePath $Log -Encoding utf8
    }

    # 追加日志，用 UTF-8 编码
    $Message | Out-File -FilePath $Log -Append -Encoding utf8
}

# ------------------------
# 菜单
# ------------------------
function Show-Menu {
    Clear-Host
    Write-Host "======================================"
    Write-Host " Office LTSC 2024 自动下载/安装助手"
    Write-Host "======================================"
    Write-Host "1. 下载 Office（联网）"
    Write-Host "2. 安装 Office（离线或联网）"
    Write-Host "3. 退出"
    Write-Host
    $choice = Read-Host "请输入数字"
    switch ($choice) {
        "1" { Download-Office }
        "2" { Install-Office }
        "3" { Exit-Script }
        default {
            Write-Host "输入无效。按 Enter 返回菜单..."
            Read-Host
            Show-Menu
        }
    }
}

# ------------------------
# 下载 Office
# ------------------------
function Download-Office {
    Write-Host "开始下载 Office..."
    Write-Log "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] 开始下载 Office..."

    # 调用 setup.exe 下载
    Start-Process -FilePath ".\setup.exe" -ArgumentList "/download `"$ConfigPath`"" -Wait -NoNewWindow

    Write-Log "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] 下载完成 √"
    Write-Host "下载完成 √"
    Read-Host "按 Enter 返回菜单"
    Show-Menu
}

# ------------------------
# 安装 Office
# ------------------------
function Install-Office {
    $ok = Read-Host "你选择安装 Office。输入 y 执行安装"
    if ($ok -eq "y") {
        Write-Host "开始安装..."
        Write-Log "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] 开始安装 Office..."

        # 自动获取 setup.exe 所在目录
        $SetupDir = Split-Path -Parent $SetupPath
        Set-Location $SetupDir  # 切换目录，避免 0-2048

        # 调用 setup.exe 安装（使用 .\ 前缀）
        Start-Process -FilePath ".\setup.exe" -ArgumentList "/configure `"$ConfigPath`"" -Wait -NoNewWindow

        Write-Log "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] 安装完成 √"
        Write-Host "安装完成 √"
    } else {
        Write-Host "安装已取消。"
        Write-Log "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] 安装已取消。"
    }
    Read-Host "按 Enter 返回菜单"
    Show-Menu
}


# ------------------------
# 退出脚本
# ------------------------
function Exit-Script {
    Write-Host "已退出."
    Read-Host "按 Enter 键关闭窗口"
    exit
}

# ------------------------
# 启动菜单
# ------------------------
Show-Menu
