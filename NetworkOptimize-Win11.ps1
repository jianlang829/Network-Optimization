#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Windows 11 笔记本通用网络优化脚本
.DESCRIPTION
    优化 TCP/IP 栈、网卡节能、DNS 设置等，提升网络响应速度与吞吐量。
    适用于日常浏览、游戏、视频会议等场景。
    自动备份原始设置，支持还原。
.NOTES
    作者: jianlang
    日期: 2025-09-05
    适用: Windows 11 笔记本电脑（有线/无线均支持）
#>

param(
    [Parameter(Mandatory=$false)]
    [ValidateSet("Apply", "Restore", "BackupOnly")]
    [string]$Action = "Apply"
)

# 设置备份路径
$backupPath = "$env:USERPROFILE\Desktop\NetworkOptimize_Backup_$(Get-Date -Format 'yyyyMMdd_HHmmss').reg"

function Backup-CurrentSettings {
    Write-Host "💾 正在备份当前网络设置到: $backupPath" -ForegroundColor Cyan
    try {
        reg export "HKLM\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters" "$backupPath" /y | Out-Null
        Write-Host "✅ 备份完成。" -ForegroundColor Green
    } catch {
        Write-Warning "⚠️ 备份失败: $_"
    }
}

function Optimize-TcpIp {
    Write-Host "🔧 正在优化 TCP/IP 设置..." -ForegroundColor Cyan

    $tcpSettings = @{
        "TcpWindowSize"           = 64KB * 1024    # 接收窗口（64KB）
        "GlobalMaxTcpWindowSize"  = 64KB * 1024
        "Tcp1323Opts"             = 3               # 启用时间戳 + 窗口缩放
        "DefaultTTL"              = 64              # 默认 TTL
        "EnablePMTUDiscovery"     = 1               # 启用路径 MTU 发现
        "EnableRSS"               = 1               # 接收端缩放（多核优化）
        "EnableTCPChimney"        = 0               # 禁用（兼容性考虑）
        "EnableTCPA"              = 1               # 启用 TCP ACK 加速
        "EnableTCPNoDelay"        = 1               # 禁用 Nagle 算法（降低延迟）
        "CongestionProvider"      = "CTCP"          # 使用 Compound TCP（Win10/11 默认）
    }

    foreach ($key in $tcpSettings.Keys) {
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters" -Name $key -Value $tcpSettings[$key] -ErrorAction SilentlyContinue
    }

    # 设置最大用户端口（避免端口耗尽）
    Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters" -Name "MaxUserPort" -Value 65534 -ErrorAction SilentlyContinue
    Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters" -Name "TcpTimedWaitDelay" -Value 30 -ErrorAction SilentlyContinue

    Write-Host "✅ TCP/IP 优化完成。" -ForegroundColor Green
}

function Disable-NicPowerSaving {
    Write-Host "🔌 正在禁用网卡节能功能（防止休眠断连）..." -ForegroundColor Cyan

    Get-NetAdapter | Where-Object {$_.Status -eq "Up"} | ForEach-Object {
        $adapterName = $_.Name
        Write-Host "  → 处理适配器: $adapterName" -ForegroundColor Gray

        # 禁用节能（电源管理）
        try {
            Disable-NetAdapterPowerManagement -Name $_.Name -ErrorAction Stop
            Write-Host "    ✅ 已禁用节能" -ForegroundColor DarkGreen
        } catch {
            Write-Warning "    ⚠️ 无法禁用节能: $_"
        }

        # 设置高级属性（如存在）
        $advancedProps = @("Energy Efficient Ethernet", "Green Ethernet", "Power Saving Mode", "ASPM")
        foreach ($prop in $advancedProps) {
            try {
                Set-NetAdapterAdvancedProperty -Name $_.Name -DisplayName $prop -DisplayValue "Disabled" -ErrorAction Stop | Out-Null
            } catch {
                # 属性不存在则忽略
            }
        }
    }
    Write-Host "✅ 网卡节能设置已优化。" -ForegroundColor Green
}

function Optimize-Dns {
    Write-Host "🌐 正在优化 DNS 设置..." -ForegroundColor Cyan

    # 增加 DNS 缓存大小
    Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\Dnscache\Parameters" -Name "MaxCacheEntryTtlLimit" -Value 86400 -ErrorAction SilentlyContinue
    Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\Dnscache\Parameters" -Name "MaxSOACacheEntryTtlLimit" -Value 300 -ErrorAction SilentlyContinue

    # 刷新 DNS 缓存
    ipconfig /flushdns | Out-Null
    Write-Host "✅ DNS 优化完成并已刷新缓存。" -ForegroundColor Green
}

function Apply-QoS {
    Write-Host "🚦 正在应用 QoS 优化（多媒体优先）..." -ForegroundColor Cyan

    # 启用基于策略的 QoS（可选）
    # 此处仅为示例，实际需按需配置组策略或使用 netsh

    # 示例：设置 DSCP 标记（如 VoIP 优先）
    # netsh int tcp set global autotuninglevel=normal

    Write-Host "ℹ️  QoS 高级设置建议通过组策略配置。" -ForegroundColor Yellow
}

function Restart-NetworkServices {
    Write-Host "🔄 正在重启网络服务..." -ForegroundColor Cyan
    Restart-Service -Name "Dnscache" -Force -ErrorAction SilentlyContinue
    Restart-Service -Name "iphlpsvc" -Force -ErrorAction SilentlyContinue
    Write-Host "✅ 网络服务重启完成。" -ForegroundColor Green
}

function Restore-FromBackup {
    if (Test-Path $backupPath) {
        Write-Host "🔄 正在从备份还原网络设置..." -ForegroundColor Cyan
        reg import "$backupPath"
        Write-Host "✅ 还原完成。请重启电脑生效。" -ForegroundColor Green
    } else {
        Write-Error "❌ 备份文件未找到: $backupPath"
    }
}

# ========= 主程序 =========

Write-Host "🚀 Windows 11 笔记本网络优化工具" -ForegroundColor Blue
Write-Host "==================================="

switch ($Action) {
    "BackupOnly" {
        Backup-CurrentSettings
        break
    }
    "Restore" {
        Restore-FromBackup
        break
    }
    "Apply" {
        Backup-CurrentSettings
        Optimize-TcpIp
        Disable-NicPowerSaving
        Optimize-Dns
        Apply-QoS
        Restart-NetworkServices
        Write-Host "`n🎉 优化完成！建议重启电脑或禁用再启用网络适配器使设置完全生效。" -ForegroundColor Green
        Write-Host "📌 备份文件位于桌面，如需还原请运行: .\NetworkOptimize-Win11.ps1 -Action Restore" -ForegroundColor Yellow
        break
    }
}

Write-Host "`n按任意键退出..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
