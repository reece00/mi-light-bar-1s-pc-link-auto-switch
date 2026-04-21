# 米家显示器挂灯1S PC-windows 自动控制

一个极致轻量的米家屏幕挂灯 1S 局域网后台控制器（仅限 Windows）。无需打开米家 App，让挂灯随电脑状态自动开关灯，没有调光等功能，需要局域网。

## 核心特性
- **电脑联动自动开关**：集成 Windows 电源事件，实现随电脑状态（休眠、唤醒、关机）自动执行挂灯开关指令。
- **轻量**：无 Electron，无重型 GUI。托盘化常驻运行，内存占用极低（< 20MB）。
- **静默守护**：除必要的状态变更日志外，全程无感运行，支持开机自动静默启动。

## 快速开始

### 0. 准备工作
在开始之前，请确保你的系统已安装以下工具：
- **Python 3.14+**: 程序核心运行环境。
- **uv**:  Python 包管理工具（推荐）。如果你还没有安装，可以在 PowerShell 中运行：
  ```powershell
  powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
  ```

### 1. 安装环境
#### 方式一：使用 uv (推荐)
```powershell
pwsh install.ps1
```
*脚本会自动创建虚拟环境并配置必要的 Windows 系统组件。*

#### 方式二：使用 pip 手动安装
```powershell
# 1. 安装依赖
pip install pillow pycryptodome pystray pywin32

# 2. 必须执行：配置 pywin32 系统组件
python -m pywin32_postinstall -install
```

### 2. 配置设备
编辑项目根目录下的 `config.json`，填入以下必要信息：
- **Token**: 设备的 32 位令牌。可使用 [Xiaomi-cloud-tokens-extractor](https://github.com/PiotrMachowski/Xiaomi-cloud-tokens-extractor) 等工具获取。
- **IP**: 挂灯的局域网 IP。

### 3. 运行
#### 使用 uv
```powershell
uv run python main.py
```

#### 手动运行
```powershell
python main.py
```
运行后，右键点击系统托盘图标可进行手动控制及自动化设置。

## 技术原理

- **控制协议**：基于小米 miIO 协议，通过 UDP 54321 端口发送 AES-CBC 加密指令。
- **自动化触发**：使用 `pywin32` 注册隐藏窗口，实时监听 Windows 系统的 `WM_POWERBROADCAST` 消息（睡眠/唤醒）及 `WM_ENDSESSION` 消息（关机）。
- **设备定位**：若初始连接失败，程序会发送 UDP 广播包。通过校验响应包中的硬件 Device ID 唯一标识，动态更新并锁定设备当前 IP。
