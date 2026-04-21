import os
import socket
import json
import time
import hashlib
import threading
import winreg
import sys
import faulthandler
from pathlib import Path
from Crypto.Cipher import AES
from Crypto.Util.Padding import pad, unpad
from PIL import Image, ImageDraw
import pystray
import shutil
import win32gui
import win32con
import win32api

# --- 极简分流日志：同时输出到控制台和文件 ---
class TeeLogger:
    def __init__(self, file_path, mode="a"):
        self.terminal = sys.stdout
        # 允许指定模式，冷启动时可以使用 "w" 覆盖日志
        self.log = open(file_path, mode, encoding="utf-8", buffering=1)

    def write(self, message):
        if self.terminal:
            self.terminal.write(message)
        if self.log:
            self.log.write(message)
            self.log.flush()

    def flush(self):
        if self.terminal:
            self.terminal.flush()
        if self.log:
            self.log.flush()

def log(msg):
    """自定义简单的带时间戳的打印"""
    timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
    print(f"{timestamp} {msg}")

# --- 核心协议层 (miIO Protocol) ---
class MiioDevice:
    def __init__(self, ip, token_hex, device_id=None):
        self.ip = ip
        self.token = bytes.fromhex(token_hex)
        # 如果配置里有 ID 就用配置的，否则初始化为 0
        self.device_id = device_id.to_bytes(4, 'big') if device_id else b'\x00\x00\x00\x00'
        self.stamp = b'\x00\x00\x00\x00'
        self.key = hashlib.md5(self.token).digest()
        self.iv = hashlib.md5(self.key + self.token).digest()

    def _md5(self, data):
        return hashlib.md5(data).digest()

    def hello(self, target_ip=None):
        ip = target_ip or self.ip
        if not ip: return None
        hello_packet = bytes.fromhex('21310020ffffffffffffffffffffffffffffffffffffffffffffffffffffffff')
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as s:
            s.settimeout(1.0) # 缩短超时，提高探测速度
            try:
                s.sendto(hello_packet, (ip, 54321))
                data, addr = s.recvfrom(1024)
                if len(data) >= 32:
                    resp_id = data[8:12]
                    # 如果我们已知 ID，则校验 ID 是否匹配
                    if self.device_id != b'\x00\x00\x00\x00' and resp_id != self.device_id:
                        return None
                    self.device_id = resp_id
                    self.stamp = data[12:16]
                    return addr[0]
            except:
                pass
        return None

    def discover_device(self):
        """精准 UDP 发现：广播 hello 并根据响应中的 Device ID 锁定 IP"""
        device_id_int = int.from_bytes(self.device_id, 'big')
        log(f"正在局域网搜索设备 (ID: {device_id_int})...")

        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as s:
            s.setsockopt(socket.SOL_SOCKET, socket.SO_BROADCAST, 1)
            s.settimeout(1.0)
            hello_packet = bytes.fromhex('21310020ffffffffffffffffffffffffffffffffffffffffffffffffffffffff')
            
            # 发送广播包
            try:
                s.sendto(hello_packet, ('<broadcast>', 54321))
            except Exception as e:
                log(f"广播发送失败: {e}")
                return None

            start_time = time.time()
            # 在 2 秒内监听所有回包
            while time.time() - start_time < 2.0:
                try:
                    data, addr = s.recvfrom(1024)
                    if len(data) >= 32:
                        resp_id = data[8:12]
                        # 核心逻辑：校验 Device ID 是否匹配
                        if self.device_id != b'\x00\x00\x00\x00' and resp_id != self.device_id:
                            continue
                        
                        found_ip = addr[0]
                        log(f"发现匹配设备! ID: {int.from_bytes(resp_id, 'big')} -> IP: {found_ip}")
                        return found_ip
                except socket.timeout:
                    continue
                except Exception as e:
                    log(f"扫描过程异常: {e}")
                    continue
        return None

    def send_command(self, method, params):
        if not self.hello():
            return None
        payload = json.dumps({"id": 1, "method": method, "params": params}).encode()
        cipher = AES.new(self.key, AES.MODE_CBC, self.iv)
        encrypted = cipher.encrypt(pad(payload, 16))
        length = len(encrypted) + 32
        header = bytearray(bytes.fromhex('2131'))
        header.extend(length.to_bytes(2, 'big'))
        header.extend(b'\x00\x00\x00\x00')
        header.extend(self.device_id)
        header.extend(self.stamp)
        header.extend(b'\xff' * 16)
        full_data = header[:16] + self.token + encrypted
        header[16:32] = self._md5(full_data)
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as s:
            s.settimeout(2.0)
            try:
                s.sendto(header + encrypted, (self.ip, 54321))
                data, _ = s.recvfrom(1024)
                if len(data) > 32:
                    return True
                return False
            except socket.timeout:
                log("错误: 指令响应超时 (挂灯未回包)")
            except Exception as e:
                log(f"错误: 网络通信异常: {e}")
        return None

# --- Windows 系统管理 ---
def set_autostart(enabled=True):
    key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
    app_name = "MiLightBarLite"
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE)
        if enabled:
            # 方案 1：利用 uvw 的“消音器”功能
            uvw_path = shutil.which("uvw") or "uvw"
            project_dir = Path(__file__).parent.absolute()
            script_path = Path(__file__).absolute()
            app_path = f'"{uvw_path}" run --project "{project_dir}" python "{script_path}" --autostart'
            log(f"设置原生静默启动项 (uvw+python): {app_path}")
            winreg.SetValueEx(key, app_name, 0, winreg.REG_SZ, app_path)
        else:
            try: winreg.DeleteValue(key, app_name)
            except: pass
        winreg.CloseKey(key)
    except Exception as e:
        log(f"设置自启动失败: {e}")
        win32api.MessageBox(0, f"设置自启动失败: {e}", "错误", win32con.MB_ICONERROR)

def listen_power_events(on_suspend, on_resume, on_shutdown):
    """使用 pywin32 监听 Windows 电源事件，比 ctypes 更稳定"""
    def wndproc(hwnd, msg, wparam, lparam):
        if msg == win32con.WM_POWERBROADCAST:
            if wparam == win32con.PBT_APMSUSPEND:
                log("[系统事件] 检测到系统准备进入睡眠/待机...")
                on_suspend()
            elif wparam == win32con.PBT_APMRESUMESUSPEND:
                log("[系统事件] 检测到系统已从睡眠中唤醒")
                on_resume()
        elif msg == win32con.WM_ENDSESSION:
            if wparam: # True 表示正在关机或注销
                log("[系统事件] 检测到系统正在关机或注销，准备执行关灯...")
                on_shutdown()
        return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)

    # 注册窗口类
    wc = win32gui.WNDCLASS()
    wc.lpfnWndProc = wndproc
    wc.lpszClassName = "MiLightPowerEventWindow"
    wc.hInstance = win32api.GetModuleHandle(None)

    try:
        class_atom = win32gui.RegisterClass(wc)
        # 创建隐藏窗口以接收消息
        hwnd = win32gui.CreateWindow(
            class_atom, "MiLightPowerEventWindow", 0, 0, 0, 0, 0, 0, 0, wc.hInstance, None
        )
        if hwnd:
            log("电源事件监听窗口已创建")
            win32gui.PumpMessages()
    except Exception as e:
        log(f"监听电源事件失败: {e}")

# --- 主程序 ---
class MiLightController:
    def __init__(self):
        self.config_path = Path("config.json")
        self.load_config()
        self.device = MiioDevice(self.config.get("ip"), self.config.get("token"), self.config.get("device_id"))
        self.icon = None

    def load_config(self):
        """加载并严格校验配置"""
        if not self.config_path.exists():
            template = {"ip": "", "token": "", "device_id": None, "auto_start": False, "on_boot": True, "on_sleep": True, "on_shutdown": True}
            with open(self.config_path, "w") as f:
                json.dump(template, f, indent=4)
            log("错误: 配置文件 [config.json] 不存在，已为你生成模板。")
            log("请在文件中填写 [token] 以及 [ip] 或 [device_id] 后重新运行。")
            sys.exit(1)

        with open(self.config_path, "r") as f:
            self.config = json.load(f)
        
        # 严格校验逻辑
        token = self.config.get("token", "").strip()
        ip = self.config.get("ip", "").strip()
        device_id = self.config.get("device_id")

        if not token:
            log("错误: [config.json] 中未填写 [token]。控制米家设备必须提供 Token。")
            sys.exit(1)
        
        if not ip and (device_id is None):
            log("错误: [config.json] 中 [ip] 和 [device_id] 全为空。")
            log("请至少填写一个：已知 IP 则填 IP，已知硬件 ID 则填 device_id。")
            sys.exit(1)
            
        # 补全可能缺失的键
        if "device_id" not in self.config: self.config["device_id"] = None
        if "on_boot" not in self.config: self.config["on_boot"] = True
        if "on_sleep" not in self.config: self.config["on_sleep"] = True
        if "on_shutdown" not in self.config: self.config["on_shutdown"] = True

    def save_config(self):
        with open(self.config_path, "w") as f:
            json.dump(self.config, f, indent=4)

    def update_device_info(self, new_ip, new_id_bytes=None):
        changed = False
        if new_ip and new_ip != self.config.get("ip"):
            log(f"更新设备 IP: {self.config.get('ip')} -> {new_ip}")
            self.config["ip"] = new_ip
            self.device.ip = new_ip
            changed = True
        
        if new_id_bytes:
            new_id = int.from_bytes(new_id_bytes, 'big')
            if new_id != self.config.get("device_id"):
                log(f"更新设备 ID: {new_id}")
                self.config["device_id"] = new_id
                changed = True
        
        if changed:
            self.save_config()

    def toggle_light(self, state, retry=0):
        retry_msg = f" (重试第 {retry} 次)" if retry > 0 else ""
        log(f"执行控灯: {state}{retry_msg}")
        res = self.device.send_command("set_power", [state])

        if res is None and retry == 0:
            log("通信失败，尝试重新定位设备...")
            new_ip = self.device.discover_device()
            if new_ip:
                self.update_device_info(new_ip, self.device.device_id)
                res = self.device.send_command("set_power", [state])
            else:
                log("自动发现失败，设备可能已离线。")
                return

        if not res:
            if retry < 2:
                log(f"控灯未成功，2秒后进行重试...")
                time.sleep(2)
                self.toggle_light(state, retry + 1)
            else:
                log(f"控灯最终失败 ({state})")

    def on_resume_with_delay(self):
        log("系统已唤醒，等待 5 秒网络恢复后开灯...")
        time.sleep(5)
        self.toggle_light("on")

    def create_icon_image(self):
        image = Image.new('RGB', (64, 64), (255, 255, 255))
        dc = ImageDraw.Draw(image)
        dc.ellipse([10, 10, 54, 54], fill=(237, 104, 45))
        return image

    def on_toggle_setting(self, key):
        def inner(icon, item):
            self.config[key] = not self.config[key]
            status = "开启" if self.config[key] else "关闭"
            log(f"设置变更: {item.text} -> {status}")
            if key == "auto_start": set_autostart(self.config[key])
            self.save_config()
        return inner

    def run(self):
        # 启动时先尝试连接，不行就用 ID 定位
        if not self.device.hello():
            log("初始 IP 无法连接，启动自动发现定位...")
            new_ip = self.device.discover_device()
            if new_ip:
                self.update_device_info(new_ip, self.device.device_id)
            else:
                log("未能在局域网找到匹配 ID 的设备。")
        else:
            log(f"设备连接成功: {self.config['ip']} (ID: {int.from_bytes(self.device.device_id, 'big')})")
            # 顺便更新一下配置里的 ID
            self.update_device_info(self.config['ip'], self.device.device_id)

        # 检查是否是开机自启动：如果是 --autostart 且配置了 on_boot，则自动开灯
        if "--autostart" in sys.argv and self.config.get("on_boot"):
            log("[冷启动] 识别到开机自启动，执行开灯...")
            # 开机时网络可能还没完全连上，稍等一下再开灯
            threading.Timer(10, lambda: self.toggle_light("on")).start()

        def on_suspend():
            if self.config.get("on_sleep"):
                self.toggle_light("off")

        def on_shutdown():
            if self.config.get("on_shutdown"):
                self.toggle_light("off")

        def on_resume():
            if self.config.get("on_boot"):
                self.on_resume_with_delay()

        threading.Thread(target=listen_power_events, args=(on_suspend, on_resume, on_shutdown), daemon=True).start()

        def get_menu():
            return pystray.Menu(
                pystray.MenuItem("开启挂灯 (ON)", lambda: self.toggle_light("on")),
                pystray.MenuItem("关闭挂灯 (OFF)", lambda: self.toggle_light("off")),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem("开机时开灯", self.on_toggle_setting("on_boot"), checked=lambda item: self.config["on_boot"]),
                pystray.MenuItem("待机时关灯", self.on_toggle_setting("on_sleep"), checked=lambda item: self.config["on_sleep"]),
                pystray.MenuItem("关机时关灯", self.on_toggle_setting("on_shutdown"), checked=lambda item: self.config["on_shutdown"]),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem("开机自启动", self.on_toggle_setting("auto_start"), checked=lambda item: self.config["auto_start"]),
                pystray.MenuItem("查看日志", lambda: os.startfile(Path(__file__).parent / "app.log")),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem("退出程序", self.quit_app)
            )
        
        self.icon = pystray.Icon("MiLight", self.create_icon_image(), "米家挂灯控制器 Lite", get_menu())
        self.icon.run()

    def quit_app(self, icon, item):
        self.icon.stop()
        os._exit(0)

if __name__ == "__main__":
    # 强制切换工作目录到脚本所在目录，防止自启动时路径错乱
    os.chdir(Path(__file__).parent.absolute())
    
    # 识别是否为冷启动（自启动）
    is_autostart = "--autostart" in sys.argv
    log_mode = "w" if is_autostart else "a"
    
    # 初始化日志系统
    sys.stdout = TeeLogger(Path(__file__).parent / "app.log", mode=log_mode)
    sys.stderr = sys.stdout

    # 启用崩溃处理器，记录底层 C 崩溃 (access violation 等)
    crash_log_path = Path(__file__).parent / "crash.log"
    # 使用 "w" 模式，每次启动覆盖旧的崩溃日志
    crash_log_file = open(crash_log_path, "w", encoding="utf-8", buffering=1)
    # 写入本次运行的启动时间
    crash_log_file.write(f"{'='*20} [会话启动: {time.strftime('%Y-%m-%d %H:%M:%S')}] {'='*20}\n")
    faulthandler.enable(file=crash_log_file)

    if is_autostart:
        log("米家挂灯控制器 Lite [冷启动] 覆盖日志并运行")
    else:
        log(f"米家挂灯控制器 Lite 启动成功 (时间: {time.strftime('%Y-%m-%d %H:%M:%S')})")

    try:
        controller = MiLightController()
        controller.run()
    except Exception as e:
        import traceback
        error_msg = f"程序发生致命错误:\n{e}\n\n{traceback.format_exc()}"
        log(error_msg)
        # 无论是否有控制台，弹窗告知用户，防止静默闪退
        win32api.MessageBox(0, error_msg, "致命错误", win32con.MB_ICONERROR)
