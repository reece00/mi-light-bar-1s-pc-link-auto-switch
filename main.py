import os
import socket
import json
import time
import hashlib
import logging
import threading
import queue
import heapq
import ctypes
import ctypes.wintypes as wintypes
import winreg
import sys
import faulthandler
from pathlib import Path
from logging.handlers import RotatingFileHandler
from Crypto.Cipher import AES
from Crypto.Util.Padding import pad
from PIL import Image, ImageDraw
import pystray
import shutil
import win32gui
import win32con
import win32api
import win32event
import winerror

_APP_LOGGER = logging.getLogger("MiLightBarLite")
_DEBUG_LOGGER = logging.getLogger("MiLightBarLiteDebug")


def setup_runtime_logger(file_path):
    _APP_LOGGER.setLevel(logging.INFO)
    _APP_LOGGER.propagate = False
    _APP_LOGGER.handlers.clear()

    formatter = logging.Formatter("%(asctime)s %(message)s", datefmt="%Y-%m-%d %H:%M:%S")
    file_handler = RotatingFileHandler(
        file_path,
        mode="a",
        maxBytes=10 * 1024,
        backupCount=2,
        encoding="utf-8",
    )
    file_handler.setFormatter(formatter)
    _APP_LOGGER.addHandler(file_handler)

    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter)
    _APP_LOGGER.addHandler(console_handler)


def setup_debug_logger(file_path):
    _DEBUG_LOGGER.setLevel(logging.DEBUG)
    _DEBUG_LOGGER.propagate = False
    _DEBUG_LOGGER.handlers.clear()

    formatter = logging.Formatter("%(asctime)s [DEBUG] %(message)s", datefmt="%Y-%m-%d %H:%M:%S")
    # 调试日志固定为单文件：每次启动覆盖，避免多副本与持续膨胀。
    file_handler = logging.FileHandler(
        file_path,
        mode="w",
        encoding="utf-8",
    )
    file_handler.setFormatter(formatter)
    _DEBUG_LOGGER.addHandler(file_handler)


def log(msg):
    """统一运行日志输出。"""
    if _APP_LOGGER.handlers:
        _APP_LOGGER.info(msg)
    else:
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        print(f"{timestamp} {msg}")


def log_debug(msg):
    if _DEBUG_LOGGER.handlers:
        _DEBUG_LOGGER.debug(msg)


GUID_CONSOLE_DISPLAY_STATE = "{6FE69556-704A-47A0-8F24-C28D936FDA47}"
PBT_POWERSETTINGCHANGE = 0x8013


class GUID(ctypes.Structure):
    _fields_ = [
        ("Data1", wintypes.DWORD),
        ("Data2", wintypes.WORD),
        ("Data3", wintypes.WORD),
        ("Data4", ctypes.c_ubyte * 8),
    ]

    def as_string(self):
        buffer = ctypes.create_unicode_buffer(39)
        ctypes.windll.ole32.StringFromGUID2(ctypes.byref(self), buffer, len(buffer))
        return buffer.value.upper()


class POWERBROADCAST_SETTING(ctypes.Structure):
    _fields_ = [
        ("PowerSetting", GUID),
        ("DataLength", wintypes.DWORD),
        ("Data", wintypes.DWORD),
    ]


def guid_from_string(guid_text):
    guid = GUID()
    ctypes.windll.ole32.CLSIDFromString(ctypes.c_wchar_p(guid_text), ctypes.byref(guid))
    return guid

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
            except socket.timeout:
                return None
            # 注意：这里不再捕获并吞没 OSError，让外层去处理网络未就绪的情况
        return None

    def discover_device(self):
        """精准 UDP 发现：广播 hello 并根据响应中的 Device ID 锁定 IP"""
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as s:
            s.setsockopt(socket.SOL_SOCKET, socket.SO_BROADCAST, 1)
            s.settimeout(1.0)
            hello_packet = bytes.fromhex('21310020ffffffffffffffffffffffffffffffffffffffffffffffffffffffff')
            
            # 发送广播包 (不再内部重试网卡，直接发)
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
            
            log(f"搜索失败: 未在局域网发现 ID 为 {int.from_bytes(self.device_id, 'big')} 的设备。")
        return None

    def send_command(self, method, params, max_retries=3, retry_delay=2):
        payload = json.dumps({"id": 1, "method": method, "params": params}).encode()
        start_t = time.time()
        attempt_count = max(1, max_retries)
        log_debug(
            f"send_command start: method={method}, ip={self.ip}, retries={attempt_count}, payload_len={len(payload)}"
        )

        for i in range(attempt_count):
            try:
                attempt_no = i + 1
                elapsed = time.time() - start_t
                log_debug(f"send_command attempt {attempt_no}/{attempt_count} start, elapsed={elapsed:.2f}s")
                # 检查是否发生了系统休眠导致的长时间跳变
                if time.time() - start_t > 20: # 单次指令逻辑不应超过 20s
                    log_debug("send_command aborted by 20s guard before hello")
                    return "TIMEOUT"

                # 1. 尝试握手 (必须拿到最新的 stamp 才能组装包)
                if not self.hello():
                    log_debug(f"hello timeout/no response at attempt {attempt_no}/{attempt_count}")
                    # hello 返回 None，说明网络通但超时，或者目标 IP 错。
                    if i < attempt_count - 1:
                        log(f"握手失败，{retry_delay}秒后重试... ({i + 1}/{attempt_count})")
                        time.sleep(retry_delay)
                        continue
                    if time.time() - start_t < 20:
                        log("错误: 握手失败 (目标无响应)")
                    return "TIMEOUT"

                # 2. 握手成功，组装加密包
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
                
                # 3. 发送控制指令
                with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as s:
                    s.settimeout(1.0)
                    s.sendto(header + encrypted, (self.ip, 54321))
                    data, _ = s.recvfrom(1024)
                    if len(data) > 32:
                        log_debug(f"send_command success at attempt {attempt_no}/{attempt_count}, resp_len={len(data)}")
                        return "SUCCESS"
                    # 收到异常短包（应用层拒绝）
                    log_debug(f"short response at attempt {attempt_no}/{attempt_count}, resp_len={len(data)}")
                    if i < attempt_count - 1:
                        log(f"收到异常短回包，{retry_delay}秒后重试... ({i + 1}/{attempt_count})")
                        time.sleep(retry_delay)
                        continue
                    log("错误: 收到异常回包")
                    return "TIMEOUT"
                    
            except socket.timeout:
                log_debug(f"socket.timeout at attempt {attempt_no}/{attempt_count}")
                if time.time() - start_t > 20:
                    return "TIMEOUT"
                if i < attempt_count - 1:
                    log(f"指令响应超时，{retry_delay}秒后重试... ({i + 1}/{attempt_count})")
                    time.sleep(retry_delay)
                    continue
                log("错误: 指令响应超时 (挂灯未回包)")
                return "TIMEOUT"
            except OSError as e:
                log_debug(f"oserror at attempt {attempt_no}/{attempt_count}: {e!r}")
                if time.time() - start_t > 20:
                    return "NETWORK_ERROR"
                # 无论是 hello() 还是 sendto() 抛出的 OSError (如网卡没好 10065) 都在这里被拦截
                if i < attempt_count - 1:
                    log(f"网络异常 ({e})，可能网卡未就绪，{retry_delay}秒后重试... ({i + 1}/{attempt_count})")
                    time.sleep(retry_delay)
                    continue
                log(f"错误: 网络通信最终失败: {e}")
                return "NETWORK_ERROR"
            except Exception as e:
                log_debug(f"unexpected exception at attempt {attempt_no}/{attempt_count}: {e!r}")
                log(f"错误: 其他通信异常: {e}")
                return "NETWORK_ERROR"
                
        log_debug("send_command exhausted retries, return NETWORK_ERROR")
        return "NETWORK_ERROR"

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
            try:
                winreg.DeleteValue(key, app_name)
            except FileNotFoundError:
                pass
        winreg.CloseKey(key)
        return True
    except Exception as e:
        log(f"设置自启动失败: {e}")
        win32api.MessageBox(0, f"设置自启动失败: {e}", "错误", win32con.MB_ICONERROR)
        return False

def listen_power_events(on_suspend, on_resume, on_shutdown, on_display_state_change):
    """监听睡眠/唤醒/关机以及显示器状态变化。"""
    display_guid = guid_from_string(GUID_CONSOLE_DISPLAY_STATE)
    display_listener_ready = False
    shutdown_queued = False

    def wndproc(hwnd, msg, wparam, lparam):
        nonlocal shutdown_queued
        if msg == win32con.WM_POWERBROADCAST:
            if wparam == win32con.PBT_APMSUSPEND:
                log("[系统事件] 检测到系统准备进入睡眠/待机...")
                on_suspend()
            elif wparam == win32con.PBT_APMRESUMESUSPEND:
                log("[系统事件] 检测到系统已从睡眠中唤醒")
                on_resume()
            elif wparam == PBT_POWERSETTINGCHANGE and display_listener_ready:
                setting = POWERBROADCAST_SETTING.from_address(lparam)
                if setting.PowerSetting.as_string() == GUID_CONSOLE_DISPLAY_STATE:
                    display_state = setting.Data
                    state_name = {0: "OFF", 1: "ON", 2: "DIM"}.get(display_state, f"UNKNOWN({display_state})")
                    log(f"[系统事件] 检测到显示器状态变化: {state_name}")
                    on_display_state_change(display_state)
        elif msg == win32con.WM_QUERYENDSESSION:
            if not shutdown_queued:
                shutdown_queued = True
                log("[系统事件] 检测到系统请求结束会话，开始同步执行关灯...")
                on_shutdown()
            return 1
        elif msg == win32con.WM_ENDSESSION:
            if wparam: # True 表示正在关机或注销
                log("[系统事件] 检测到系统正在关机或注销")
                if not shutdown_queued:
                    shutdown_queued = True
                    log("[系统事件] 未提前排队，改为在结束会话阶段兜底执行关灯...")
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
            try:
                notify_handle = ctypes.windll.user32.RegisterPowerSettingNotification(
                    wintypes.HANDLE(hwnd), ctypes.byref(display_guid), 0
                )
                if not notify_handle:
                    raise ctypes.WinError()
                display_listener_ready = True
                log("显示器状态监听已注册")
            except Exception as e:
                log(f"显示器状态监听注册失败，将仅保留睡眠/唤醒/关机监听: {e}")
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
        self.last_display_state = None
        self.post_resume_not_before = 0.0
        self.action_id_lock = threading.Lock()
        self.latest_action_id = 0
        self.shutdown_event = threading.Event()
        self.command_queue = queue.Queue()
        self.schedule_condition = threading.Condition()
        self.scheduled_actions = []
        self.command_worker = threading.Thread(target=self._command_worker_loop, daemon=True)
        self.scheduler_worker = threading.Thread(target=self._scheduler_loop, daemon=True)
        self.command_worker.start()
        self.scheduler_worker.start()

    def load_config(self):
        """加载并严格校验配置"""
        if not self.config_path.exists():
            template = {
                "ip": "", 
                "token": "", 
                "device_id": None, 
                "auto_start": False, 
                "on_boot": True, 
                "on_resume": True,
                "on_sleep": True, 
                "on_shutdown": True,
                "on_display_sync": False,
                "resume_delay": 10,
                "post_resume_delay": 4,
                "discovery_retries": 3,
                "command_retries": 2
            }
            with open(self.config_path, "w") as f:
                json.dump(template, f, indent=4)
            msg = "错误: 配置文件 [config.json] 不存在，已为你生成模板。\n请在文件中填写 [token] 以及 [ip] 或 [device_id] 后重新运行。"
            log(msg)
            win32api.MessageBox(0, msg, "配置错误", win32con.MB_ICONERROR)
            sys.exit(1)

        with open(self.config_path, "r") as f:
            self.config = json.load(f)
        
        # 严格校验逻辑
        token = self.config.get("token", "").strip()
        ip = self.config.get("ip", "").strip()
        device_id = self.config.get("device_id")

        if not token:
            msg = "错误: [config.json] 中未填写 [token]。控制米家设备必须提供 Token。"
            log(msg)
            win32api.MessageBox(0, msg, "配置错误", win32con.MB_ICONERROR)
            sys.exit(1)
        
        if not ip and (device_id is None):
            msg = "错误: [config.json] 中 [ip] 和 [device_id] 全为空。\n请至少填写一个：已知 IP 则填 IP，已知硬件 ID 则填 device_id。"
            log(msg)
            win32api.MessageBox(0, msg, "配置错误", win32con.MB_ICONERROR)
            sys.exit(1)
            
        # 补全可能缺失的键
        if "device_id" not in self.config: self.config["device_id"] = None
        if "on_boot" not in self.config: self.config["on_boot"] = True
        if "on_resume" not in self.config: self.config["on_resume"] = self.config["on_boot"]
        if "on_sleep" not in self.config: self.config["on_sleep"] = True
        if "on_shutdown" not in self.config: self.config["on_shutdown"] = True
        if "on_display_sync" not in self.config: self.config["on_display_sync"] = False

        # 新增可选配置及其默认值
        if "resume_delay" not in self.config: self.config["resume_delay"] = 10
        if "post_resume_delay" not in self.config: self.config["post_resume_delay"] = 4
        if "discovery_retries" not in self.config: self.config["discovery_retries"] = 3
        if "command_retries" not in self.config: self.config["command_retries"] = 2

    def save_config(self):
        temp_path = self.config_path.with_name(f"{self.config_path.name}.tmp")
        try:
            with open(temp_path, "w", encoding="utf-8") as f:
                json.dump(self.config, f, indent=4)
                f.flush()
                os.fsync(f.fileno())
            temp_path.replace(self.config_path)
        finally:
            if temp_path.exists():
                try:
                    temp_path.unlink()
                except OSError:
                    pass

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

    def _next_action_id(self):
        with self.action_id_lock:
            self.latest_action_id += 1
            return self.latest_action_id

    def _is_action_current(self, action_id):
        with self.action_id_lock:
            return action_id == self.latest_action_id

    def submit_light_action(self, state, is_auto=False, force_no_retry=False, delay=0, reason=""):
        action = {
            "action_id": self._next_action_id(),
            "state": state,
            "is_auto": is_auto,
            "force_no_retry": force_no_retry,
            "reason": reason,
        }
        if delay > 0:
            self._schedule_action(action, delay)
        else:
            self.command_queue.put(action)

    def _schedule_action(self, action, delay):
        due_at = time.monotonic() + delay
        with self.schedule_condition:
            heapq.heappush(self.scheduled_actions, (due_at, action["action_id"], action))
            self.schedule_condition.notify()

    def _scheduler_loop(self):
        while not self.shutdown_event.is_set():
            with self.schedule_condition:
                while not self.scheduled_actions and not self.shutdown_event.is_set():
                    self.schedule_condition.wait()

                if self.shutdown_event.is_set():
                    return

                due_at, _, action = self.scheduled_actions[0]
                wait_seconds = due_at - time.monotonic()
                if wait_seconds > 0:
                    self.schedule_condition.wait(wait_seconds)
                    continue

                heapq.heappop(self.scheduled_actions)

            if not self._is_action_current(action["action_id"]):
                continue

            self.command_queue.put(action)

    def _command_worker_loop(self):
        while not self.shutdown_event.is_set():
            try:
                action = self.command_queue.get(timeout=0.5)
            except queue.Empty:
                continue

            if action is None:
                self.command_queue.task_done()
                return

            try:
                if not self._is_action_current(action["action_id"]):
                    log_debug(
                        f"worker skip stale action_id={action['action_id']}, latest={self.latest_action_id}, state={action['state']}"
                    )
                    continue
                log_debug(
                    f"worker executing action_id={action['action_id']}, state={action['state']}, reason={action['reason']}"
                )
                self.toggle_light(
                    action["state"],
                    is_auto=action["is_auto"],
                    force_no_retry=action["force_no_retry"],
                    action_id=action["action_id"],
                    reason=action["reason"],
                )
            except Exception as e:
                log(f"控灯任务执行异常: {e}")
            finally:
                self.command_queue.task_done()

    def toggle_light(self, state, is_auto=False, force_no_retry=False, action_id=None, reason=""):
        start_time = time.time()
        log_debug(
            f"toggle_light start: state={state}, is_auto={is_auto}, force_no_retry={force_no_retry}, action_id={action_id}, reason={reason}"
        )

        # 根据控制来源动态决定重试策略
        if is_auto and not force_no_retry:
            max_retries = self.config.get("command_retries", 2)
            # 自动模式下，底层网络抗压使用默认的 3 次
            network_retries = 3
        else:
            # 手动控制或强制不重试：极速失败，不进行业务重试
            max_retries = 0
            network_retries = 2

        for retry in range(max_retries + 1):
            if action_id is not None and not self._is_action_current(action_id):
                log(f"指令 {state} 已过期，停止执行。原因: {reason}")
                log_debug(f"toggle_light stop stale action before retry loop: action_id={action_id}")
                return

            retry_msg = f" (重试第 {retry} 次)" if retry > 0 else ""
            source_msg = "[自动]" if is_auto else "[手动]"
            log(f"执行控灯 {source_msg}: {state}{retry_msg} | {reason}")
            log_debug(f"toggle_light retry={retry}, network_retries={network_retries}")
            
            # 让底层带容错地去发指令
            res = self.device.send_command("set_power", [state], max_retries=network_retries)
            
            # 检查执行时长。如果跨度超过 30 秒，极大概率是中途系统休眠了
            if time.time() - start_time > 30:
                log(f"指令 {state} 执行跨度异常 ({int(time.time() - start_time)}s)，检测到休眠恢复，忽略此过期指令。")
                log_debug("toggle_light aborted by 30s guard")
                return

            if res == "SUCCESS":
                log_debug("toggle_light result SUCCESS")
                return # 完美成功
                
            if res == "NETWORK_ERROR":
                # 如果是自动控制，网卡可能只是暂时没好，允许通过外层业务重试来“死磕”
                if is_auto and retry < max_retries:
                    log("本地网络暂时不可达，等待下一轮自动重试...")
                else:
                    log("本地网络瘫痪，取消控灯。")
                    log_debug("toggle_light result NETWORK_ERROR and stop")
                    return # 手动控制或重试耗尽，彻底放弃，绝不广播
                
            elif res == "TIMEOUT":
                log_debug("toggle_light result TIMEOUT")
                # 超时/无响应（可能是没理我，也可能是 IP 变了）
                # 只有在第一次尝试失败时，才去尝试重新定位设备
                if retry == 0:
                    log("错误: 目标无响应，正在重新搜索定位设备...")
                    discovery_retries = max(1, self.config.get("discovery_retries", 3))
                    for discovery_retry in range(discovery_retries):
                        if action_id is not None and not self._is_action_current(action_id):
                            log(f"指令 {state} 在重新定位设备前已过期，停止执行。原因: {reason}")
                            log_debug("toggle_light stop stale action during discovery")
                            return
                        if discovery_retry > 0:
                            log(f"重新搜索定位设备，第 {discovery_retry + 1} 次尝试...")
                        new_ip = self.device.discover_device()
                        if new_ip:
                            self.update_device_info(new_ip, self.device.device_id)
                            log_debug(f"discovery success, new_ip={new_ip}")
                            # 定位成功，立刻用新 IP 重发一次，如果不成功，进入下一轮业务重试
                            res = self.device.send_command("set_power", [state], max_retries=1)
                            if res == "SUCCESS":
                                log_debug("toggle_light redispatch success after discovery")
                                return
                            break
                
            # 业务重试前的等待
            if retry < max_retries:
                log(f"控灯未成功，2秒后进行业务重试...")
                time.sleep(2)
                if action_id is not None and not self._is_action_current(action_id):
                    log(f"指令 {state} 在重试间隙已过期，停止。原因: {reason}")
                    log_debug("toggle_light stop stale action after retry sleep")
                    return
                
        log(f"控灯最终失败 ({state})")
        log_debug("toggle_light final failure")

    def handle_shutdown_event(self):
        if not self.config.get("on_shutdown"):
            return

        log("执行控灯 [自动]: off | 系统正在关机或注销")
        res = self.device.send_command("set_power", ["off"], max_retries=1, retry_delay=0)
        if res == "SUCCESS":
            log("[系统事件] 同步关灯成功。")
        else:
            log(f"[系统事件] 同步关灯失败: {res}")

    def handle_auto_event(self, event_name, display_state=None):
        if event_name == "cold_start":
            if self.config.get("on_boot"):
                self.submit_light_action("on", is_auto=True, delay=10, reason="冷启动开机自启动")
            return

        if event_name == "suspend":
            if self.config.get("on_sleep"):
                self.submit_light_action("off", is_auto=True, force_no_retry=True, reason="系统准备进入睡眠/待机")
            return

        if event_name == "shutdown":
            if self.config.get("on_shutdown"):
                self.submit_light_action("off", is_auto=True, force_no_retry=True, reason="系统正在关机或注销")
            return

        if event_name == "resume":
            post_resume_delay = max(0, self.config.get("post_resume_delay", 4))
            self.post_resume_not_before = time.monotonic() + post_resume_delay
            if self.config.get("on_display_sync"):
                log_debug(
                    f"[自动][恢复] 已启用显示器联动：等待显示器 ON 事件触发开灯，"
                    f"恢复后冷却窗口={post_resume_delay}s。"
                )
                return
            if post_resume_delay > 0:
                log(f"[自动][恢复] 检测到睡眠恢复，网络冷却窗口 {post_resume_delay}s（仅影响恢复后的开灯发送）。")
            if self.config.get("on_resume"):
                delay = self.config.get("resume_delay", 10)
                self.submit_light_action("on", is_auto=True, delay=delay, reason="系统唤醒后延时开灯")
            return

        if event_name == "display_state":
            if display_state == self.last_display_state:
                return

            self.last_display_state = display_state

            if not self.config.get("on_display_sync"):
                return

            if display_state == 0:
                self.submit_light_action("off", is_auto=True, force_no_retry=True, reason="显示器已关闭")
            elif display_state == 1:
                now = time.monotonic()
                extra_delay = max(0, self.post_resume_not_before - now)
                if extra_delay > 0:
                    log_debug(f"[自动][显示器] 收到 ON 事件，恢复后网络冷却未结束，{extra_delay:.1f}s 后开灯。")
                self.submit_light_action("on", is_auto=True, delay=extra_delay, reason="显示器 ON 事件触发")
            elif display_state == 2:
                pass
            return

        log(f"收到未知自动事件: {event_name}")

    def create_icon_image(self):
        image = Image.new('RGB', (64, 64), (255, 255, 255))
        dc = ImageDraw.Draw(image)
        dc.ellipse([10, 10, 54, 54], fill=(237, 104, 45))
        return image

    def on_toggle_setting(self, key):
        def inner(icon, item):
            new_value = not self.config[key]
            if key == "auto_start" and not set_autostart(new_value):
                log(f"设置变更失败: {item.text} 保持不变")
                return

            self.config[key] = new_value
            status = "开启" if new_value else "关闭"
            log(f"设置变更: {item.text} -> {status}")
            self.save_config()
        return inner

    def run(self):
        log(f"服务启动: IP={self.config['ip']}, ID={self.config.get('device_id')}")

        # 检查是否是开机自启动：如果是 --autostart 且配置了 on_boot，则自动开灯
        if "--autostart" in sys.argv and self.config.get("on_boot"):
            self.handle_auto_event("cold_start")

        def on_suspend():
            self.handle_auto_event("suspend")

        def on_shutdown():
            self.handle_shutdown_event()

        def on_resume():
            self.handle_auto_event("resume")

        def on_display_state_change(display_state):
            self.handle_auto_event("display_state", display_state=display_state)

        threading.Thread(
            target=listen_power_events,
            args=(on_suspend, on_resume, on_shutdown, on_display_state_change),
            daemon=True,
        ).start()

        def get_menu():
            return pystray.Menu(
                pystray.MenuItem("开启挂灯 (ON)", lambda: self.submit_light_action("on", reason="托盘手动开启挂灯")),
                pystray.MenuItem("关闭挂灯 (OFF)", lambda: self.submit_light_action("off", reason="托盘手动关闭挂灯")),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem("开机时开灯", self.on_toggle_setting("on_boot"), checked=lambda item: self.config["on_boot"]),
                pystray.MenuItem("唤醒后开灯", self.on_toggle_setting("on_resume"), checked=lambda item: self.config["on_resume"]),
                pystray.MenuItem("待机时关灯", self.on_toggle_setting("on_sleep"), checked=lambda item: self.config["on_sleep"]),
                pystray.MenuItem("关机时关灯", self.on_toggle_setting("on_shutdown"), checked=lambda item: self.config["on_shutdown"]),
                pystray.MenuItem("显示器熄灭/唤醒联动", self.on_toggle_setting("on_display_sync"), checked=lambda item: self.config["on_display_sync"]),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem("开机自启动", self.on_toggle_setting("auto_start"), checked=lambda item: self.config["auto_start"]),
                pystray.MenuItem("查看日志", lambda: os.startfile(Path(__file__).parent / "app.log")),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem("退出程序", self.quit_app)
            )
        
        self.icon = pystray.Icon("MiLight", self.create_icon_image(), "米家挂灯控制器 Lite", get_menu())
        self.icon.run()

    def quit_app(self, icon, item):
        log("收到退出请求，正在停止调度与托盘...")
        self.shutdown_event.set()
        with self.schedule_condition:
            self.schedule_condition.notify_all()
        self.command_queue.put(None)
        self.icon.stop()

if __name__ == "__main__":
    # 防止多实例运行
    mutex_name = "Global\\MiLightBarLite_SingleInstance_Mutex"
    mutex = win32event.CreateMutex(None, False, mutex_name)
    if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
        # 如果是静默自启动，就不弹窗了，直接退出
        if "--autostart" not in sys.argv:
            win32api.MessageBox(0, "程序已经在运行中，请检查系统托盘。", "提示", win32con.MB_ICONINFORMATION)
        sys.exit(0)

    # 强制切换工作目录到脚本所在目录，防止自启动时路径错乱
    os.chdir(Path(__file__).parent.absolute())
    
    # 识别是否为冷启动（自启动）
    is_autostart = "--autostart" in sys.argv
    app_log_path = Path(__file__).parent / "app.log"
    debug_log_path = Path(__file__).parent / "debug.log"

    # 初始化运行日志系统（大小轮转，保留 2 份历史）
    setup_runtime_logger(app_log_path)
    setup_debug_logger(debug_log_path)

    # 启用崩溃处理器，记录底层 C 崩溃 (access violation 等)
    crash_log_path = Path(__file__).parent / "crash.log"
    # 使用 "w" 模式，每次启动覆盖旧的崩溃日志
    crash_log_file = open(crash_log_path, "w", encoding="utf-8", buffering=1)
    # 写入本次运行的启动时间
    crash_log_file.write(f"{'='*20} [会话启动: {time.strftime('%Y-%m-%d %H:%M:%S')}] {'='*20}\n")
    faulthandler.enable(file=crash_log_file)

    start_mode = "autostart" if is_autostart else "manual"
    log(f"米家挂灯控制器 Lite 启动成功 (mode={start_mode})")

    try:
        controller = MiLightController()
        controller.run()
    except Exception as e:
        import traceback
        error_msg = f"程序发生致命错误:\n{e}\n\n{traceback.format_exc()}"
        log(error_msg)
        # 无论是否有控制台，弹窗告知用户，防止静默闪退
        win32api.MessageBox(0, error_msg, "致命错误", win32con.MB_ICONERROR)
