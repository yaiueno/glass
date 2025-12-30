import sys
import os
import time
import ctypes
import subprocess
import threading
import math
import struct
import traceback
import datetime

# === ログ保存機能 ===
class DualLogger:
    def __init__(self):
        filename = f"log_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        self.terminal = sys.stdout
        self.log_file = open(filename, "w", encoding='utf-8')
        start_msg = f"=== Log Started: {filename} ===\n"
        self.terminal.write(start_msg)
        self.log_file.write(start_msg)

    def write(self, message):
        self.terminal.write(message)
        self.log_file.write(message)
        self.log_file.flush()

    def flush(self):
        self.terminal.flush()
        self.log_file.flush()

if not os.environ.get("IDLE_MODE"):
    logger = DualLogger()
    sys.stdout = logger
    sys.stderr = logger

# === ライブラリ読み込み ===
try:
    import win32api
    import win32con
    import pywintypes
    import moderngl
    import pygame
    from pygame.locals import *
    import glm
    import numpy as np
    import hid
    import dxcam
    
    # --- DXCam修正パッチ ---
    def patched_release(self):
        if hasattr(self, '_duplicator') and self._duplicator is not None:
            try:
                self._duplicator.Release()
            except: pass
            self._duplicator = None
    dxcam.DXCamera.release = patched_release
    # ----------------------

except ImportError as e:
    print(f"エラー: ライブラリ不足 ({e})")
    sys.exit()

# 高DPI & タイマー精度
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
    ctypes.windll.winmm.timeBeginPeriod(1)
except:
    ctypes.windll.user32.SetProcessDPIAware()

# --- 設定 ---
DRIVER_PATH = r"C:\Program Files\usbmmidd_v2"
INSTALLER_EXE = "deviceinstaller64.exe"
TARGET_WIDTH = 6000
TARGET_HEIGHT = 1080

# ビュワー設定
VIEWER_CONFIG = {
    'RADIUS': 3.0,       # 基準距離
    'ARC_ANGLE': 140.0,  # 湾曲具合
    'SEGMENTS': 64,
    'FOV': 45.0,         # Zoom用視野角
}

# IMU設定
VENDOR_ID = 0x4817
PRODUCT_ID = 0x4242
INVERT_PITCH = False 
INVERT_YAW = False
INVERT_ROLL = True

# === シェーダー ===
# メイン画面用
VERTEX_SHADER = '''
#version 330
in vec3 in_position;
in vec2 in_texcoord;
uniform mat4 m_proj;
uniform mat4 m_view;
uniform mat4 m_model;
out vec2 v_texcoord;
void main() {
    gl_Position = m_proj * m_view * m_model * vec4(in_position, 1.0);
    v_texcoord = in_texcoord; 
}
'''

FRAGMENT_SHADER = '''
#version 330
uniform sampler2D Texture;
in vec2 v_texcoord;
out vec4 f_color;
void main() {
    f_color = vec4(texture(Texture, v_texcoord).rgb, 1.0);
}
'''

# FPS表示用
UI_VERTEX_SHADER = '''
#version 330
in vec2 in_pos;
in vec2 in_uv;
out vec2 v_uv;
void main() {
    gl_Position = vec4(in_pos, 0.0, 1.0);
    v_uv = in_uv;
}
'''

UI_FRAGMENT_SHADER = '''
#version 330
uniform sampler2D uiTexture;
in vec2 v_uv;
out vec4 f_color;
void main() {
    vec4 c = texture(uiTexture, v_uv);
    if(c.a < 0.1) discard;
    f_color = c;
}
'''

def is_admin():
    try: return ctypes.windll.shell32.IsUserAnAdmin()
    except: return False

# --- 仮想ディスプレイ管理 ---
class VirtualDisplayManager:
    def __init__(self, driver_path, exe_name):
        self.driver_path = driver_path
        self.exe_path = os.path.join(driver_path, exe_name)

    def _run(self, args):
        try:
            subprocess.run([self.exe_path] + args, cwd=self.driver_path, 
                          check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            return True
        except: return False

    def enable(self):
        print(f">>> 仮想ディスプレイ({TARGET_WIDTH}x{TARGET_HEIGHT}) セットアップ...")
        self._run(["enableidd", "0"]) 
        # 【修正】ここを0.5秒から2.0秒に延長。ドライバの確実なアンロード待ち
        time.sleep(2.0)
        self._run(["enableidd", "1"])
        
        # 【修正】有効化後、即座に解像度変更に行かず少し待つ
        print(">>> ドライバ反映待ち(3秒)...")
        time.sleep(3.0)
        
        print(">>> 認識待機中(最大10秒)...")
        for i in range(10):
            if self._force_resolution():
                print(f">>> [成功] 解像度設定完了: {TARGET_WIDTH}x{TARGET_HEIGHT}")
                return
            time.sleep(1.0)
        print("!!! 警告: 解像度変更に失敗しました")

    def disable(self):
        print("\n>>> 仮想ディスプレイを削除中...")
        self._run(["enableidd", "0"])

    def _force_resolution(self):
        i = 0
        while True:
            try:
                dev = win32api.EnumDisplayDevices(None, i)
                if self._check_and_apply(dev.DeviceName): return True
                i += 1
            except: break
        return False

    def _check_and_apply(self, dev_name):
        j = 0
        target_mode = None
        while True:
            try:
                m = win32api.EnumDisplaySettings(dev_name, j)
                if m.PelsWidth == TARGET_WIDTH and m.PelsHeight == TARGET_HEIGHT:
                    target_mode = m
                    break
                j += 1
            except: break
        
        if not target_mode: return False
        curr = win32api.EnumDisplaySettings(dev_name, win32con.ENUM_CURRENT_SETTINGS)
        if curr.PelsWidth == TARGET_WIDTH and curr.PelsHeight == TARGET_HEIGHT:
            return True
        target_mode.Fields = win32con.DM_PELSWIDTH | win32con.DM_PELSHEIGHT
        return win32api.ChangeDisplaySettingsEx(dev_name, target_mode) == win32con.DISP_CHANGE_SUCCESSFUL

# --- メッシュ生成 ---
def create_mesh(radius, arc_deg, aspect, segs):
    verts, inds = [], []
    w_arc = radius * math.radians(arc_deg)
    h = w_arc / aspect
    half_ang = math.radians(arc_deg) / 2.0
    for i in range(segs + 1):
        t = i / segs
        theta = -half_ang + (t * 2 * half_ang)
        x = radius * math.sin(theta)
        z = -radius * math.cos(theta)
        verts.extend([x, h/2, z, t, 0.0])
        verts.extend([x, -h/2, z, t, 1.0])
    for i in range(segs):
        base = i * 2
        inds.extend([base, base+1, base+2, base+2, base+1, base+3])
    return np.array(verts, dtype='f4'), np.array(inds, dtype='i4')

# --- FPS表示クラス ---
class FpsOverlay:
    def __init__(self, ctx):
        self.ctx = ctx
        self.prog = ctx.program(vertex_shader=UI_VERTEX_SHADER, fragment_shader=UI_FRAGMENT_SHADER)
        self.font = pygame.font.SysFont("Arial", 24, bold=True)
        self.tex_w, self.tex_h = 150, 50
        self.texture = self.ctx.texture((self.tex_w, self.tex_h), 4)
        self.texture.filter = (moderngl.LINEAR, moderngl.LINEAR)
        
        # 右上配置
        vertices = np.array([
            0.85, 0.95, 0.0, 1.0,
            0.85, 0.85, 0.0, 0.0,
            0.98, 0.85, 1.0, 0.0,
            0.85, 0.95, 0.0, 1.0,
            0.98, 0.85, 1.0, 0.0,
            0.98, 0.95, 1.0, 1.0,
        ], dtype='f4')
        self.vbo = self.ctx.buffer(vertices)
        self.vao = self.ctx.vertex_array(self.prog, [(self.vbo, '2f 2f', 'in_pos', 'in_uv')])
        self.surface = pygame.Surface((self.tex_w, self.tex_h), pygame.SRCALPHA)
        self.last_update = 0

    def render(self, fps_val):
        now = time.time()
        if now - self.last_update > 0.2:
            self.surface.fill((0,0,0,0))
            pygame.draw.rect(self.surface, (0, 0, 0, 150), (0, 0, self.tex_w, self.tex_h), border_radius=5)
            text = self.font.render(f"FPS: {int(fps_val)}", True, (0, 255, 0))
            rect = text.get_rect(center=(self.tex_w//2, self.tex_h//2))
            self.surface.blit(text, rect)
            self.texture.write(pygame.image.tostring(self.surface, 'RGBA', True))
            self.last_update = now
        
        self.ctx.disable(moderngl.DEPTH_TEST)
        self.texture.use(0)
        self.vao.render()
        self.ctx.enable(moderngl.DEPTH_TEST)

# --- メインビュワー ---
def run_viewer():
    pygame.init()
    pygame.font.init()
    
    # パフォーマンス設定
    pygame.display.gl_set_attribute(pygame.GL_MULTISAMPLEBUFFERS, 0)
    pygame.display.gl_set_attribute(pygame.GL_MULTISAMPLESAMPLES, 0)
    pygame.display.gl_set_attribute(pygame.GL_SWAP_CONTROL, 0)
    
    pygame.display.set_mode((1280, 720), DOUBLEBUF | OPENGL | RESIZABLE)
    pygame.display.set_caption("Ultra-Wide Glass")
    
    ctx = moderngl.create_context()
    ctx.enable(moderngl.DEPTH_TEST)
    prog = ctx.program(vertex_shader=VERTEX_SHADER, fragment_shader=FRAGMENT_SHADER)
    
    fps_overlay = FpsOverlay(ctx)

    # --- モニター探索ロジック ---
    print(">>> ゲーミングノート対応: 全GPUをスキャンします...")
    
    # 【修正】DXCamの初期化前に、システム全体の安定を待つ
    print(">>> ディスプレイ認識安定待ち(5秒)...")
    time.sleep(5.0)

    camera = None
    found_gpu = False
    
    real_monitor_exists = False
    try:
        for i in range(10):
            try:
                dev = win32api.EnumDisplayDevices(None, i-1)
                settings = win32api.EnumDisplaySettings(dev.DeviceName, win32con.ENUM_CURRENT_SETTINGS)
                if settings.PelsWidth == 6000:
                    print(f">>> [Windows認識] 仮想ディスプレイ発見: {dev.DeviceString} ({settings.PelsWidth}x{settings.PelsHeight})")
                    real_monitor_exists = True
                    break
            except: break
    except: pass
    if not real_monitor_exists: print("!!! 警告: Windows自体が6000pxのモニターを認識していません。")

    for device_idx in range(4): 
        if found_gpu: break
        try:
            try:
                _dummy = dxcam.create(device_idx=device_idx, output_idx=0)
                del _dummy
            except: continue

            for out_idx in range(6):
                try:
                    cam_check = dxcam.create(device_idx=device_idx, output_idx=out_idx)
                    w, h = cam_check.width, cam_check.height
                    print(f"    GPU[{device_idx}] Output[{out_idx}]: {w}x{h}")
                    if w >= 5900:
                        print(f">>> 発見成功！ GPU[{device_idx}] Output[{out_idx}] を使用します。")
                        camera = cam_check
                        found_gpu = True
                        break
                    cam_check.release()
                    del cam_check
                except: continue
        except: pass

    if not camera:
        print("\n" + "="*60)
        print("【重要】モニターが見つかりません。")
        print("※起動タイミングの問題かもしれないので、もう一度実行してみてください。")
        print("="*60 + "\n")
        input("Enterキーを押して終了..."); sys.exit()

    # キャプチャ開始
    camera.start(target_fps=144, video_mode=True)
    cw, ch = camera.width, camera.height
    
    texture = ctx.texture((cw, ch), 3)
    texture.filter = (moderngl.LINEAR, moderngl.LINEAR)
    
    def update_mesh():
        vbo_d, ibo_d = create_mesh(VIEWER_CONFIG['RADIUS'], VIEWER_CONFIG['ARC_ANGLE'], cw/ch, VIEWER_CONFIG['SEGMENTS'])
        vbo = ctx.buffer(vbo_d); ibo = ctx.buffer(ibo_d)
        vao = ctx.vertex_array(prog, [(vbo, '3f 2f', 'in_position', 'in_texcoord')], ibo)
        return vbo, ibo, vao

    vbo, ibo, vao = update_mesh()

    # IMU接続
    h_imu = None
    try:
        h_imu = hid.device()
        path = next(d['path'] for d in hid.enumerate(VENDOR_ID, PRODUCT_ID) if d['interface_number'] == 0)
        h_imu.open_path(path)
        h_imu.set_nonblocking(1)
        print("IMU Connected.")
    except: pass

    base_q = glm.quat(1,0,0,0)
    curr_q = glm.quat(1,0,0,0)
    clock = pygame.time.Clock()
    running = True
    need_reset = True 
    is_fullscreen = False

    print("\n=== 操作ガイド ===")
    print(" [SPACE]        : 視点リセット")
    print(" [Alt + ↑/↓]   : 拡大 / 縮小 (FOV変更)")
    print(" [Alt + ←/→]   : 湾曲率の調整")
    print(" [ESC]          : 終了")
    print("==================")

    while running:
        dt = clock.tick(144) / 1000.0
        fps_val = clock.get_fps()
        
        for event in pygame.event.get():
            if event.type == QUIT: running = False
            if event.type == VIDEORESIZE: ctx.viewport = (0,0,event.w,event.h)
            if event.type == KEYDOWN:
                if event.key == K_ESCAPE: running = False
                if event.key == K_SPACE:
                    base_q = curr_q; print("Reset View")
                if event.key == K_F11:
                    is_fullscreen = not is_fullscreen
                    if is_fullscreen: pygame.display.toggle_fullscreen()
                    else: pygame.display.toggle_fullscreen()

        # --- 入力処理 ---
        rebuild = False
        is_alt = win32api.GetAsyncKeyState(win32con.VK_MENU) & 0x8000
        if is_alt:
            # Zoom In (FOVを小さく)
            if win32api.GetAsyncKeyState(win32con.VK_UP) & 0x8000:
                VIEWER_CONFIG['FOV'] = max(10.0, VIEWER_CONFIG['FOV'] - 30.0 * dt)
            # Zoom Out (FOVを大きく)
            if win32api.GetAsyncKeyState(win32con.VK_DOWN) & 0x8000:
                VIEWER_CONFIG['FOV'] = min(120.0, VIEWER_CONFIG['FOV'] + 30.0 * dt)
                
            # Curve (湾曲)
            if win32api.GetAsyncKeyState(win32con.VK_LEFT) & 0x8000:
                VIEWER_CONFIG['ARC_ANGLE'] = max(10.0, VIEWER_CONFIG['ARC_ANGLE'] - 30.0 * dt); rebuild = True
            if win32api.GetAsyncKeyState(win32con.VK_RIGHT) & 0x8000:
                VIEWER_CONFIG['ARC_ANGLE'] += 30.0 * dt; rebuild = True
        
        if rebuild:
            vbo.release(); ibo.release(); vao.release()
            vbo, ibo, vao = update_mesh()

        # IMU
        if h_imu:
            try:
                last_d = None
                while True:
                    d = h_imu.read(64)
                    if not d: break
                    last_d = d
                if last_d:
                    v = struct.unpack('>iiii', bytearray(last_d[4:20]))
                    norm = math.sqrt(sum(x*x for x in v))
                    if norm > 0:
                        w, x, y, z = v[0]/norm, v[1]/norm, v[2]/norm, v[3]/norm
                        if not INVERT_PITCH: x = -x
                        if INVERT_YAW: y = -y
                        if INVERT_ROLL: z = -z
                        raw_q = glm.quat(w, x, y, z)
                        if need_reset: base_q = raw_q; need_reset = False
                        curr_q = glm.slerp(curr_q, raw_q, 0.15)
            except: pass

        # カメラ更新
        try:
            img = camera.get_latest_frame() 
            if img is not None:
                texture.write(img.tobytes())
        except: pass

        ctx.clear(0.0, 0.0, 0.0)
        
        view_q = glm.inverse(base_q) * curr_q
        view_rot = glm.mat4_cast(glm.inverse(view_q))
        prog['m_view'].write(glm.translate(glm.vec3(0,0,0)) * view_rot)
        
        # FOV適用 (Zoom)
        aspect_ratio = pygame.display.get_surface().get_width()/pygame.display.get_surface().get_height()
        prog['m_proj'].write(glm.perspective(glm.radians(VIEWER_CONFIG['FOV']), aspect_ratio, 0.1, 100.0))
        
        prog['m_model'].write(glm.mat4(1.0))
        texture.use(0)
        vao.render()
        
        # FPS描画
        fps_overlay.render(fps_val)

        pygame.display.flip()

    try: camera.stop(); camera.release()
    except: pass
    if h_imu: h_imu.close()
    try: ctypes.windll.winmm.timeEndPeriod(1)
    except: pass
    pygame.quit()

if __name__ == '__main__':
    if not is_admin():
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
        sys.exit()

    vm = VirtualDisplayManager(DRIVER_PATH, INSTALLER_EXE)
    try:
        vm.enable()
        run_viewer()
    except KeyboardInterrupt: pass
    except Exception as e:
        traceback.print_exc()
        input("Critical Error. Press Enter...")
    finally:
        
        vm.disable()