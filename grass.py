import hid
import time
import struct
import math
import pyautogui

# --- 設定エリア ---
VENDOR_ID = 0x4817
PRODUCT_ID = 0x4242

# 動作範囲（マウス移動の感度）
RANGE_X_DEG = 25.0
RANGE_Y_DEG = 15.0
SMOOTHING = 0.2

# ■ クリックの感度（首をかしげる角度）
CLICK_THRESHOLD = 20.0  # 20度以上傾けるとクリック

# 反転設定
INVERT_X = True
INVERT_Y = False

pyautogui.FAILSAFE = False
# ------------------

def quaternion_to_euler(v0, v1, v2, v3):
    norm = math.sqrt(v0**2 + v1**2 + v2**2 + v3**2)
    if norm == 0: return 0, 0, 0
    w, x, y, z = v0/norm, v1/norm, v2/norm, v3/norm

    # Roll (赤: 上下 / Pitch)
    sinr_cosp = 2 * (w * x + y * z)
    cosr_cosp = 1 - 2 * (x * x + y * y)
    angle_pitch = math.degrees(math.atan2(sinr_cosp, cosr_cosp))

    # Yaw (青: 左右 / Yaw)
    siny_cosp = 2 * (w * z + x * y)
    cosy_cosp = 1 - 2 * (y * y + z * z)
    angle_yaw = math.degrees(math.atan2(siny_cosp, cosy_cosp))
    
    # ★追加: Tilt (緑: 首かしげ / Roll)
    # クォータニオンからの変換式
    sinp = 2 * (w * y - z * x)
    if abs(sinp) >= 1:
        angle_roll = math.copysign(90, sinp)
    else:
        angle_roll = math.degrees(math.asin(sinp))

    return angle_pitch, angle_yaw, angle_roll

# 画面サイズ
SCREEN_W, SCREEN_H = pyautogui.size()
CENTER_X = SCREEN_W / 2
CENTER_Y = SCREEN_H / 2

print(f"=== Head Mouse + Tilt Click ===")
print("Ctrl+C で終了")
print("初期化中... 正面を見て静止 (3秒)")

try:
    target_path = None
    for d in hid.enumerate(VENDOR_ID, PRODUCT_ID):
        if d['interface_number'] == 0:
            target_path = d['path']
            break
    
    h = hid.device()
    if target_path:
        h.open_path(target_path)
    else:
        h.open(VENDOR_ID, PRODUCT_ID)
    h.set_nonblocking(1)

    # キャリブレーション
    base_pitch = 0
    base_yaw = 0
    base_roll = 0
    samples = 0
    
    for _ in range(10): h.read(64)

    start_calib = time.time()
    while time.time() - start_calib < 3.0:
        data = h.read(64)
        if data and len(data) >= 20:
            vals = struct.unpack('>iiii', bytearray(data[4:20]))
            p, y, r = quaternion_to_euler(vals[0], vals[1], vals[2], vals[3])
            base_pitch += p
            base_yaw += y
            base_roll += r
            samples += 1
        time.sleep(0.01)
    
    if samples == 0: exit()

    base_pitch /= samples
    base_yaw /= samples
    base_roll /= samples
    
    print(">>> 操作スタート！ <<<")
    print("  - 右に首をかしげる: 左クリック")
    print("  - 左に首をかしげる: 右クリック")

    curr_cursor_x, curr_cursor_y = pyautogui.position()
    
    # クリック状態管理（連打防止）
    is_clicking = False

    while True:
        data = h.read(64)
        if data and len(data) >= 20:
            vals = struct.unpack('>iiii', bytearray(data[4:20]))
            raw_p, raw_y, raw_r = quaternion_to_euler(vals[0], vals[1], vals[2], vals[3])

            # 差分計算
            diff_pitch = raw_p - base_pitch
            diff_yaw = raw_y - base_yaw
            diff_roll = raw_r - base_roll # クリック判定用

            # 180度補正
            for d in [diff_pitch, diff_yaw, diff_roll]:
                if d > 180: d -= 360
                if d < -180: d += 360

            # --- マウス移動 ---
            if INVERT_X: diff_yaw *= -1
            if INVERT_Y: diff_pitch *= -1

            target_x = CENTER_X + (diff_yaw / RANGE_X_DEG) * (SCREEN_W / 2)
            target_y = CENTER_Y + (diff_pitch / RANGE_Y_DEG) * (SCREEN_H / 2)
            
            target_x = max(0, min(SCREEN_W - 1, target_x))
            target_y = max(0, min(SCREEN_H - 1, target_y))

            curr_cursor_x += (target_x - curr_cursor_x) * SMOOTHING
            curr_cursor_y += (target_y - curr_cursor_y) * SMOOTHING
            
            pyautogui.moveTo(curr_cursor_x, curr_cursor_y, duration=0)

            # --- クリック判定 ---
            # 右にかしげる (Roll > Threshold) -> 左クリック
            if diff_roll > CLICK_THRESHOLD:
                if not is_clicking:
                    pyautogui.click() # 左クリック
                    print("Left Click!")
                    is_clicking = True
            
            # 左にかしげる (Roll < -Threshold) -> 右クリック
            elif diff_roll < -CLICK_THRESHOLD:
                if not is_clicking:
                    pyautogui.rightClick() # 右クリック
                    print("Right Click!")
                    is_clicking = True
            
            # 頭を戻した (不感帯に戻った)
            elif abs(diff_roll) < (CLICK_THRESHOLD - 5):
                is_clicking = False

        time.sleep(0.01)

except KeyboardInterrupt:
    print("\n終了")
except Exception as e:
    print(f"\nエラー: {e}")
finally:
    try: h.close()
    except: pass