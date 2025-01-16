import os
import sys
import ctypes
import time
import threading
import tkinter as tk
from tkinter import filedialog
import pydirectinput
from mido import MidiFile
from pynput.keyboard import GlobalHotKeys

def check_admin():
    try:
        is_admin = (os.getuid() == 0)
    except AttributeError:
        is_admin = (ctypes.windll.shell32.IsUserAnAdmin() != 0)
    if not is_admin:
        print("当前非管理员权限，尝试提权...")
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
        sys.exit(0)

KEY_MAP = {
    72: '1',
    74: '2',
    76: '3',
    77: '4',
    79: '5',
    81: '6',
    83: '7',
    60: 'q',
    62: 'w',
    64: 'e',
    65: 'f',
    67: 't',
    69: 'y',
    71: 'u',
    48: 'a',
    50: 's',
    52: 'd',
    53: 'f',
    55: 'g',
    57: 'h',
    59: 'j'
}
WHITE_KEYS = sorted(KEY_MAP.keys())

def get_closest_white_key(pitch: int) -> int:
    if not WHITE_KEYS:
        return None
    if pitch < WHITE_KEYS[0]:
        return WHITE_KEYS[0]
    if pitch > WHITE_KEYS[-1]:
        return WHITE_KEYS[-1]
    closest = WHITE_KEYS[0]
    min_diff = abs(pitch - closest)
    for w in WHITE_KEYS:
        diff = abs(pitch - w)
        if diff < min_diff:
            closest = w
            min_diff = diff
        elif diff == min_diff:
            if w < closest:
                closest = w
                min_diff = diff
    return closest

def get_key_for_pitch(pitch: int) -> str:
    if pitch in KEY_MAP:
        return KEY_MAP[pitch]
    cwk = get_closest_white_key(pitch)
    if cwk is not None:
        return KEY_MAP.get(cwk, None)
    return None

def parse_midi_all_tempo(midi_path: str):
    mid = MidiFile(midi_path)
    all_events = []
    for track in mid.tracks:
        abs_tick = 0
        for msg in track:
            abs_tick += msg.time
            if msg.type == 'set_tempo':
                all_events.append((abs_tick, 'set_tempo', None, None, msg.tempo))
            elif msg.type in ('note_on', 'note_off'):
                all_events.append((abs_tick, msg.type, msg.note, msg.velocity, None))
            else:
                pass
    all_events.sort(key=lambda e: e[0])
    results = []
    current_time = 0.0
    prev_tick = 0
    current_tempo = 500000
    from mido import tick2second
    for i, (abs_tick, etype, note, velocity, tempo_val) in enumerate(all_events):
        delta_tick = abs_tick - prev_tick
        if delta_tick < 0:
            delta_tick = 0
        if delta_tick > 0:
            delta_seconds = tick2second(delta_tick, mid.ticks_per_beat, current_tempo)
            current_time += delta_seconds
        prev_tick = abs_tick
        if etype == 'set_tempo':
            current_tempo = tempo_val
        else:
            results.append((current_time, etype, note, velocity))
    results.sort(key=lambda e: e[0])
    return results

def post_process_events(events, chord_min_interval=0.01, note_min_interval=0.03):
    if not events:
        return []
    new_events = []
    i = 0
    n = len(events)
    while i < n:
        base_time, base_type, base_note, base_vel = events[i]
        chord_group = [(base_time, base_type, base_note, base_vel)]
        j = i + 1
        while j < n and abs(events[j][0] - base_time) < 1e-9:
            chord_group.append(events[j])
            j += 1
        chord_group.sort(key=lambda x: x[2])
        for k, (t, etype, note, vel) in enumerate(chord_group):
            new_t = base_time + k * chord_min_interval
            new_events.append((new_t, etype, note, vel))
        i = j
    new_events.sort(key=lambda x: x[0])
    adjusted = [new_events[0]]
    for idx in range(1, len(new_events)):
        p_time, p_type, p_note, p_vel = adjusted[-1]
        c_time, c_type, c_note, c_vel = new_events[idx]
        if c_time < p_time + note_min_interval:
            c_time = p_time + note_min_interval
        adjusted.append((c_time, c_type, c_note, c_vel))
    adjusted.sort(key=lambda x: x[0])
    return adjusted

class MidiPlayerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("MIDI Player (多轨Tempo + 后台解析) from:瑶光沁雪-泪冠哀歌，本程序免费发布，如有倒卖请及时退款。")
        self.events = []
        self.play_thread = None
        self.stop_flag = False
        self.parse_thread = None
        frame_top = tk.Frame(self)
        frame_top.pack(pady=5)
        self.btn_select = tk.Button(frame_top, text="选择 MIDI 文件", command=self.on_select_file)
        self.btn_select.pack(side=tk.LEFT, padx=5)
        self.label_file = tk.Label(frame_top, text="未选择文件", width=35, anchor='w')
        self.label_file.pack(side=tk.LEFT, padx=5)
        frame_intervals = tk.Frame(self)
        frame_intervals.pack(pady=5)
        tk.Label(frame_intervals, text="和弦最小间隔 (秒):").pack(side=tk.LEFT, padx=2)
        self.entry_chord = tk.Entry(frame_intervals, width=5)
        self.entry_chord.pack(side=tk.LEFT, padx=2)
        self.entry_chord.insert(0, "0")
        tk.Label(frame_intervals, text="音符最小间隔 (秒)").pack(side=tk.LEFT, padx=2)
        self.entry_note = tk.Entry(frame_intervals, width=5)
        self.entry_note.pack(side=tk.LEFT, padx=2)
        self.entry_note.insert(0, "0")
        self.label_status = tk.Label(self, text="准备就绪")
        self.label_status.pack(pady=5)
        frame_play = tk.Frame(self)
        frame_play.pack(pady=5)
        self.btn_start = tk.Button(frame_play, text="开始播放 (Ctrl+F10)", command=self.start_play)
        self.btn_start.pack(side=tk.LEFT, padx=10)
        self.btn_stop = tk.Button(frame_play, text="停止播放 (Ctrl+F11)", command=self.stop_play)
        self.btn_stop.pack(side=tk.LEFT, padx=10)
        self.label_author = tk.Label(self, text="瑶光沁雪-泪冠哀歌，本程序免费发布，如遇倒卖请及时退款。")
        self.label_author.pack(side=tk.BOTTOM, pady=5)
        self.global_hotkey_listener = None
        self.start_global_hotkeys()

    def on_select_file(self):
        self.stop_play()
        self.events = []
        file_path = filedialog.askopenfilename(title="选择 MIDI 文件", filetypes=[("MIDI files", "*.mid *.midi"), ("All files", "*.*")])
        if file_path:
            self.label_file.config(text=file_path)
            self.label_status.config(text="后台解析中，请稍候...")
            if self.parse_thread and self.parse_thread.is_alive():
                self.label_status.config(text="等待上一个解析完成...")
                return
            self.parse_thread = threading.Thread(target=self.parse_midi_in_background, args=(file_path,), daemon=True)
            self.parse_thread.start()
        else:
            self.label_file.config(text="未选择文件")
            self.label_status.config(text="未选择文件")

    def parse_midi_in_background(self, file_path: str):
        raw_events = parse_midi_all_tempo(file_path)
        try:
            chord_val = float(self.entry_chord.get())
        except ValueError:
            chord_val = 0.01
        try:
            note_val = float(self.entry_note.get())
        except ValueError:
            note_val = 0.03
        processed = post_process_events(raw_events, chord_val, note_val)
        self.events = processed
        self.after(0, self.on_parse_done)

    def on_parse_done(self):
        self.label_status.config(text=f"解析完成，共 {len(self.events)} 个事件")

    def start_play(self):
        if not self.events:
            self.label_status.config(text="没有可播放的事件，请先加载并解析MIDI")
            return
        self.stop_play()
        self.stop_flag = False
        self.play_thread = threading.Thread(target=self.play_midi_events, daemon=True)
        self.play_thread.start()
        self.label_status.config(text="正在播放...")

    def stop_play(self):
        self.stop_flag = True
        if self.play_thread and self.play_thread.is_alive():
            self.play_thread.join()
        self.play_thread = None
        self.label_status.config(text="播放已停止")

    def play_midi_events(self):
        start_time = time.time()
        idx = 0
        total = len(self.events)
        while idx < total and not self.stop_flag:
            event_time, event_type, pitch, velocity = self.events[idx]
            elapsed = time.time() - start_time
            if elapsed >= event_time:
                key_to_press = get_key_for_pitch(pitch)
                if key_to_press is not None:
                    if event_type == 'note_on' and velocity > 0:
                        pydirectinput.keyDown(key_to_press)
                    else:
                        pydirectinput.keyUp(key_to_press)
                idx += 1
            else:
                time.sleep(0.001)
        pressed_pitches = set(e[2] for e in self.events)
        for p in pressed_pitches:
            k = get_key_for_pitch(p)
            if k is not None:
                pydirectinput.keyUp(k)
        self.after(0, lambda: self.label_status.config(text="播放结束"))

    def start_global_hotkeys(self):
        def on_activate_start():
            self.start_play()
        def on_activate_stop():
            self.stop_play()
        def run_hotkeys():
            with GlobalHotKeys({
                '<ctrl>+<f10>': on_activate_start,
                '<ctrl>+<f11>': on_activate_stop
            }) as h:
                h.join()
        self.global_hotkey_listener = threading.Thread(target=run_hotkeys, daemon=True)
        self.global_hotkey_listener.start()

def main():
    check_admin()
    app = MidiPlayerApp()
    app.mainloop()

if __name__ == '__main__':
    main()
