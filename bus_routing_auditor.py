#!/usr/bin/env python3
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json, os, csv, threading, time, re, socket

try:
    from waapi import WaapiClient, CannotConnectToWaapiException
    WAAPI_AVAILABLE = True
except ImportError:
    WAAPI_AVAILABLE = False

VERSION     = "V.1.0.0"
WAAPI_URL   = "ws://127.0.0.1:8080/waapi"
SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(SCRIPT_DIR, "bus_routing_rules.json")
WORD_SEP    = re.compile(r'[ _\-\./\\()\[\]{},;:\s]+')
DEFAULT_CONFIG = {
    "name_rules": [
        {"keyword": "UI",    "expected_bus_keyword": "UI",    "case_sensitive": False},
        {"keyword": "Music", "expected_bus_keyword": "Music", "case_sensitive": False},
        {"keyword": "AMB",   "expected_bus_keyword": "AMB",   "case_sensitive": False},
        {"keyword": "VO",    "expected_bus_keyword": "VO",    "case_sensitive": False},
    ],
    "workunit_rules": [
        {"work_unit_keyword": "UI",    "expected_bus_keyword": "UI",    "case_sensitive": False},
        {"work_unit_keyword": "Music", "expected_bus_keyword": "Music", "case_sensitive": False},
        {"work_unit_keyword": "AMB",   "expected_bus_keyword": "AMB",   "case_sensitive": False},
        {"work_unit_keyword": "VO",    "expected_bus_keyword": "VO",    "case_sensitive": False},
    ],
    "flag_unset_bus": True,
}
STRINGS = {
    "ko": {
        "reconnect":      "⟳  재연결",
        "lang_toggle":    "EN",
        "tab1":           "  스캔 1  ·  에셋 이름 기준  ",
        "tab2":           "  스캔 2  ·  Work Unit 기준  ",
        "rule_title":     "룰 설정",
        "desc_name":      "이름에 키워드 포함 → 버스 이름에도 해당 키워드 있어야 함",
        "desc_wu":        "Work Unit / 경로에 키워드 포함 → 버스 이름에도 해당 키워드 있어야 함",
        "kw_name":        "에셋 이름 키워드",
        "kw_wu":          "Work Unit / 경로 키워드",
        "bus_kw_hdr":     "→  버스 이름 키워드",
        "case_hdr":       "대소문자",
        "flag_unset":     "버스 미설정(상속됨) 항목도 위반으로 표시",
        "add_rule":       "+ 룰 추가",
        "save_rules":     "룰 저장",
        "search_lbl":     "검색",
        "search_hint":    "단어 단위 검색 (예: UI  /  AMB Metal  /  PC Attack)",
        "inherited":      "↑ 상속됨",
        "overridden":     "◆ 오버라이드됨",
        "col_name":       "에셋 이름",
        "col_work_unit":  "Work Unit",
        "col_cur_bus":    "현재 버스",
        "col_trigger":    "위반 원인",
        "col_exp_kw":     "기대 버스 키워드",
        "col_path":       "전체 경로",
        "scan":           "▶  스캔 실행",
        "view_wwise":     "⊙  Wwise에서 보기",
        "reroute":        "⟲  일괄 재라우팅",
        "select_all":     "전체 선택",
        "export_csv":     "↓  CSV 내보내기",
        "connecting":     "Wwise 연결 중...",
        "connected_fmt":  "연결됨  ·  {}  ·  버스 {}개",
        "connect_fail":   "연결 실패  —  Project > User Preferences > WAAPI 활성화 필요  ({})",
        "no_waapi":       "waapi-client 미설치  →  install.bat 을 실행하거나  pip install waapi-client",
        "scanning":       "스캔 중...",
        "scan_done_name": "[이름 스캔 완료]  {}개 Sound  ·  위반 {}개",
        "scan_done_wu":   "[Work Unit 스캔 완료]  {}개 Sound  ·  위반 {}개",
        "scan_fail":      "스캔 실패: {}",
        "filter_applied": "필터 적용  ·  {} / {}개",
        "no_violations":  "총 {}개  ·  위반 없음 ✓",
        "violations":     "총 {}개  ·  위반 {}개",
        "no_wwise":       "Wwise에 연결되지 않았습니다.",
        "select_item":    "항목을 선택하세요.",
        "select_reroute": "재라우팅할 항목을 선택하세요.",
        "no_export":      "내보낼 결과가 없습니다.",
        "rules_saved":    "룰이 저장되었습니다.",
        "save_title":     "저장",
        "error_title":    "오류",
        "info_title":     "안내",
        "scan_error":     "스캔 오류",
        "reroute_title":  "버스 재라우팅",
        "reroute_fmt":    "{}개 항목을 재라우팅합니다.",
        "reroute_bus_sel":"각 키워드 그룹에 적용할 버스를 선택하세요:",
        "cancel":         "취소",
        "apply_reroute":  "재라우팅 적용",
        "reroute_done":   "{}개 재라우팅 완료.",
        "reroute_partial":"{} 완료  /  {} 실패\n\n{}",
        "partial_title":  "일부 실패",
        "complete_title": "완료",
        "csv_title":      "CSV로 내보내기",
        "csv_done":       "저장 완료:\n{}",
        "not_found":      "Wwise에서 오브젝트를 선택할 수 없었습니다.\nProject Explorer에서 수동으로 찾아주세요.",
        "bus_missing":    '버스 "{}" 없음',
        "search_all":     "전체",
        "save_reminder":  "⚠ 룰 변경 시 저장 후 스캔",
        "flag_unset_note": "(위반된 최상위 항목의 하위 에셋들이 모두 노출됩니다)",
        "tip_kw_wu": (
            "┌─ Work Unit / 경로 키워드 — 매칭 방식 ────────────────────\n"
            "│  입력한 키워드가 아래 항목 중 하나라도 포함되면 매칭됩니다:\n"
            "│    · Work Unit 이름\n"
            "│    · Work Unit 전체 경로\n"
            "│    · Sound의 상위 계층 경로 (Sound 이름 자신 제외)\n"
            "│\n"
            "│  ⚠ 부분 이름도 매칭됩니다 (단어 경계 기준)\n"
            "│  예) 키워드: VIDEO\n"
            "│      → VIDEO_Main ✓  /  VIDEO_Dialogue ✓  /  VIDEO ✓\n"
            "│      → VIDEOGAME ✗  (단어 경계 없음)\n"
            "│\n"
            "│  특정 Work Unit만 검색하려면 이름 전체를 키워드로 입력하세요.\n"
            "│  예) 'VIDEO_Main' 만 검색 → 키워드: VIDEO_Main\n"
            "└──────────────────────────────────────────────────────────"
        ),
        "tip_name": (
            "┌─ 스캔 1  에셋 이름 기준 ────────────────────────────────\n"
            "│  Sound 오브젝트의 이름을 검사합니다.\n"
            "│\n"
            "│  동작 방식\n"
            "│  이름에 '에셋 이름 키워드'가 단어 토큰으로 포함되어 있으면\n"
            "│  해당 Sound가 라우팅된 버스 이름에도 '버스 이름 키워드'가\n"
            "│  단어 토큰으로 포함되어야 합니다. 없으면 위반으로 표시합니다.\n"
            "│\n"
            "│  단어 토큰이란?\n"
            "│  _  (언더바)  /  Space  /  -  /  .  /  /  등으로 구분된 단어.\n"
            "│  'UI'는 UI_Click ✓, NPC_UI ✓ 에 매칭되지만\n"
            "│  BUILD ✗, QUIT ✗ 에는 매칭되지 않습니다.\n"
            "│  키워드가 'UI_Mix' 처럼 _ 포함 시 연속 토큰 매칭을 사용합니다.\n"
            "│\n"
            "│  예시 룰:  PC  →  PC\n"
            "│    PC_Attack_01  라우팅→  PC_Bus      ✓\n"
            "│    PC_Attack_01  라우팅→  NPC_Bus     ✗  (NPC ≠ PC)\n"
            "└──────────────────────────────────────────────────────────"
        ),
        "tip_wu": (
            "┌─ 스캔 2  Work Unit / 경로 기준 ─────────────────────────\n"
            "│  Sound의 소속 경로를 검사합니다.\n"
            "│  검색 대상: Work Unit 이름 + Work Unit 전체 경로\n"
            "│             + 오브젝트 상위 경로 (Sound 자신의 이름 제외)\n"
            "│\n"
            "│  예시 룰:  경로 키워드: UI  →  버스 이름 키워드: UI\n"
            "│    \\SFX\\UI\\Buttons\\Click_01  라우팅→  UI_Bus  ✓\n"
            "│    \\SFX\\UI\\Buttons\\Click_01  라우팅→  SFX_Bus ✗\n"
            "│\n"
            "│  ▸ 상속됨 필터  (노랑)\n"
            "│    버스가 부모 오브젝트에서 상속된 항목\n"
            "│  ▸ 오버라이드됨 필터  (빨강)\n"
            "│    버스가 이 Sound에 직접 명시적으로 설정된 항목\n"
            "└──────────────────────────────────────────────────────────"
        ),
    },
    "en": {
        "reconnect":      "⟳  Reconnect",
        "lang_toggle":    "KO",
        "tab1":           "  Scan 1  ·  Asset Name  ",
        "tab2":           "  Scan 2  ·  Work Unit  ",
        "rule_title":     "Rule Settings",
        "desc_name":      "If name contains keyword → bus name must also contain keyword",
        "desc_wu":        "If Work Unit / path contains keyword → bus name must also contain keyword",
        "kw_name":        "Asset Name Keyword",
        "kw_wu":          "Work Unit / Path Keyword",
        "bus_kw_hdr":     "→  Bus Name Keyword",
        "case_hdr":       "Case",
        "flag_unset":     "Also flag unset (inherited) bus as violation",
        "add_rule":       "+ Add Rule",
        "save_rules":     "Save Rules",
        "search_lbl":     "Search",
        "search_hint":    "Word search (e.g.: UI  /  AMB Metal  /  PC Attack)",
        "inherited":      "↑ Inherited",
        "overridden":     "◆ Overridden",
        "col_name":       "Asset Name",
        "col_work_unit":  "Work Unit",
        "col_cur_bus":    "Current Bus",
        "col_trigger":    "Violation Cause",
        "col_exp_kw":     "Expected Bus KW",
        "col_path":       "Full Path",
        "scan":           "▶  Run Scan",
        "view_wwise":     "⊙  View in Wwise",
        "reroute":        "⟲  Bulk Reroute",
        "select_all":     "Select All",
        "export_csv":     "↓  Export CSV",
        "connecting":     "Connecting to Wwise...",
        "connected_fmt":  "Connected  ·  {}  ·  {} buses",
        "connect_fail":   "Connection failed  —  Enable WAAPI in Project > User Preferences  ({})",
        "no_waapi":       "waapi-client not installed  →  run install.bat or  pip install waapi-client",
        "scanning":       "Scanning...",
        "scan_done_name": "[Name scan done]  {} Sounds  ·  {} violations",
        "scan_done_wu":   "[Work Unit scan done]  {} Sounds  ·  {} violations",
        "scan_fail":      "Scan failed: {}",
        "filter_applied": "Filtered  ·  {} / {}",
        "no_violations":  "{} total  ·  No violations ✓",
        "violations":     "{} total  ·  {} violations",
        "no_wwise":       "Not connected to Wwise.",
        "select_item":    "Please select an item.",
        "select_reroute": "Please select items to reroute.",
        "no_export":      "No results to export.",
        "rules_saved":    "Rules saved.",
        "save_title":     "Saved",
        "error_title":    "Error",
        "info_title":     "Info",
        "scan_error":     "Scan Error",
        "reroute_title":  "Bus Reroute",
        "reroute_fmt":    "Rerouting {} items.",
        "reroute_bus_sel":"Select target bus for each keyword group:",
        "cancel":         "Cancel",
        "apply_reroute":  "Apply Reroute",
        "reroute_done":   "{} items rerouted.",
        "reroute_partial":"{} done  /  {} failed\n\n{}",
        "partial_title":  "Partial Failure",
        "complete_title": "Complete",
        "csv_title":      "Export as CSV",
        "csv_done":       "Saved:\n{}",
        "not_found":      "Could not select object in Wwise.\nPlease find it manually in Project Explorer.",
        "bus_missing":    'Bus "{}" not found',
        "search_all":     "All",
        "save_reminder":  "⚠ Save rules before scanning",
        "flag_unset_note": "(All sub-assets under the violating top-level item are also shown)",
        "tip_kw_wu": (
            "┌─ Work Unit / Path Keyword — How matching works ─────────\n"
            "│  A keyword matches if it appears in any of the following:\n"
            "│    · Work Unit name\n"
            "│    · Work Unit full path\n"
            "│    · Sound's parent hierarchy path (Sound's own name excluded)\n"
            "│\n"
            "│  ⚠ Partial names also match (at word boundaries)\n"
            "│  e.g. keyword: VIDEO\n"
            "│      → VIDEO_Main ✓  /  VIDEO_Dialogue ✓  /  VIDEO ✓\n"
            "│      → VIDEOGAME ✗  (no word boundary)\n"
            "│\n"
            "│  To target a specific Work Unit, enter its full name.\n"
            "│  e.g. search only 'VIDEO_Main' → keyword: VIDEO_Main\n"
            "└──────────────────────────────────────────────────────────"
        ),
        "tip_name": (
            "┌─ Scan 1  Asset Name ────────────────────────────────────\n"
            "│  Inspects the name of each Sound object.\n"
            "│\n"
            "│  How it works\n"
            "│  If the name contains an 'Asset Name Keyword' as a word token,\n"
            "│  the bus name must also contain the 'Bus Name Keyword'.\n"
            "│  A violation is flagged if it does not.\n"
            "│\n"
            "│  What is a word token?\n"
            "│  Words separated by _  /  Space  /  -  /  .  /  /  etc.\n"
            "│  'UI' matches UI_Click ✓, NPC_UI ✓\n"
            "│  but NOT BUILD ✗, QUIT ✗.\n"
            "│  Keywords containing '_' (e.g. 'UI_Mix') use consecutive token matching.\n"
            "│\n"
            "│  Example rule:  PC  →  PC\n"
            "│    PC_Attack_01  routed→  PC_Bus      ✓\n"
            "│    PC_Attack_01  routed→  NPC_Bus     ✗  (NPC ≠ PC)\n"
            "└──────────────────────────────────────────────────────────"
        ),
        "tip_wu": (
            "┌─ Scan 2  Work Unit / Path ──────────────────────────────\n"
            "│  Inspects the hierarchy path of each Sound object.\n"
            "│  Search targets: Work Unit name + Work Unit full path\n"
            "│                  + object parent path (excludes Sound's own name)\n"
            "│\n"
            "│  Example rule:  Path KW: UI  →  Bus KW: UI\n"
            "│    \\SFX\\UI\\Buttons\\Click_01  routed→  UI_Bus  ✓\n"
            "│    \\SFX\\UI\\Buttons\\Click_01  routed→  SFX_Bus ✗\n"
            "│\n"
            "│  ▸ Inherited filter  (yellow)\n"
            "│    Bus is inherited from a parent object\n"
            "│  ▸ Overridden filter  (red)\n"
            "│    Bus is explicitly set on this Sound\n"
            "└──────────────────────────────────────────────────────────"
        ),
    },
}

# ── Color System ─────────────────────────────────────────────────────────────
BG      = "#0D1117"   # main window bg
BG2     = "#161B22"   # card / rule panel
BG3     = "#1C2128"   # elevated panel / header areas
PANEL   = "#21262D"   # inputs, raised buttons
BORDER  = "#30363D"   # subtle borders
BORDER2 = "#444C56"   # hover-state borders
ACCENT  = "#58A6FF"   # primary blue
DANGER  = "#DA3633"   # danger
OK_CLR  = "#3FB950"   # success
WARN    = "#D29922"   # warning/amber
ERR_CLR = "#F85149"   # error red (bright)
FG      = "#E6EDF3"   # primary text
FG_DIM  = "#8B949E"   # secondary text
FG_MUT  = "#484F58"   # muted / placeholder
SEL_BG  = "#1F6FEB"   # treeview selection

# ── Typography ───────────────────────────────────────────────────────────────
_UI     = "Segoe UI"
_MN     = "Consolas"
FONT_H1    = (_UI, 11, "bold")
FONT_H2    = (_UI, 10, "bold")
FONT_UI    = (_UI,  9)
FONT_UIB   = (_UI,  9, "bold")
FONT_SM    = (_UI,  8)
FONT_CODE  = (_MN,  9)
FONT_CODE_B= (_MN,  9, "bold")
FONT_CODE_L= (_MN, 10, "bold")

# Legacy aliases — used directly in logic code, keep same names
MONO    = FONT_CODE
MONO_B  = FONT_CODE_B
MONO_LG = FONT_CODE_L

# ── Button animation presets  (bg, fg, hover_bg, press_bg) ───────────────────
_BP = {
    "primary": ("#1F6FEB", "#FFFFFF",  "#388BFD", "#1158C7"),
    "danger":  ("#3D0E11", "#F28B82",  "#5A1317", "#2A080A"),
    "ghost":   (PANEL,     FG,         "#2D333B", BG3),
    "add":     ("#0A2016", OK_CLR,     "#0F2E1E", "#071612"),
    "add_bus": ("#0A1830", ACCENT,     "#0E2244", "#071020"),
    "remove":  (PANEL,     FG_DIM,     "#3D1A1A", "#2A0F0F"),
    "lang":    ("#1E1833", "#BC8CFF",  "#271F42", "#140F24"),
    "icon":    (BG3,       FG_DIM,     PANEL,     BG2),
    "muted":   (BG2,       FG_DIM,     BG3,       BG),
}

FIND_CMD_PRIMARY   = ["FindInProjectExplorerSelectionChannel1","FindInProjectExplorer","FindInProjectExplorer1"]
SEARCH_FIELDS      = ["", "name", "work_unit", "current_bus", "expected_bus_keyword"]
SEARCH_FIELD_SKEYS = ["search_all", "col_name", "col_work_unit", "col_cur_bus", "col_exp_kw"]


# ── Animated button factory ───────────────────────────────────────────────────
def _ab(parent, text, cmd=None, preset="ghost", font=None, padx=14, pady=0, width=None, **kw):
    """Return an animated tk.Button with hover/press state transitions."""
    bg, fg, hov, prs = _BP[preset]
    extra = {"width": width} if width is not None else {}
    b = tk.Button(parent, text=text, command=cmd,
                  bg=bg, fg=fg, relief="flat", bd=0,
                  font=font or FONT_UIB, padx=padx, pady=pady + 6,
                  cursor="hand2", disabledforeground=FG_MUT,
                  activebackground=hov, activeforeground=fg,
                  **extra, **kw)
    def _enter(e):
        if str(b.cget("state")) != "disabled": b.config(bg=hov)
    def _leave(e):
        if str(b.cget("state")) != "disabled": b.config(bg=bg)
        else: b.config(bg=bg)
    def _press(e):
        if str(b.cget("state")) != "disabled": b.config(bg=prs)
    def _release(e):
        if str(b.cget("state")) != "disabled": b.config(bg=hov)
    b.bind("<Enter>",           _enter)
    b.bind("<Leave>",           _leave)
    b.bind("<Button-1>",        _press)
    b.bind("<ButtonRelease-1>", _release)
    return b


def _styled_entry(parent, var, width=16):
    """Entry with focus-glow animation."""
    e = tk.Entry(parent, textvariable=var,
                 bg=PANEL, fg=FG, insertbackground=FG,
                 width=width, font=FONT_CODE, relief="flat",
                 highlightthickness=1,
                 highlightbackground=BORDER,
                 highlightcolor=ACCENT)
    e.bind("<FocusIn>",  lambda ev: e.config(highlightbackground=ACCENT))
    e.bind("<FocusOut>", lambda ev: e.config(highlightbackground=BORDER))
    return e


class BusRoutingAuditor:
    def __init__(self, root):
        self.root = root
        self.root.title("Bus Routing Auditor  —  Wwise")
        self.root.geometry("1200x800")
        self.root.minsize(900, 620)
        self.root.configure(bg=BG)
        self.client = None
        self._was_connected = False
        self.results_name = []; self.results_wu = []
        self.filtered_name = []; self.filtered_wu = []
        self._scanned_name = False; self._scanned_wu = False
        self.buses = {}
        self._find_cmd = None
        self._sort_reverse = {}
        self._status_anim_id = None
        self.config = self._load_config()
        self._lang = "ko"
        self._lang_updaters = []
        self._nb = None; self._lang_btn = None
        self._search_field_idx_name = None; self._search_field_idx_wu = None
        self._apply_styles()
        self._build_ui()
        self._start_wwise_watchdog()
        self.root.after(200, lambda: threading.Thread(target=self._connect_waapi, daemon=True).start())

    # ── i18n ─────────────────────────────────────────────────────────────────
    def _t(self, key):
        return STRINGS[self._lang].get(key, STRINGS["ko"].get(key, key))

    def _toggle_lang(self):
        self._lang = "en" if self._lang == "ko" else "ko"
        self._refresh_lang()

    def _refresh_lang(self):
        for fn in self._lang_updaters:
            try: fn()
            except Exception: pass

    # ── Config ───────────────────────────────────────────────────────────────
    def _load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                for k, v in DEFAULT_CONFIG.items():
                    data.setdefault(k, v)
                return data
            except Exception: pass
        return json.loads(json.dumps(DEFAULT_CONFIG))

    def _save_config(self):
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(self.config, f, indent=2, ensure_ascii=False)

    # ── Watchdog ─────────────────────────────────────────────────────────────
    def _start_wwise_watchdog(self):
        def _watch():
            while True:
                time.sleep(3)
                if not self._was_connected: continue
                try:
                    s = socket.create_connection(("127.0.0.1", 8080), timeout=2); s.close()
                except Exception:
                    self.root.after(0, self.root.destroy); return
        threading.Thread(target=_watch, daemon=True).start()

    # ── WAAPI ─────────────────────────────────────────────────────────────────
    def _connect_waapi(self):
        self._set_status(self._t("connecting"), WARN, pulsing=True)
        if not WAAPI_AVAILABLE:
            self._set_status(self._t("no_waapi"), ERR_CLR); return
        try:
            if self.client:
                try: self.client.disconnect()
                except Exception: pass
            self.client = WaapiClient(url=WAAPI_URL)
            self._find_cmd = None
            self._fetch_buses()
            r = self.client.call("ak.wwise.core.object.get",
                {"from": {"ofType": ["Project"]}, "options": {"return": ["name"]}})
            items = r.get("return", [])
            proj_name = items[0]["name"] if items else "—"
            self._was_connected = True
            self._set_status(self._t("connected_fmt").format(proj_name, len(self.buses)), OK_CLR)
        except Exception as e:
            self.client = None
            self._set_status(self._t("connect_fail").format(e), ERR_CLR)

    def _fetch_buses(self):
        r = self.client.call("ak.wwise.core.object.get",
            {"from": {"ofType": ["Bus","AuxBus"]}, "options": {"return": ["id","name","path"]}})
        self.buses = {obj["name"]: obj["id"] for obj in r.get("return", [])}

    _HIERARCHY_TYPES = ["Sound","ActorMixer","BlendContainer","RandomSequenceContainer","SwitchContainer","WorkUnit","Folder"]

    def _get_all_sounds(self):
        r = self.client.call("ak.wwise.core.object.get", {
            "from": {"ofType": self._HIERARCHY_TYPES},
            "options": {"return": ["id","name","path","type","@OutputBus","@OverrideOutput","workunit"]},
        })
        all_objects = r.get("return", [])
        effective = self._resolve_effective_buses(all_objects)
        _DEFAULT = "masteraudiobus"
        sounds = []
        for obj in all_objects:
            if obj.get("type") != "Sound": continue
            path = obj.get("path", "")
            eff = effective.get(path) or {"name": "Master Audio Bus", "id": ""}
            if eff.get("name", "").replace(" ", "").lower() == _DEFAULT:
                own = obj.get("@OutputBus") or {}
                own_name = own.get("name", "")
                if own_name and own_name.replace(" ", "").lower() != _DEFAULT:
                    eff = own
            obj["_effective_bus_name"] = eff.get("name", "")
            obj["_effective_bus_id"]   = eff.get("id", "")
            obj["_bus_inherited"]      = not obj.get("@OverrideOutput", False)
            sounds.append(obj)
        return sounds

    def _resolve_effective_buses(self, all_objects):
        DEFAULT_BUS = "masteraudiobus"
        overrides = {}; stale = {}
        for obj in all_objects:
            bus = obj.get("@OutputBus") or {}
            bus_name = bus.get("name", "")
            if not bus_name: continue
            path = obj["path"]
            if obj.get("@OverrideOutput", False):
                overrides[path] = bus
            elif bus_name.replace(" ","").lower() != DEFAULT_BUS and obj.get("type") != "Sound":
                stale[path] = bus
        override_cache = {}

        def _find_override(path):
            chain = []; cur = path
            while True:
                if cur in overrides:
                    val = overrides[cur]; override_cache[cur] = val
                    for p in chain: override_cache[p] = val
                    return val
                if cur in override_cache:
                    val = override_cache[cur]
                    for p in chain: override_cache[p] = val
                    return val
                chain.append(cur)
                sep = cur.rfind("\\")
                if sep <= 0: break
                cur = cur[:sep]
            for p in chain: override_cache[p] = {}
            return {}

        def _find_stale_highest(path):
            cur = path; last_val = {}
            while True:
                if cur in stale: last_val = stale[cur]
                sep = cur.rfind("\\")
                if sep <= 0: break
                cur = cur[:sep]
            return last_val

        final = {}
        for obj in all_objects:
            path = obj["path"]
            bus = _find_override(path)
            if not bus: bus = _find_stale_highest(path)
            final[path] = bus
        return final

    def _select_in_wwise(self, object_id):
        if not self.client: return
        if self._find_cmd:
            try:
                self.client.call("ak.wwise.ui.commands.execute",
                    {"command": self._find_cmd, "objects": [object_id]})
            except Exception:
                self._find_cmd = None
        if not self._find_cmd:
            for cmd in FIND_CMD_PRIMARY:
                try:
                    self.client.call("ak.wwise.ui.commands.execute",
                        {"command": cmd, "objects": [object_id]})
                    self._find_cmd = cmd; break
                except Exception: continue
            else:
                self.root.after(0, lambda: messagebox.showwarning(
                    self._t("info_title"), self._t("not_found"), parent=self.root))
                return
        try:
            self.client.call("ak.wwise.ui.commands.execute",
                {"command": "Inspect", "objects": [object_id]})
        except Exception: pass

    def _set_output_bus(self, object_id, bus_id):
        self.client.call("ak.wwise.core.object.setReference",
            {"object": object_id, "reference": "OutputBus", "value": bus_id})

    @staticmethod
    def _word_match(text, keyword, cs):
        flags = 0 if cs else re.IGNORECASE
        sep = r'[ _\-\./\\()\[\]{},;:\s]'
        pat = r'(?:^|' + sep + r')' + re.escape(keyword) + r'(?:$|' + sep + r')'
        return bool(re.search(pat, text, flags))

    def _bus_display(self, bus_name, inherited):
        return (bus_name + "  ↑") if (bus_name and inherited) else bus_name

    def _check_name_rules(self, sounds):
        rules = self.config.get("name_rules", [])
        flag_unset = self.config.get("flag_unset_bus", True)
        violations = []
        for sound in sounds:
            name = sound.get("name",""); path = sound.get("path","")
            bus_name = sound.get("_effective_bus_name",""); bus_id = sound.get("_effective_bus_id","")
            inherited = sound.get("_bus_inherited", False)
            wu_obj = sound.get("workunit") or {}; wu_name = wu_obj.get("name","")
            for rule in rules:
                cs = rule.get("case_sensitive", False)
                all_src = [rule["keyword"]] + rule.get("extra_keywords", [])
                all_bus = [rule["expected_bus_keyword"]] + rule.get("extra_bus_keywords", [])
                if not any(self._word_match(name, sk, cs) for sk in all_src): continue
                no_bus = not bus_name
                mismatch = (not no_bus) and not any(self._word_match(bus_name, bk, cs) for bk in all_bus)
                if mismatch or (no_bus and flag_unset):
                    violations.append({"id": sound.get("id",""), "name": name, "path": path,
                        "work_unit": wu_name, "current_bus": self._bus_display(bus_name, inherited),
                        "current_bus_id": bus_id, "trigger": f'이름에 "{" | ".join(all_src)}" 포함',
                        "expected_bus_keyword": " | ".join(all_bus), "unset": no_bus, "inherited": inherited})
        return violations

    def _check_workunit_rules(self, sounds):
        rules = self.config.get("workunit_rules", [])
        flag_unset = self.config.get("flag_unset_bus", True)
        violations = []
        for sound in sounds:
            name = sound.get("name",""); path = sound.get("path","")
            bus_name = sound.get("_effective_bus_name",""); bus_id = sound.get("_effective_bus_id","")
            inherited = sound.get("_bus_inherited", False)
            wu_obj = sound.get("workunit") or {}; wu_name = wu_obj.get("name",""); wu_path = wu_obj.get("path","")
            parent_path = path[:path.rfind("\\")] if "\\" in path else ""
            search = f"{wu_name} {wu_path} {parent_path}"
            for rule in rules:
                cs = rule.get("case_sensitive", False)
                all_src = [rule["work_unit_keyword"]] + rule.get("extra_work_unit_keywords", [])
                all_kws = [rule["expected_bus_keyword"]] + rule.get("extra_bus_keywords", [])
                if not any(self._word_match(search, sk, cs) for sk in all_src): continue
                no_bus = not bus_name
                mismatch = (not no_bus) and not any(self._word_match(bus_name, ek, cs) for ek in all_kws)
                if mismatch or (no_bus and flag_unset):
                    violations.append({"id": sound.get("id",""), "name": name, "path": path,
                        "work_unit": wu_name, "current_bus": self._bus_display(bus_name, inherited),
                        "current_bus_id": bus_id, "trigger": f'경로에 "{" | ".join(all_src)}" 포함',
                        "expected_bus_keyword": " | ".join(all_kws), "unset": no_bus, "inherited": inherited})
        return violations

    def _run_scan(self, mode):
        if not self.client:
            messagebox.showerror(self._t("error_title"), self._t("no_wwise"), parent=self.root); return
        self._set_status(self._t("scanning"), WARN)
        self.btn_scan_name.config(state="disabled"); self.btn_scan_wu.config(state="disabled")
        def worker():
            try:
                sounds = self._get_all_sounds()
                if mode == "name":
                    self.results_name = self._check_name_rules(sounds)
                    self._scanned_name = True
                    results = self.results_name
                    self.root.after(0, lambda: self._apply_filter("name"))
                    self.root.after(0, lambda: self._set_count(self.lbl_count_name, len(sounds), len(results)))
                    self.root.after(0, lambda: self._set_status(
                        self._t("scan_done_name").format(len(sounds), len(results)), ERR_CLR if results else OK_CLR))
                else:
                    self.results_wu = self._check_workunit_rules(sounds)
                    self._scanned_wu = True
                    results = self.results_wu
                    self.root.after(0, lambda: self._apply_filter("workunit"))
                    self.root.after(0, lambda: self._set_count(self.lbl_count_wu, len(sounds), len(results)))
                    self.root.after(0, lambda: self._set_status(
                        self._t("scan_done_wu").format(len(sounds), len(results)), ERR_CLR if results else OK_CLR))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror(self._t("scan_error"), str(e), parent=self.root))
                self.root.after(0, lambda: self._set_status(self._t("scan_fail").format(e), ERR_CLR))
            finally:
                self.root.after(0, lambda: self.btn_scan_name.config(state="normal"))
                self.root.after(0, lambda: self.btn_scan_wu.config(state="normal"))
        threading.Thread(target=worker, daemon=True).start()

    def _apply_filter(self, mode):
        is_name = (mode == "name")
        src    = self.results_name if is_name else self.results_wu
        tree   = self.tree_name    if is_name else self.tree_wu
        search = (self._search_var_name if is_name else self._search_var_wu).get().strip()
        search_tokens = [t for t in WORD_SEP.split(search) if t]
        filtered = []
        for r in src:
            if not is_name:
                is_inh = r.get("inherited", True)
                if is_inh  and not self._wu_show_inherited.get(): continue
                if not is_inh and not self._wu_show_override.get(): continue
            if search_tokens:
                fv = self._search_field_idx_name if is_name else self._search_field_idx_wu
                field_key = SEARCH_FIELDS[fv.get() if fv else 0]
                if field_key:
                    target = r.get(field_key, "")
                else:
                    target = " ".join([r.get("name",""), r.get("work_unit",""), r.get("path",""),
                                       r.get("current_bus",""), r.get("expected_bus_keyword","")])
                hay = {t.upper() for t in WORD_SEP.split(target) if t}
                if not all(st.upper() in hay for st in search_tokens): continue
            filtered.append(r)
        tree.delete(*tree.get_children())
        if is_name: self.filtered_name = filtered
        else:       self.filtered_wu   = filtered
        for i, r in enumerate(filtered):
            row_base = "row_e" if i % 2 == 0 else "row_o"
            if r.get("unset"):   tags = (row_base, "unset")
            else:                tags = (row_base, "violation")
            tree.insert("", "end", values=(
                r["name"], r["work_unit"], r["current_bus"],
                r["trigger"], r["expected_bus_keyword"], r["path"],
            ), tags=tags)
        total = len(src); shown = len(filtered)
        scanned = self._scanned_name if is_name else self._scanned_wu
        lbl = self.lbl_count_name if is_name else self.lbl_count_wu
        if search_tokens or (not is_name and not (self._wu_show_inherited.get() and self._wu_show_override.get())):
            lbl.config(text=self._t("filter_applied").format(shown, total), fg=WARN if shown < total else FG_DIM)
        else:
            if not scanned: lbl.config(text="—", fg=FG_DIM)
            elif shown == 0:
                lbl.config(text=self._t("no_violations").format(total), fg=OK_CLR)
                tree.insert("", "end",
                    values=("✓  " + self._t("no_violations").format(total), "", "", "", "", ""),
                    tags=("ok_msg",))
            else: lbl.config(text=self._t("violations").format(total, shown), fg=ERR_CLR)

    # ── ttk Styles ───────────────────────────────────────────────────────────
    def _apply_styles(self):
        s = ttk.Style(); s.theme_use("clam")

        # Notebook / Tabs
        s.configure("TNotebook", background=BG, borderwidth=0, tabmargins=[0, 0, 0, 0])
        s.configure("TNotebook.Tab",
                    background=BG2, foreground=FG_DIM,
                    padding=[20, 8], font=FONT_UIB,
                    borderwidth=0)
        s.map("TNotebook.Tab",
              background=[("selected", BG3)],
              foreground=[("selected", FG)])

        # Treeview
        s.configure("Treeview",
                    background=BG2, foreground=FG,
                    fieldbackground=BG2,
                    rowheight=26,
                    font=FONT_CODE,
                    borderwidth=0,
                    relief="flat")
        s.configure("Treeview.Heading",
                    background=BG3, foreground=FG_DIM,
                    font=FONT_UIB, relief="flat",
                    padding=[8, 6])
        s.map("Treeview",
              background=[("selected", SEL_BG)],
              foreground=[("selected", "#FFFFFF")])
        s.map("Treeview.Heading",
              background=[("active", PANEL)])

        # Scrollbars
        s.configure("Vertical.TScrollbar",
                    background=PANEL, troughcolor=BG2,
                    arrowcolor=FG_DIM, borderwidth=0,
                    relief="flat", width=8)
        s.configure("Horizontal.TScrollbar",
                    background=PANEL, troughcolor=BG2,
                    arrowcolor=FG_DIM, borderwidth=0,
                    relief="flat", width=8)
        s.map("Vertical.TScrollbar",   background=[("active", BORDER2)])
        s.map("Horizontal.TScrollbar", background=[("active", BORDER2)])

        # Combobox
        s.configure("TCombobox",
                    fieldbackground=PANEL, background=BG3,
                    foreground=FG, selectbackground=PANEL,
                    selectforeground=FG, insertcolor=FG,
                    arrowcolor=FG_DIM, borderwidth=0,
                    relief="flat", padding=[6, 4])
        s.map("TCombobox",
              fieldbackground=[("readonly", PANEL), ("readonly !focus", PANEL)],
              foreground=[("readonly", FG), ("readonly !focus", FG)],
              selectbackground=[("readonly", PANEL), ("readonly !focus", PANEL)],
              selectforeground=[("readonly", FG), ("readonly !focus", FG)],
              arrowcolor=[("active", ACCENT), ("pressed", ACCENT)])

    # ── Main Window ───────────────────────────────────────────────────────────
    def _build_ui(self):
        # ── Header bar ──────────────────────────────────────────────────────
        hdr = tk.Frame(self.root, bg=BG3, height=52)
        hdr.pack(fill="x"); hdr.pack_propagate(False)

        # Left: app identity
        id_frame = tk.Frame(hdr, bg=BG3)
        id_frame.pack(side="left", padx=(16, 0), pady=0)
        tk.Label(id_frame, text="◈", bg=BG3, fg=ACCENT,
                 font=(_UI, 16)).pack(side="left", padx=(0, 8))
        tk.Label(id_frame, text="Bus Routing Auditor", bg=BG3, fg=FG,
                 font=FONT_H1).pack(side="left")
        tk.Label(id_frame, text=f"  {VERSION}", bg=BG3, fg=FG_MUT,
                 font=FONT_SM).pack(side="left", pady=(3, 0))

        # Right: buttons
        btn_area = tk.Frame(hdr, bg=BG3)
        btn_area.pack(side="right", padx=(0, 12))

        self._lang_btn = _ab(btn_area, self._t("lang_toggle"),
                             self._toggle_lang, preset="lang",
                             font=FONT_UIB, padx=12)
        self._lang_btn.pack(side="right", padx=(4, 0), pady=10)
        self._lang_updaters.append(lambda: self._lang_btn.config(text=self._t("lang_toggle")))

        btn_rc = _ab(btn_area, self._t("reconnect"),
                     lambda: threading.Thread(target=self._connect_waapi, daemon=True).start(),
                     preset="ghost", font=FONT_UI, padx=12)
        btn_rc.pack(side="right", padx=4, pady=10)
        self._lang_updaters.append(lambda w=btn_rc: w.config(text=self._t("reconnect")))

        # Center: status
        status_area = tk.Frame(hdr, bg=BG3)
        status_area.pack(side="left", fill="x", expand=True, padx=20)
        self._status_dot = tk.Label(status_area, text="●", bg=BG3, fg=FG_MUT,
                                    font=(_UI, 10))
        self._status_dot.pack(side="left", padx=(0, 6))
        self._status_lbl = tk.Label(status_area, text="초기화 중...", bg=BG3,
                                    fg=FG_DIM, font=FONT_UI, anchor="w")
        self._status_lbl.pack(side="left", fill="x")

        # Accent separator line
        tk.Frame(self.root, bg=ACCENT, height=2).pack(fill="x")

        # ── Notebook ────────────────────────────────────────────────────────
        nb = ttk.Notebook(self.root)
        nb.pack(fill="both", expand=True)
        self._nb = nb

        tab1 = tk.Frame(nb, bg=BG); nb.add(tab1, text=self._t("tab1"))
        self._lang_updaters.append(lambda: self._nb.tab(0, text=self._t("tab1")))
        self._build_tab(tab1, "name")

        tab2 = tk.Frame(nb, bg=BG); nb.add(tab2, text=self._t("tab2"))
        self._lang_updaters.append(lambda: self._nb.tab(1, text=self._t("tab2")))
        self._build_tab(tab2, "workunit")

    # ── Tab ──────────────────────────────────────────────────────────────────
    def _build_tab(self, parent, mode):
        is_name  = (mode == "name")
        rule_key = "name_rules" if is_name else "workunit_rules"
        reg = self._lang_updaters

        # ── Rule Panel ──────────────────────────────────────────────────────
        outer = tk.Frame(parent, bg=BORDER); outer.pack(fill="x", padx=12, pady=(10, 0))
        rp = tk.Frame(outer, bg=BG2); rp.pack(fill="both", padx=1, pady=(0, 1))

        # Top accent bar
        tk.Frame(rp, bg=ACCENT, height=2).pack(fill="x")

        # Panel header
        ph = tk.Frame(rp, bg=BG2); ph.pack(fill="x", padx=14, pady=(10, 6))

        lbl_rt = tk.Label(ph, text=self._t("rule_title"), bg=BG2, fg=FG, font=FONT_H2)
        lbl_rt.pack(side="left")
        reg.append(lambda w=lbl_rt: w.config(text=self._t("rule_title")))

        sep = tk.Frame(ph, bg=BORDER, width=1); sep.pack(side="left", fill="y", padx=12, pady=2)

        desc_key = "desc_name" if is_name else "desc_wu"
        lbl_desc = tk.Label(ph, text=self._t(desc_key), bg=BG2, fg=FG_DIM, font=FONT_UI)
        lbl_desc.pack(side="left")
        reg.append(lambda w=lbl_desc, k=desc_key: w.config(text=self._t(k)))

        tip_key = "tip_name" if is_name else "tip_wu"
        help_btn = tk.Label(ph, text=" ? ", bg=PANEL, fg=ACCENT, font=FONT_UIB,
                            cursor="question_arrow", relief="flat")
        help_btn.pack(side="left", padx=(8, 0))
        self._attach_tooltip(help_btn, lambda k=tip_key: self._t(k))

        # Column headers
        col_hdr = tk.Frame(rp, bg=BG3); col_hdr.pack(fill="x", padx=14, pady=(4, 2))
        kw_key = "kw_name" if is_name else "kw_wu"

        lbl_kw = tk.Label(col_hdr, text=self._t(kw_key), bg=BG3, fg=FG_MUT,
                          font=FONT_SM, width=27, anchor="w")
        lbl_kw.grid(row=0, column=0, padx=(4, 0))
        reg.append(lambda w=lbl_kw, k=kw_key: w.config(text=self._t(k)))

        if not is_name:
            kw_tip = tk.Label(col_hdr, text=" ? ", bg=PANEL, fg=ACCENT,
                              font=FONT_UIB, cursor="question_arrow", relief="flat")
            kw_tip.grid(row=0, column=1, padx=(2, 6))
            self._attach_tooltip(kw_tip, lambda: self._t("tip_kw_wu"))
        else:
            tk.Label(col_hdr, text="", bg=BG3, width=4).grid(row=0, column=1)

        lbl_bus_kw = tk.Label(col_hdr, text=self._t("bus_kw_hdr"), bg=BG3, fg=FG_MUT,
                              font=FONT_SM, width=22, anchor="w")
        lbl_bus_kw.grid(row=0, column=2, padx=4)
        reg.append(lambda w=lbl_bus_kw: w.config(text=self._t("bus_kw_hdr")))

        lbl_case = tk.Label(col_hdr, text=self._t("case_hdr"), bg=BG3, fg=FG_MUT,
                            font=FONT_SM, width=8, anchor="w")
        lbl_case.grid(row=0, column=3, padx=4)
        reg.append(lambda w=lbl_case: w.config(text=self._t("case_hdr")))

        # Rule rows container
        rows_frame = tk.Frame(rp, bg=BG2); rows_frame.pack(fill="x", padx=14, pady=(0, 4))
        rule_rows = []

        # ── add_row ──────────────────────────────────────────────────────────
        def add_row(kw="", bus="", cs=False, extra_src=None, extra_buses=None):
            row_frame = tk.Frame(rows_frame, bg=BG2, pady=1)
            row_frame.pack(fill="x")

            # Subtle separator line above each row
            tk.Frame(row_frame, bg=BORDER, height=1).pack(fill="x")

            kw_var  = tk.StringVar(value=kw)
            bus_var = tk.StringVar(value=bus)
            cs_var  = tk.BooleanVar(value=cs)
            src_extra_list = []; bus_extra_list = []

            g = tk.Frame(row_frame, bg=BG2); g.pack(fill="x", pady=2)

            # col 0: or-label placeholder (src side)
            # col 1: source entry
            # col 2: + button (src)
            # col 3: arrow
            # col 4: or-label placeholder (bus side)
            # col 5: bus entry
            # col 6: + button (bus)
            # col 7: case checkbox
            # col 8: × remove button

            tk.Label(g, bg=BG2, width=3).grid(row=0, column=0)
            _styled_entry(g, kw_var, 18).grid(row=0, column=1, padx=(0, 2), pady=2, ipady=3)

            src_add = _ab(g, "+", lambda: add_src_extra(), preset="add",
                          font=FONT_CODE_B, padx=6, pady=0, width=2)
            src_add.grid(row=0, column=2, padx=2)

            tk.Label(g, text="→", bg=BG2, fg=FG_MUT, font=FONT_UI).grid(
                row=0, column=3, padx=(18, 18))

            tk.Label(g, bg=BG2, width=3).grid(row=0, column=4)
            _styled_entry(g, bus_var, 18).grid(row=0, column=5, padx=(0, 2), pady=2, ipady=3)

            bus_add = _ab(g, "+", lambda: add_bus_extra(), preset="add_bus",
                          font=FONT_CODE_B, padx=6, pady=0, width=2)
            bus_add.grid(row=0, column=6, padx=2)

            # Case checkbox — custom styled
            cb_frame = tk.Frame(g, bg=BG2); cb_frame.grid(row=0, column=7, padx=10)
            tk.Checkbutton(cb_frame, variable=cs_var, bg=BG2,
                           selectcolor=PANEL, activebackground=BG2,
                           relief="flat", bd=0).pack()

            def remove():
                rule_rows.remove((kw_var, bus_var, cs_var, src_extra_list, bus_extra_list, row_frame))
                row_frame.destroy()

            rm_btn = _ab(g, "✕", remove, preset="remove",
                         font=FONT_CODE_B, padx=6, pady=0)
            rm_btn.grid(row=0, column=8, padx=(2, 0))

            # extra_slots for OR rows
            extra_slots = {}

            def _alloc_src_slot():
                for r in sorted(extra_slots):
                    if not extra_slots[r]["src"]: return r
                r = (max(extra_slots, default=0) + 1)
                extra_slots[r] = {"src": False, "bus": False}; return r

            def _alloc_bus_slot():
                for r in sorted(extra_slots):
                    if not extra_slots[r]["bus"]: return r
                r = (max(extra_slots, default=0) + 1)
                extra_slots[r] = {"src": False, "bus": False}; return r

            def add_src_extra(val=""):
                r = _alloc_src_slot(); extra_slots[r]["src"] = True
                evar = tk.StringVar(value=val)
                lbl = tk.Label(g, text="or", bg=BG2, fg=OK_CLR, font=FONT_UIB, width=3)
                lbl.grid(row=r, column=0, sticky="e")
                e = _styled_entry(g, evar, 18)
                e.grid(row=r, column=1, padx=(0, 2), pady=2, ipady=3)
                def rm():
                    lbl.destroy(); e.destroy(); rm_b.destroy()
                    extra_slots[r]["src"] = False
                    src_extra_list[:] = [(v, x) for v, x in src_extra_list if v is not evar]
                rm_b = _ab(g, "−", rm, preset="remove", font=FONT_CODE_B, padx=6, pady=0, width=2)
                rm_b.grid(row=r, column=2, padx=2, pady=2)
                src_extra_list.append((evar, None))

            def add_bus_extra(val=""):
                r = _alloc_bus_slot(); extra_slots[r]["bus"] = True
                evar = tk.StringVar(value=val)
                lbl = tk.Label(g, text="or", bg=BG2, fg=ACCENT, font=FONT_UIB, width=3)
                lbl.grid(row=r, column=4, sticky="e")
                e = _styled_entry(g, evar, 18)
                e.grid(row=r, column=5, padx=(0, 2), pady=2, ipady=3)
                def rm():
                    lbl.destroy(); e.destroy(); rm_b.destroy()
                    extra_slots[r]["bus"] = False
                    bus_extra_list[:] = [(v, x) for v, x in bus_extra_list if v is not evar]
                rm_b = _ab(g, "−", rm, preset="remove", font=FONT_CODE_B, padx=6, pady=0, width=2)
                rm_b.grid(row=r, column=6, padx=2, pady=2)
                bus_extra_list.append((evar, None))

            rule_rows.append((kw_var, bus_var, cs_var, src_extra_list, bus_extra_list, row_frame))
            for es in (extra_src   or []): add_src_extra(es)
            for eb in (extra_buses or []): add_bus_extra(eb)

        for rule in self.config.get(rule_key, []):
            extra_src   = rule.get("extra_keywords", []) if is_name else rule.get("extra_work_unit_keywords", [])
            extra_buses = rule.get("extra_bus_keywords", [])
            if is_name:
                add_row(rule.get("keyword",""), rule.get("expected_bus_keyword",""),
                        rule.get("case_sensitive", False), extra_src, extra_buses)
            else:
                add_row(rule.get("work_unit_keyword",""), rule.get("expected_bus_keyword",""),
                        rule.get("case_sensitive", False), extra_src, extra_buses)

        # ── Rule panel footer ───────────────────────────────────────────────
        tk.Frame(rp, bg=BORDER, height=1).pack(fill="x", padx=14)
        footer = tk.Frame(rp, bg=BG2); footer.pack(fill="x", padx=14, pady=(6, 10))

        # Left: flag unset checkbox + note
        flag_var = tk.BooleanVar(value=self.config.get("flag_unset_bus", True))
        def on_flag(): self.config["flag_unset_bus"] = flag_var.get()
        cb_flag = tk.Checkbutton(footer, text=self._t("flag_unset"),
                                 variable=flag_var, command=on_flag,
                                 bg=BG2, fg=FG_DIM, font=FONT_UI,
                                 selectcolor=PANEL, activebackground=BG2)
        cb_flag.pack(side="left")
        reg.append(lambda w=cb_flag: w.config(text=self._t("flag_unset")))

        lbl_fn = tk.Label(footer, text=self._t("flag_unset_note"),
                          bg=BG2, fg=FG_MUT, font=FONT_SM)
        lbl_fn.pack(side="left", padx=(2, 0))
        reg.append(lambda w=lbl_fn: w.config(text=self._t("flag_unset_note")))

        # Right: buttons
        def save_rules():
            rules = []
            for kw_v, bus_v, cs_v, src_el, bus_el, _ in rule_rows:
                kw, bus = kw_v.get().strip(), bus_v.get().strip()
                if kw and bus:
                    rule = {"expected_bus_keyword": bus, "case_sensitive": cs_v.get()}
                    src_extras = [ev.get().strip() for ev, _ in src_el if ev.get().strip()]
                    bus_extras  = [ev.get().strip() for ev, _ in bus_el if ev.get().strip()]
                    if is_name:
                        rule["keyword"] = kw
                        if src_extras: rule["extra_keywords"] = src_extras
                    else:
                        rule["work_unit_keyword"] = kw
                        if src_extras: rule["extra_work_unit_keywords"] = src_extras
                    if bus_extras: rule["extra_bus_keywords"] = bus_extras
                    rules.append(rule)
            self.config[rule_key] = rules
            self._save_config()
            messagebox.showinfo(self._t("save_title"), self._t("rules_saved"), parent=self.root)

        btn_save = _ab(footer, self._t("save_rules"), save_rules,
                       preset="primary", font=FONT_UIB, padx=14)
        btn_save.pack(side="right")
        reg.append(lambda w=btn_save: w.config(text=self._t("save_rules")))

        lbl_rem = tk.Label(footer, text=self._t("save_reminder"),
                           bg=BG2, fg=WARN, font=FONT_UI)
        lbl_rem.pack(side="right", padx=(0, 12))
        reg.append(lambda w=lbl_rem: w.config(text=self._t("save_reminder")))

        btn_add = _ab(footer, self._t("add_rule"), lambda: add_row(),
                      preset="muted", font=FONT_UI, padx=12)
        btn_add.pack(side="right", padx=(0, 6))
        reg.append(lambda w=btn_add: w.config(text=self._t("add_rule")))

        # ── Filter Bar ──────────────────────────────────────────────────────
        fb = tk.Frame(parent, bg=BG3); fb.pack(fill="x", padx=12, pady=(6, 0))

        # Left accent bar
        tk.Frame(fb, bg=ACCENT, width=2).pack(side="left", fill="y", padx=(0, 10))

        lbl_srch = tk.Label(fb, text=self._t("search_lbl"), bg=BG3, fg=ACCENT,
                            font=FONT_UIB, width=4, anchor="w")
        lbl_srch.pack(side="left", padx=(0, 6), pady=7)
        reg.append(lambda w=lbl_srch: w.config(text=self._t("search_lbl")))

        field_idx_var = tk.IntVar(value=0)
        field_combo = ttk.Combobox(fb, values=[self._t(k) for k in SEARCH_FIELD_SKEYS],
                                   state="readonly", font=FONT_UI, width=16)
        field_combo.current(0); field_combo.pack(side="left", padx=(0, 6), pady=7)
        field_combo.bind("<<ComboboxSelected>>",
            lambda e, v=field_idx_var, c=field_combo, m=mode:
                (v.set(c.current()), self._apply_filter(m)))

        def _refresh_fc(c=field_combo, v=field_idx_var):
            c.config(values=[self._t(k) for k in SEARCH_FIELD_SKEYS]); c.current(v.get())
        reg.append(_refresh_fc)

        if is_name: self._search_field_idx_name = field_idx_var
        else:       self._search_field_idx_wu   = field_idx_var

        search_var = tk.StringVar()
        search_entry = tk.Entry(fb, textvariable=search_var,
                                bg=PANEL, fg=FG, insertbackground=FG,
                                font=FONT_CODE, relief="flat",
                                highlightthickness=1,
                                highlightbackground=BORDER,
                                highlightcolor=ACCENT, width=34)
        search_entry.pack(side="left", padx=(0, 4), pady=7, ipady=4)
        search_entry.bind("<FocusIn>",  lambda e: search_entry.config(highlightbackground=ACCENT))
        search_entry.bind("<FocusOut>", lambda e: search_entry.config(highlightbackground=BORDER))

        clr_btn = _ab(fb, "×",
                      lambda: search_var.set(""),
                      preset="icon", font=(_UI, 12, "bold"), padx=8, pady=0)
        clr_btn.pack(side="left", padx=(0, 10))

        lbl_hint = tk.Label(fb, text=self._t("search_hint"), bg=BG3, fg=FG_MUT, font=FONT_SM)
        lbl_hint.pack(side="left")
        reg.append(lambda w=lbl_hint: w.config(text=self._t("search_hint")))

        if is_name:
            self._search_var_name = search_var
            search_var.trace_add("write", lambda *_: self._apply_filter("name"))
        else:
            self._search_var_wu = search_var
            search_var.trace_add("write", lambda *_: self._apply_filter("workunit"))
            tk.Frame(fb, bg=BORDER, width=1).pack(side="left", fill="y", padx=(12, 10), pady=6)
            self._wu_show_inherited = tk.BooleanVar(value=True)
            self._wu_show_override  = tk.BooleanVar(value=True)
            def _wu_toggle(*_): self._apply_filter("workunit")
            cb_inh = tk.Checkbutton(fb, text=self._t("inherited"),
                                    variable=self._wu_show_inherited, command=_wu_toggle,
                                    bg=BG3, fg=WARN, font=FONT_UIB,
                                    selectcolor=BG2, activebackground=BG3,
                                    activeforeground=WARN)
            cb_inh.pack(side="left", padx=(0, 4))
            reg.append(lambda w=cb_inh: w.config(text=self._t("inherited")))

            cb_ovr = tk.Checkbutton(fb, text=self._t("overridden"),
                                    variable=self._wu_show_override, command=_wu_toggle,
                                    bg=BG3, fg=ERR_CLR, font=FONT_UIB,
                                    selectcolor=BG2, activebackground=BG3,
                                    activeforeground=ERR_CLR)
            cb_ovr.pack(side="left", padx=4)
            reg.append(lambda w=cb_ovr: w.config(text=self._t("overridden")))

        # ── Treeview ────────────────────────────────────────────────────────
        cols     = ("name", "work_unit", "current_bus", "trigger", "expected_kw", "path")
        col_keys = ("col_name","col_work_unit","col_cur_bus","col_trigger","col_exp_kw","col_path")
        col_wids = (200, 130, 170, 240, 130, 1200)

        tbl = tk.Frame(parent, bg=BORDER, bd=0)
        tbl.pack(fill="both", expand=True, padx=12, pady=(4, 0))

        vsb = ttk.Scrollbar(tbl, orient="vertical")
        vsb.pack(side="right", fill="y")
        hsb = ttk.Scrollbar(tbl, orient="horizontal")
        hsb.pack(side="bottom", fill="x")

        tree = ttk.Treeview(tbl, columns=cols, show="headings",
                            yscrollcommand=vsb.set, xscrollcommand=hsb.set,
                            selectmode="extended")
        tree.pack(fill="both", expand=True)
        vsb.config(command=tree.yview); hsb.config(command=tree.xview)

        for col, key, w in zip(cols, col_keys, col_wids):
            tree.heading(col, text=self._t(key), command=lambda c=col, t=tree: self._sort(t, c))
            tree.column(col, width=w, minwidth=60, stretch=False)
            reg.append(lambda t=tree, c=col, k=key: t.heading(c, text=self._t(k)))

        # Row tags: alternating + state
        tree.tag_configure("row_e",   background=BG2,    foreground=FG)
        tree.tag_configure("row_o",   background="#131920", foreground=FG)
        tree.tag_configure("violation", foreground=ERR_CLR)
        tree.tag_configure("unset",     foreground=WARN)
        tree.tag_configure("ok_msg",    foreground=OK_CLR)

        tree.bind("<Double-1>", lambda e, t=tree, m=is_name: self._on_dbl(e, t, m))
        if is_name: self.tree_name = tree
        else:       self.tree_wu   = tree

        # ── Action Bar ──────────────────────────────────────────────────────
        bar = tk.Frame(parent, bg=BG3, height=48)
        bar.pack(fill="x", padx=12, pady=(3, 8))
        bar.pack_propagate(False)

        # Scan button — primary action
        scan_btn = _ab(bar, self._t("scan"),
                       lambda: self._run_scan(mode),
                       preset="primary", font=FONT_UIB, padx=18)
        scan_btn.pack(side="left", padx=(0, 2), pady=8)
        reg.append(lambda w=scan_btn: w.config(text=self._t("scan")))

        tk.Frame(bar, bg=BORDER, width=1).pack(side="left", fill="y", padx=10, pady=10)

        btn_view = _ab(bar, self._t("view_wwise"),
                       lambda t=tree, m=is_name: self._action_select(t, m),
                       preset="ghost", font=FONT_UI, padx=12)
        btn_view.pack(side="left", padx=2, pady=8)
        reg.append(lambda w=btn_view: w.config(text=self._t("view_wwise")))

        btn_fix = _ab(bar, self._t("reroute"),
                      lambda t=tree, m=is_name: self._action_fix(t, m),
                      preset="danger", font=FONT_UIB, padx=12)
        btn_fix.pack(side="left", padx=2, pady=8)
        reg.append(lambda w=btn_fix: w.config(text=self._t("reroute")))

        btn_sel = _ab(bar, self._t("select_all"),
                      lambda t=tree: t.selection_set(t.get_children()),
                      preset="ghost", font=FONT_UI, padx=12)
        btn_sel.pack(side="left", padx=2, pady=8)
        reg.append(lambda w=btn_sel: w.config(text=self._t("select_all")))

        # Right side
        tk.Frame(bar, bg=BORDER, width=1).pack(side="right", fill="y", padx=10, pady=10)

        btn_csv = _ab(bar, self._t("export_csv"),
                      lambda m=is_name: self._export_csv(m),
                      preset="ghost", font=FONT_UI, padx=12)
        btn_csv.pack(side="right", padx=2, pady=8)
        reg.append(lambda w=btn_csv: w.config(text=self._t("export_csv")))

        tk.Label(bar, text=VERSION, bg=BG3, fg=FG_MUT, font=FONT_SM).pack(
            side="right", padx=(0, 6))

        count_lbl = tk.Label(bar, text="—", bg=BG3, fg=FG_DIM, font=FONT_UI)
        count_lbl.pack(side="right", padx=12)

        if is_name: self.btn_scan_name = scan_btn;  self.lbl_count_name = count_lbl
        else:       self.btn_scan_wu   = scan_btn;  self.lbl_count_wu   = count_lbl

    # ── Sort ─────────────────────────────────────────────────────────────────
    def _sort(self, tree, col):
        rev  = self._sort_reverse.get(col, False)
        data = [(tree.set(iid, col), iid) for iid in tree.get_children("")]
        data.sort(key=lambda x: x[0].lower(), reverse=rev)
        for i, (_, iid) in enumerate(data): tree.move(iid, "", i)
        self._sort_reverse[col] = not rev
        raw = tree.heading(col)["text"].rstrip(" ↑↓")
        tree.heading(col, text=raw + (" ↑" if not rev else " ↓"))

    def _set_count(self, lbl, total, n):
        if n == 0: lbl.config(text=self._t("no_violations").format(total), fg=OK_CLR)
        else:      lbl.config(text=self._t("violations").format(total, n), fg=ERR_CLR)

    def _get_filtered(self, is_name):
        return self.filtered_name if is_name else self.filtered_wu

    def _on_dbl(self, event, tree, is_name):
        sel = tree.selection()
        if not sel or not self.client: return
        results = self._get_filtered(is_name); idx = tree.index(sel[0])
        if 0 <= idx < len(results):
            threading.Thread(target=self._select_in_wwise, args=(results[idx]["id"],), daemon=True).start()

    def _action_select(self, tree, is_name):
        sel = tree.selection()
        if not sel:
            messagebox.showinfo(self._t("info_title"), self._t("select_item"), parent=self.root); return
        if not self.client:
            messagebox.showerror(self._t("error_title"), self._t("no_wwise"), parent=self.root); return
        results = self._get_filtered(is_name); idx = tree.index(sel[0])
        if 0 <= idx < len(results):
            threading.Thread(target=self._select_in_wwise, args=(results[idx]["id"],), daemon=True).start()

    def _action_fix(self, tree, is_name):
        sel = tree.selection()
        if not sel:
            messagebox.showinfo(self._t("info_title"), self._t("select_reroute"), parent=self.root); return
        if not self.client:
            messagebox.showerror(self._t("error_title"), self._t("no_wwise"), parent=self.root); return
        results = self._get_filtered(is_name)
        to_fix = [results[tree.index(iid)] for iid in sel if 0 <= tree.index(iid) < len(results)]
        if to_fix: self._show_fix_dialog(to_fix, is_name)

    # ── Reroute Dialog ───────────────────────────────────────────────────────
    def _show_fix_dialog(self, violations, is_name):
        dlg = tk.Toplevel(self.root)
        dlg.title(self._t("reroute_title"))
        dlg.geometry("580x380"); dlg.configure(bg=BG); dlg.grab_set()
        dlg.resizable(False, False)

        # Header
        dh = tk.Frame(dlg, bg=BG3); dh.pack(fill="x")
        tk.Frame(dh, bg=DANGER, height=2).pack(fill="x")
        tk.Label(dh, text=self._t("reroute_title"), bg=BG3, fg=FG,
                 font=FONT_H1).pack(side="left", padx=18, pady=12)

        # Body
        body = tk.Frame(dlg, bg=BG); body.pack(fill="both", expand=True, padx=18, pady=(12, 6))
        tk.Label(body, text=self._t("reroute_fmt").format(len(violations)),
                 bg=BG, fg=FG, font=FONT_UIB).pack(anchor="w")
        tk.Label(body, text=self._t("reroute_bus_sel"),
                 bg=BG, fg=FG_DIM, font=FONT_UI).pack(anchor="w", pady=(4, 10))

        by_kw = {}
        for v in violations: by_kw.setdefault(v["expected_bus_keyword"], []).append(v)
        bus_names = sorted(self.buses.keys()); kw_bus_vars = {}

        grid = tk.Frame(body, bg=BG); grid.pack(fill="x")
        for i, (kw, items) in enumerate(by_kw.items()):
            row_bg = BG2 if i % 2 == 0 else BG
            rf = tk.Frame(grid, bg=row_bg); rf.pack(fill="x", pady=1)
            tk.Label(rf, text=f'"{kw}"', bg=row_bg, fg=ACCENT,
                     font=FONT_CODE_B, width=24, anchor="w").pack(side="left", padx=(10, 4), pady=6)
            tk.Label(rf, text=f"({len(items)})", bg=row_bg, fg=FG_DIM,
                     font=FONT_UI).pack(side="left", padx=(0, 12))
            default = next((b for b in bus_names if kw.upper() in b.upper()), "")
            var = tk.StringVar(value=default); kw_bus_vars[kw] = var
            ttk.Combobox(rf, textvariable=var, values=bus_names,
                         width=30, font=FONT_CODE).pack(side="left", pady=6)

        def do_fix():
            fixed, errors = 0, []
            for kw, var in kw_bus_vars.items():
                bus_id = self.buses.get(var.get())
                if not bus_id: errors.append(self._t("bus_missing").format(var.get())); continue
                for v in by_kw[kw]:
                    try: self._set_output_bus(v["id"], bus_id); fixed += 1
                    except Exception as e: errors.append(f'{v["name"]}: {e}')
            dlg.destroy()
            if errors:
                messagebox.showwarning(self._t("partial_title"),
                    self._t("reroute_partial").format(fixed, len(errors), "\n".join(errors)),
                    parent=self.root)
            else:
                messagebox.showinfo(self._t("complete_title"),
                    self._t("reroute_done").format(fixed), parent=self.root)
            self._run_scan("name" if is_name else "workunit")

        # Footer
        tk.Frame(dlg, bg=BORDER, height=1).pack(fill="x")
        bbar = tk.Frame(dlg, bg=BG3); bbar.pack(fill="x", pady=0)
        _ab(bbar, self._t("cancel"), dlg.destroy,
            preset="ghost", font=FONT_UI, padx=16).pack(side="right", padx=(4, 12), pady=10)
        _ab(bbar, self._t("apply_reroute"), do_fix,
            preset="danger", font=FONT_UIB, padx=16).pack(side="right", padx=4, pady=10)

    # ── Export ───────────────────────────────────────────────────────────────
    def _export_csv(self, is_name):
        results = self._get_filtered(is_name)
        if not results:
            messagebox.showinfo(self._t("info_title"), self._t("no_export"), parent=self.root); return
        path = filedialog.asksaveasfilename(defaultextension=".csv",
            filetypes=[("CSV","*.csv"),("*","*.*")], title=self._t("csv_title"), parent=self.root)
        if not path: return
        fields = ["name","work_unit","current_bus","trigger","expected_bus_keyword","path"]
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore")
            w.writeheader(); w.writerows(results)
        messagebox.showinfo(self._t("complete_title"), self._t("csv_done").format(path), parent=self.root)

    # ── Tooltip ──────────────────────────────────────────────────────────────
    @staticmethod
    def _attach_tooltip(widget, text_fn):
        tip = [None]
        def _show(e):
            if tip[0]: return
            text = text_fn() if callable(text_fn) else text_fn
            x = widget.winfo_rootx() + 20
            y = widget.winfo_rooty() + widget.winfo_height() + 6
            w = tk.Toplevel(widget); w.wm_overrideredirect(True)
            w.wm_geometry(f"+{x}+{y}")
            outer = tk.Frame(w, bg=ACCENT, padx=1, pady=1); outer.pack()
            tk.Label(outer, text=text, bg=BG3, fg=FG,
                     font=FONT_CODE, justify="left",
                     padx=12, pady=8).pack()
            tip[0] = w
        def _hide(e):
            if tip[0]: tip[0].destroy(); tip[0] = None
        widget.bind("<Enter>", _show, add="+")
        widget.bind("<Leave>", _hide, add="+")

    # ── Status ───────────────────────────────────────────────────────────────
    def _set_status(self, msg, color=FG_DIM, pulsing=False):
        if self._status_anim_id:
            try: self.root.after_cancel(self._status_anim_id)
            except Exception: pass
            self._status_anim_id = None

        def _update():
            self._status_lbl.config(text=msg, fg=color)
            self._status_dot.config(fg=color)

        self.root.after(0, _update)

        if pulsing:
            pulse_colors = [color, FG_MUT]
            def _pulse(idx=0):
                try:
                    self._status_dot.config(fg=pulse_colors[idx % 2])
                    self._status_anim_id = self.root.after(600, lambda: _pulse(idx + 1))
                except Exception: pass
            self.root.after(0, _pulse)


def main():
    import ctypes
    _TITLE = "Bus Routing Auditor  —  Wwise"
    hwnd = ctypes.windll.user32.FindWindowW(None, _TITLE)
    if hwnd:
        ctypes.windll.user32.ShowWindow(hwnd, 9)
        ctypes.windll.user32.SetForegroundWindow(hwnd)
        return
    root = tk.Tk()
    BusRoutingAuditor(root)
    root.mainloop()


if __name__ == "__main__":
    main()
