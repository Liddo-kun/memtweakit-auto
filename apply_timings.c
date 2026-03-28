/*
 * apply_timings.c - MemTweakIt Automation
 *
 * Launches MemTweakIt, sets memory timing values via Win32 SendMessage,
 * then clicks OK/Apply. Reads desired values from timings.ini.
 *
 * Build:
 *   cl /O2 /W3 /MT /D_UNICODE /DUNICODE apply_timings.c /Fe:apply_timings.exe user32.lib kernel32.lib shell32.lib
 *
 * Usage:
 *   apply_timings.exe                  Apply timings from timings.ini
 *   apply_timings.exe --discover       Discover controls, generate template INI
 *   apply_timings.exe --ini custom.ini Use custom INI file
 */

#define _WIN32_WINNT 0x0601
#define _CRT_SECURE_NO_WARNINGS

#ifndef UNICODE
#define UNICODE
#endif
#ifndef _UNICODE
#define _UNICODE
#endif

#include <windows.h>
#include <commctrl.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <wchar.h>
#include <limits.h>
#include <stdarg.h>
#include <conio.h>

/* shell32.dll export, linked via shell32.lib */
BOOL WINAPI IsUserAnAdmin(void);

/* ---- Constants ---- */

#define MAX_TIMINGS  256
#define MAX_NAME     256
#define MAX_VALUE    64
#define MAX_CONTROLS 512

/* ---- Data structures ---- */

typedef struct {
    wchar_t name[MAX_NAME];
    wchar_t value[MAX_VALUE];
    int applied;  /* 0=pending, 1=set, -1=failed */
} Timing;

typedef struct {
    wchar_t exe_path[MAX_PATH];
    wchar_t action[32];
    int start_delay;
    Timing timings[MAX_TIMINGS];
    int timing_count;
} Config;

typedef struct {
    HWND hwnd;
    RECT rect;
    wchar_t text[MAX_NAME];
} Control;

/* ---- Utility ---- */

static wchar_t *trim(wchar_t *s)
{
    while (*s == L' ' || *s == L'\t') s++;
    size_t len = wcslen(s);
    while (len > 0 && (s[len-1] == L' '  || s[len-1] == L'\t' ||
                       s[len-1] == L'\r' || s[len-1] == L'\n'))
        s[--len] = L'\0';
    return s;
}

/* Get directory of this exe, with trailing backslash */
static void get_exe_dir(wchar_t *buf)
{
    GetModuleFileNameW(NULL, buf, MAX_PATH);
    wchar_t *sep = wcsrchr(buf, L'\\');
    if (sep) *(sep + 1) = L'\0';
}

/* ---- Logging ---- */

static void LOG(const wchar_t *fmt, ...)
{
    va_list ap;
    va_start(ap, fmt);
    vwprintf(fmt, ap);
    va_end(ap);
}

/* ---- INI parser ---- */

static int load_config(const wchar_t *path, Config *cfg)
{
    FILE *f = _wfopen(path, L"r, ccs=UTF-8");
    if (!f) {
        LOG(L"ERROR: Cannot open %s\n", path);
        return 0;
    }

    memset(cfg, 0, sizeof(*cfg));
    wcscpy(cfg->action, L"ok");
    cfg->start_delay = 3;

    wchar_t line[1024];
    int section = 0;  /* 0=none, 1=Settings, 2=Timings */

    while (fgetws(line, 1024, f)) {
        wchar_t *s = trim(line);
        if (!*s || *s == L';' || *s == L'#') continue;

        if (*s == L'[') {
            wchar_t *end = wcschr(s, L']');
            if (end) {
                *end = L'\0';
                if (_wcsicmp(s + 1, L"Settings") == 0)      section = 1;
                else if (_wcsicmp(s + 1, L"Timings") == 0)   section = 2;
                else section = 0;
            }
            continue;
        }

        wchar_t *eq = wcschr(s, L'=');
        if (!eq) continue;
        *eq = L'\0';

        wchar_t *key = trim(s);
        wchar_t *val = trim(eq + 1);

        /* Strip inline comment (;) */
        wchar_t *semi = wcschr(val, L';');
        if (semi) { *semi = L'\0'; val = trim(val); }

        if (section == 1) {
            if (_wcsicmp(key, L"Path") == 0)
                wcsncpy(cfg->exe_path, val, MAX_PATH - 1);
            else if (_wcsicmp(key, L"StartDelay") == 0)
                cfg->start_delay = _wtoi(val);
            else if (_wcsicmp(key, L"Action") == 0)
                wcsncpy(cfg->action, val, 31);
        } else if (section == 2 && cfg->timing_count < MAX_TIMINGS) {
            Timing *t = &cfg->timings[cfg->timing_count++];
            wcsncpy(t->name, key, MAX_NAME - 1);
            wcsncpy(t->value, val, MAX_VALUE - 1);
            t->applied = 0;
        }
    }

    fclose(f);

    /* Default exe path: MemTweakIt.exe next to this exe */
    if (cfg->exe_path[0] == L'\0') {
        get_exe_dir(cfg->exe_path);
        wcscat(cfg->exe_path, L"MemTweakIt.exe");
    }

    return 1;
}

/* ---- Process launcher ---- */

static HANDLE launch_memtweakit(Config *cfg, DWORD *out_pid)
{
    STARTUPINFOW si = {0};
    PROCESS_INFORMATION pi = {0};
    si.cb = sizeof(si);

    wchar_t workdir[MAX_PATH];
    wcsncpy(workdir, cfg->exe_path, MAX_PATH - 1);
    workdir[MAX_PATH - 1] = L'\0';
    wchar_t *sep = wcsrchr(workdir, L'\\');
    if (!sep) sep = wcsrchr(workdir, L'/');
    if (sep) *sep = L'\0';

    if (!CreateProcessW(cfg->exe_path, NULL, NULL, NULL, FALSE,
                        0, NULL, workdir, &si, &pi)) {
        LOG(L"ERROR: CreateProcessW failed (%lu)\n", GetLastError());
        return NULL;
    }

    CloseHandle(pi.hThread);
    *out_pid = pi.dwProcessId;
    return pi.hProcess;  /* PROCESS_ALL_ACCESS - reuse for VirtualAllocEx */
}

/* ---- Window finder ---- */

static HWND find_tab_control(HWND parent)
{
    return FindWindowExW(parent, NULL, L"SysTabControl32", NULL);
}

typedef struct {
    DWORD target_pid;
    HWND  found_hwnd;
} FindWndData;

static BOOL CALLBACK find_wnd_cb(HWND hwnd, LPARAM lp)
{
    FindWndData *d = (FindWndData *)lp;
    DWORD pid;
    GetWindowThreadProcessId(hwnd, &pid);
    if (pid != d->target_pid || !IsWindowVisible(hwnd))
        return TRUE;
    if (find_tab_control(hwnd)) {
        d->found_hwnd = hwnd;
        return FALSE;
    }
    return TRUE;
}

static HWND find_window_by_pid(DWORD pid, HANDLE hproc, int timeout_sec)
{
    FindWndData d = {0};
    d.target_pid = pid;

    for (int ms = 0; ms < timeout_sec * 1000; ms += 500) {
        if (WaitForSingleObject(hproc, 0) == WAIT_OBJECT_0) {
            LOG(L"ERROR: MemTweakIt exited prematurely.\n");
            return NULL;
        }
        d.found_hwnd = NULL;
        EnumWindows(find_wnd_cb, (LPARAM)&d);
        if (d.found_hwnd) return d.found_hwnd;
        Sleep(500);
    }
    return NULL;
}

/* ---- Tab switcher ---- */

static BOOL switch_tab(HWND tab_ctrl, HANDLE hproc, int index)
{
    if ((int)SendMessageW(tab_ctrl, TCM_GETCURSEL, 0, 0) == index)
        return TRUE;

    int cx = 45 + index * 90, cy = 12;  /* fallback estimate */

    if (hproc) {
        LPVOID remote = VirtualAllocEx(hproc, NULL, sizeof(RECT),
                                       MEM_COMMIT, PAGE_READWRITE);
        if (remote) {
            SendMessageW(tab_ctrl, TCM_GETITEMRECT, (WPARAM)index, (LPARAM)remote);
            RECT r;
            SIZE_T n;
            if (ReadProcessMemory(hproc, remote, &r, sizeof(r), &n) &&
                n == sizeof(r)) {
                cx = (r.left + r.right) / 2;
                cy = (r.top + r.bottom) / 2;
            }
            VirtualFreeEx(hproc, remote, 0, MEM_RELEASE);
        }
    }

    LPARAM lp = MAKELPARAM(cx, cy);
    SendMessageW(tab_ctrl, WM_LBUTTONDOWN, MK_LBUTTON, lp);
    SendMessageW(tab_ctrl, WM_LBUTTONUP, 0, lp);

    /* SendMessage is synchronous - tab change should be immediate.
       Poll briefly in case target needs extra processing. */
    for (int w = 0; w < 50; w++) {
        if ((int)SendMessageW(tab_ctrl, TCM_GETCURSEL, 0, 0) == index)
            return TRUE;
        Sleep(10);
    }
    return FALSE;
}

/* ---- Control enumerator ---- */

typedef struct {
    Control *statics;  int *ns;  int max_s;
    Control *combos;   int *nc;  int max_c;
} EnumCtlData;

static BOOL CALLBACK enum_ctl_cb(HWND hwnd, LPARAM lp)
{
    EnumCtlData *d = (EnumCtlData *)lp;
    if (!IsWindowVisible(hwnd)) return TRUE;

    wchar_t cls[64];
    GetClassNameW(hwnd, cls, 64);

    if (wcscmp(cls, L"ComboBox") == 0 && *d->nc < d->max_c) {
        Control *c = &d->combos[(*d->nc)++];
        c->hwnd = hwnd;
        GetWindowRect(hwnd, &c->rect);
        c->text[0] = L'\0';
    }
    else if (wcscmp(cls, L"Static") == 0 && *d->ns < d->max_s) {
        wchar_t buf[MAX_NAME];
        GetWindowTextW(hwnd, buf, MAX_NAME);
        wchar_t *t = trim(buf);
        if (*t) {
            Control *c = &d->statics[(*d->ns)++];
            c->hwnd = hwnd;
            GetWindowRect(hwnd, &c->rect);
            wcsncpy(c->text, t, MAX_NAME - 1);
            c->text[MAX_NAME - 1] = L'\0';
        }
    }

    return TRUE;
}

static void get_controls(HWND parent,
                         Control *statics, int *ns,
                         Control *combos,  int *nc)
{
    *ns = *nc = 0;
    EnumCtlData d = { statics, ns, MAX_CONTROLS, combos, nc, MAX_CONTROLS };
    EnumChildWindows(parent, enum_ctl_cb, (LPARAM)&d);
}

/* ---- Spatial matching ---- */

static wchar_t *find_label_for_combo(RECT *cr, Control *statics, int ns)
{
    int combo_cy   = (cr->top + cr->bottom) / 2;
    int combo_left = cr->left;
    int best = INT_MAX;
    wchar_t *label = NULL;

    for (int i = 0; i < ns; i++) {
        int scy = (statics[i].rect.top + statics[i].rect.bottom) / 2;
        int dx  = combo_left - statics[i].rect.right;
        int dy  = abs(scy - combo_cy);
        if (dy < 25 && dx > -5 && dx < 400) {
            int score = abs(dx) + dy * 10;
            if (score < best) { best = score; label = statics[i].text; }
        }
    }
    return label;
}

/* Forward declaration */
static void get_combo_text(HWND hwnd, wchar_t *out, int max);

/* ---- Value setter ---- */

static LRESULT find_combo_item(HWND combo, const wchar_t *value)
{
    /* Try exact string match (works for standard combos) */
    LRESULT raw = SendMessageW(combo, CB_FINDSTRINGEXACT,
                               (WPARAM)-1, (LPARAM)value);
    /* WoW64 fix: 32-bit CB_ERR (0xFFFFFFFF) can be zero-extended to
       0x00000000FFFFFFFF instead of sign-extended to 0xFFFFFFFFFFFFFFFF.
       Cast to int to get correct sign. */
    int target32 = (int)raw;
    if (target32 >= 0) return (LRESULT)target32;

    /* CB_FINDSTRINGEXACT fails cross-process on owner-drawn combos.
       For numeric values, verify text at that index matches. */
    LRESULT count = SendMessageW(combo, CB_GETCOUNT, 0, 0);
    wchar_t *end;
    long idx = wcstol(value, &end, 10);
    if (*end == L'\0' && idx >= 0 && idx < count) {
        wchar_t buf[MAX_VALUE] = {0};
        LRESULT n = SendMessageW(combo, CB_GETLBTEXTLEN, (WPARAM)idx, 0);
        if (n > 0 && n < MAX_VALUE) {
            SendMessageW(combo, CB_GETLBTEXT, (WPARAM)idx, (LPARAM)buf);
            if (wcscmp(trim(buf), value) == 0) return (LRESULT)idx;
        }
        /* Items are sequential integers; trust index even if text read fails */
        return (LRESULT)idx;
    }
    return -1;
}

/* Read current combo value as trimmed text (same logic as discover mode) */
static void get_combo_value(HWND hwnd, wchar_t *out, int max)
{
    get_combo_text(hwnd, out, max);
    wchar_t *t = trim(out);
    if (t != out) {
        size_t len = wcslen(t);
        wmemmove(out, t, len + 1);
    }
}

/* Return: 0=failed, 1=set, 2=already correct */
static int set_combo_value(HWND combo, const wchar_t *value)
{
    /* Text-based comparison - same method as discover mode */
    wchar_t cur[MAX_VALUE];
    get_combo_value(combo, cur, MAX_VALUE);
    if (wcscmp(cur, value) == 0) return 2; /* already correct */

    /* Find target index */
    LRESULT target = find_combo_item(combo, value);
    if (target < 0) {
        /* Typable combo (CBS_DROPDOWN) - no dropdown items.
           Select all text and type the new value. */
        SendMessageW(combo, CB_SETEDITSEL, 0, MAKELPARAM(0, -1));
        for (int i = 0; value[i]; i++) {
            PostMessageW(combo, WM_CHAR, (WPARAM)value[i], 0);
            Sleep(5);
        }
        /* Wait for posted chars to be processed */
        for (int retry = 0; retry < 50; retry++) {
            get_combo_value(combo, cur, MAX_VALUE);
            if (wcscmp(cur, value) == 0) return 1;
            Sleep(10);
        }
        return 0;
    }

    LRESULT current = SendMessageW(combo, CB_GETCURSEL, 0, 0);
    if (current < 0) current = 0;
    if (current == target) return 2; /* index matches despite text mismatch */

    int diff = (int)(target - current);
    int steps = abs(diff);
    if (steps > 5000) return 0; /* sanity limit */
    WPARAM vk = (diff > 0) ? VK_DOWN : VK_UP;
    LPARAM kdown = (vk == VK_DOWN) ? 0x00500001L : 0x00480001L;
    LPARAM kup   = (vk == VK_DOWN) ? 0xC0500001L : 0xC0480001L;

    /* Navigate from current to target using PostMessage. */
    for (int i = 0; i < steps; i++) {
        PostMessageW(combo, WM_KEYDOWN, vk, kdown);
        PostMessageW(combo, WM_KEYUP, vk, kup);
        if (i % 50 == 49) Sleep(5);
    }

    /* Wait for posted key messages to be processed */
    for (int retry = 0; retry < 50; retry++) {
        get_combo_value(combo, cur, MAX_VALUE);
        if (wcscmp(cur, value) == 0) return 1;
        Sleep(10);
    }
    return 0;
}

/* ---- Button clicker ---- */

static BOOL click_button(HWND parent, const wchar_t *text)
{
    for (HWND h = NULL; (h = FindWindowExW(parent, h, L"Button", NULL));) {
        if (!IsWindowVisible(h) || !IsWindowEnabled(h)) continue;
        wchar_t raw[256], clean[256];
        GetWindowTextW(h, raw, 256);
        /* Strip '&' accelerator chars */
        int j = 0;
        for (int i = 0; raw[i] && j < 255; i++)
            if (raw[i] != L'&') clean[j++] = raw[i];
        clean[j] = L'\0';
        if (_wcsicmp(trim(clean), text) == 0) {
            SendMessageW(h, BM_CLICK, 0, 0);
            return TRUE;
        }
    }
    return FALSE;
}

/* ---- Combo value reader (for discover) ---- */

static void get_combo_text(HWND hwnd, wchar_t *out, int max)
{
    LRESULT idx = SendMessageW(hwnd, CB_GETCURSEL, 0, 0);
    if (idx < 0) { GetWindowTextW(hwnd, out, max); return; }
    LRESULT n = SendMessageW(hwnd, CB_GETLBTEXTLEN, (WPARAM)idx, 0);
    if (n <= 0 || n >= max) { out[0] = L'\0'; return; }
    SendMessageW(hwnd, CB_GETLBTEXT, (WPARAM)idx, (LPARAM)out);
}

/* ---- Discover mode ---- */

static void discover(HWND main_hwnd, HWND tab_ctrl, HANDLE hproc,
                     const wchar_t *out_path)
{
    int n_tabs = (int)SendMessageW(tab_ctrl, TCM_GETITEMCOUNT, 0, 0);
    LOG(L"Window: HWND=0x%p  Tabs=%d\n\n", (void *)main_hwnd, n_tabs);

    FILE *fout = _wfopen(out_path, L"w, ccs=UTF-8");
    if (fout) {
        fwprintf(fout, L"; MemTweakIt Automation - Timing Configuration\n");
        fwprintf(fout, L"; Generated by: apply_timings.exe --discover\n");
        fwprintf(fout, L";\n");
        fwprintf(fout, L"; Edit values below, then run: apply_timings.exe\n");
        fwprintf(fout, L"; Comment out (;) or delete any timing you don't want to change.\n\n");
        fwprintf(fout, L"[Settings]\n");
        fwprintf(fout, L"; Path to MemTweakIt.exe (defaults to same directory as apply_timings.exe)\n");
        fwprintf(fout, L"; Path = C:\\path\\to\\MemTweakIt.exe\n");
        fwprintf(fout, L"StartDelay = 3\n");
        fwprintf(fout, L"; Action: apply = click Apply only, ok = click OK (apply+close)\n");
        fwprintf(fout, L"Action = ok\n\n");
        fwprintf(fout, L"[Timings]\n");
    }

    int total = 0;
    Control *statics = (Control *)malloc(MAX_CONTROLS * sizeof(Control));
    Control *combos  = (Control *)malloc(MAX_CONTROLS * sizeof(Control));

    for (int i = 0; i < n_tabs; i++) {
        if (!switch_tab(tab_ctrl, hproc, i)) {
            LOG(L"  [!] Could not switch to tab %d\n", i);
            continue;
        }

        int ns = 0, nc = 0;
        get_controls(main_hwnd, statics, &ns, combos, &nc);
        if (nc == 0) continue;

        LOG(L"--- Timings #%d (%d controls) ---\n", i + 1, nc);
        if (fout) fwprintf(fout, L"; -- Timings #%d --\n", i + 1);

        for (int j = 0; j < nc; j++) {
            wchar_t *label = find_label_for_combo(&combos[j].rect, statics, ns);
            if (!label) continue;

            wchar_t val[MAX_VALUE];
            get_combo_text(combos[j].hwnd, val, MAX_VALUE);
            BOOL enabled = IsWindowEnabled(combos[j].hwnd);

            LOG(L"  %-45s = %s%s\n", label, val,
                    enabled ? L"" : L"  [read-only]");

            if (enabled && fout) {
                fwprintf(fout, L"%s = %s\n", label, val);
                total++;
            }
        }
        if (fout) fwprintf(fout, L"\n");
        LOG(L"\n");
    }

    switch_tab(tab_ctrl, hproc, 0);
    free(statics);
    free(combos);
    if (fout) fclose(fout);

    LOG(L"Total editable controls: %d\n", total);
    LOG(L"Template written to: %s\n", out_path);
}

/* ---- Main ---- */

int wmain(int argc, wchar_t *argv[])
{
    int discover_mode = 0;
    wchar_t *ini_arg = NULL;
    int user_specified_ini = 0;

    for (int i = 1; i < argc; i++) {
        if (wcscmp(argv[i], L"--discover") == 0)
            discover_mode = 1;
        else if (wcscmp(argv[i], L"--ini") == 0 && i + 1 < argc) {
            ini_arg = argv[++i];
            user_specified_ini = 1;
        }
    }

    HWND con = GetConsoleWindow();

    if (!IsUserAnAdmin()) {
        if (con) ShowWindow(con, SW_RESTORE);
        LOG(L"ERROR: Must run as administrator (MemTweakIt needs admin).\n");
        return 1;
    }

    /* Default INI: same directory as this exe */
    wchar_t exe_dir[MAX_PATH];
    get_exe_dir(exe_dir);

    wchar_t default_ini[MAX_PATH];
    if (!ini_arg) {
        wcscpy(default_ini, exe_dir);
        wcscat(default_ini, L"timings.ini");
        ini_arg = default_ini;
    }

    int ini_exists = (GetFileAttributesW(ini_arg) != INVALID_FILE_ATTRIBUTES);
    int first_run = (!ini_exists && !discover_mode && !user_specified_ini);

    Config cfg;
    if (ini_exists) {
        if (!load_config(ini_arg, &cfg)) return 1;
    } else {
        if (!first_run && !discover_mode) {
            /* User specified --ini with a nonexistent file */
            LOG(L"ERROR: Cannot open %s\n", ini_arg);
            return 1;
        }
        /* No INI - set up defaults for discover / first run */
        memset(&cfg, 0, sizeof(cfg));
        wcscpy(cfg.exe_path, exe_dir);
        wcscat(cfg.exe_path, L"MemTweakIt.exe");
        wcscpy(cfg.action, L"ok");
        cfg.start_delay = 3;
    }

    /* Minimize console for normal background operation */
    if (!first_run && con) ShowWindow(con, SW_MINIMIZE);

    if (GetFileAttributesW(cfg.exe_path) == INVALID_FILE_ATTRIBUTES) {
        if (con) ShowWindow(con, SW_RESTORE);
        LOG(L"ERROR: MemTweakIt not found at: %s\n", cfg.exe_path);
        if (first_run) {
            LOG(L"\nPlace MemTweakIt.exe in the same directory as apply_timings.exe.\n");
            LOG(L"Press any key to exit...\n");
            _getwch();
        }
        return 1;
    }

    LOG(L"Launching MemTweakIt...\n");
    DWORD pid;
    HANDLE hproc = launch_memtweakit(&cfg, &pid);
    if (!hproc) return 1;

    int timeout = cfg.start_delay > 15 ? cfg.start_delay : 15;
    LOG(L"Waiting up to %ds for window (PID %lu)...\n", timeout, pid);

    HWND hwnd = find_window_by_pid(pid, hproc, timeout);
    if (!hwnd) {
        LOG(L"ERROR: Timed out waiting for MemTweakIt window.\n");
        CloseHandle(hproc);
        return 1;
    }

    ShowWindow(hwnd, SW_MINIMIZE);

    HWND tab_ctrl = find_tab_control(hwnd);
    if (!tab_ctrl) {
        LOG(L"ERROR: Tab control not found.\n");
        CloseHandle(hproc);
        return 1;
    }

    int n_tabs = (int)SendMessageW(tab_ctrl, TCM_GETITEMCOUNT, 0, 0);
    LOG(L"MemTweakIt found (HWND=0x%p, PID=%lu, Tabs=%d)\n",
            (void *)hwnd, pid, n_tabs);

    if (first_run) {
        /* First run: discover current timings and write timings.ini */
        LOG(L"\nNo timings.ini found. Discovering current timings...\n\n");
        discover(hwnd, tab_ctrl, hproc, ini_arg);
        click_button(hwnd, L"OK");
        CloseHandle(hproc);
        LOG(L"\nCreated %s with your current timing values.\n", ini_arg);
        LOG(L"Edit the file to set your desired values, then run apply_timings.exe again.\n\n");
        LOG(L"Press any key to exit...\n");
        _getwch();
        return 0;
    }

    if (discover_mode) {
        wchar_t disc_path[MAX_PATH];
        wcscpy(disc_path, exe_dir);
        wcscat(disc_path, L"timings_discovered.ini");
        discover(hwnd, tab_ctrl, hproc, disc_path);
        CloseHandle(hproc);
        return 0;
    }

    /* ---- Apply timings ---- */

    if (cfg.timing_count == 0) {
        LOG(L"No timings to apply.\n");
        CloseHandle(hproc);
        return 0;
    }

    LOG(L"Loaded %d timing(s) from %s\n", cfg.timing_count, ini_arg);

    Control *statics = (Control *)malloc(MAX_CONTROLS * sizeof(Control));
    Control *combos  = (Control *)malloc(MAX_CONTROLS * sizeof(Control));
    int applied = 0, skipped = 0, failed = 0;

    /* No SetForegroundWindow/AttachThreadInput - they cause MemTweakIt
       to reset some combo Win32 states via WM_ACTIVATE. PostMessage
       doesn't need foreground or focus. */

    for (int ti = 0; ti < n_tabs; ti++) {
        /* Check if all timings resolved */
        int pending = 0;
        for (int k = 0; k < cfg.timing_count; k++)
            if (cfg.timings[k].applied == 0) pending++;
        if (!pending) break;

        if (!switch_tab(tab_ctrl, hproc, ti)) continue;

        int ns = 0, nc = 0;
        get_controls(hwnd, statics, &ns, combos, &nc);
        if (nc == 0) continue;

        for (int j = 0; j < nc; j++) {
            wchar_t *label = find_label_for_combo(&combos[j].rect, statics, ns);
            if (!label) continue;

            for (int k = 0; k < cfg.timing_count; k++) {
                if (cfg.timings[k].applied != 0) continue;
                if (wcscmp(cfg.timings[k].name, label) != 0) continue;

                if (!IsWindowEnabled(combos[j].hwnd)) {
                    LOG(L"  SKIP %s (read-only)\n", label);
                    cfg.timings[k].applied = -1;
                    failed++;
                    break;
                }

                {
                    int result = set_combo_value(combos[j].hwnd,
                                                cfg.timings[k].value);
                    if (result == 2) {
                        LOG(L"  OK   %s = %s\n", label, cfg.timings[k].value);
                        cfg.timings[k].applied = 1;
                        skipped++;
                    } else if (result == 1) {
                        LOG(L"  SET  %s = %s\n", label, cfg.timings[k].value);
                        cfg.timings[k].applied = 1;
                        applied++;
                    } else {
                        wchar_t cur[MAX_VALUE];
                        get_combo_value(combos[j].hwnd, cur, MAX_VALUE);
                        LOG(L"  FAIL %s = %s (cur=\"%s\")\n",
                                label, cfg.timings[k].value, cur);
                        cfg.timings[k].applied = -1;
                        failed++;
                    }
                }
                break;
            }
        }
    }


    int not_found = 0;
    for (int k = 0; k < cfg.timing_count; k++) {
        if (cfg.timings[k].applied == 0) {
            LOG(L"  MISS %s (control not found in any tab)\n",
                    cfg.timings[k].name);
            not_found++;
        }
    }

    LOG(L"\nResult: %d applied, %d already OK, %d failed, %d not found\n",
            applied, skipped, failed, not_found);

    /* Verification pass: read back all values to check if they persisted */
    LOG(L"\nVerifying values persisted...\n");
    int reset_count = 0;
    for (int ti = 0; ti < n_tabs; ti++) {
        if (!switch_tab(tab_ctrl, hproc, ti)) continue;
        int ns2 = 0, nc2 = 0;
        get_controls(hwnd, statics, &ns2, combos, &nc2);
        for (int j = 0; j < nc2; j++) {
            wchar_t *label = find_label_for_combo(&combos[j].rect, statics, ns2);
            if (!label) continue;
            for (int k = 0; k < cfg.timing_count; k++) {
                if (cfg.timings[k].applied != 1) continue;
                if (wcscmp(cfg.timings[k].name, label) != 0) continue;
                wchar_t actual[MAX_VALUE];
                get_combo_text(combos[j].hwnd, actual, MAX_VALUE);
                if (wcscmp(trim(actual), cfg.timings[k].value) != 0) {
                    LOG(L"  RESET %s: expected=%s actual=%s\n",
                        label, cfg.timings[k].value, trim(actual));
                    reset_count++;
                }
                break;
            }
        }
    }
    if (reset_count == 0)
        LOG(L"  All values verified OK\n");
    else
        LOG(L"  %d values were reset by tab switching!\n", reset_count);

    free(statics);
    free(combos);

    /* Click Apply first (commits to hardware), then OK (closes window) */
    if (click_button(hwnd, L"Apply"))
        LOG(L"Clicked Apply\n");
    else
        LOG(L"WARNING: Apply button not found\n");

    Sleep(100);

    if (_wcsicmp(cfg.action, L"ok") == 0) {
        if (click_button(hwnd, L"OK"))
            LOG(L"Clicked OK\n");
        else
            LOG(L"WARNING: OK button not found\n");
    }

    CloseHandle(hproc);
    return (failed == 0 && not_found == 0) ? 0 : 1;
}
