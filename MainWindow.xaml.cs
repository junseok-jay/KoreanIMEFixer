using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Windows;
using System.Windows.Threading;
using KoreanIMEFixer.Input;
using KoreanIMEFixer.IME;
using KoreanIMEFixer.Logging;
using Keys = System.Windows.Forms.Keys;

namespace KoreanIMEFixer
{
    public partial class MainWindow : Window
    {
        private readonly GlobalKeyboardHook _hook;
        private readonly NoopAnalyzer _analyzer;
        private readonly IMEStateRestorer _restorer;
        private DateTime _lastRestore = DateTime.MinValue;
        private readonly DispatcherTimer _focusTimer;
        private bool _aggressiveMode = false;
        private bool _autoEnabled = false;
        private bool _doubleEnterFixEnabled = true;
        private bool _forceNextKeyFallback = false;
        // buffer for fallback composition: hold first jamo and wait a short window for second
        private char? _bufferedFallbackJamo = null;
        private readonly DispatcherTimer _fallbackBufferTimer;
        private string[] _processWhitelist = new[] { "code", "notion" };
        private readonly string _configPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "KoreanIMEFixer", "config.json");

        // Lightweight no-op analyzer to keep Enter-only behavior while disabling other features.
        private class NoopAnalyzer
        {
            public event EventHandler? FirstLetterBugDetected;
            public void BeginTypingExpectingKorean() { }
            public void EndTypingExpectingKorean() { }
            public bool IsFirstLetterBugCandidate(System.Windows.Forms.Keys k) => false;
            public void OnKeyPressed(System.Windows.Forms.Keys k) { }
        }

        public MainWindow()
        {
            InitializeComponent();

            _hook = new GlobalKeyboardHook();
            _analyzer = new NoopAnalyzer();
            _restorer = new IMEStateRestorer();

            LogService.Write("Application started. Monitoring is available but not auto-enabled.");

            _focusTimer = new DispatcherTimer();
            _focusTimer.Interval = TimeSpan.FromMilliseconds(500);
            _focusTimer.Tick += FocusTimer_Tick;
            _focusTimer.Start();

            _fallbackBufferTimer = new DispatcherTimer();
            _fallbackBufferTimer.Interval = TimeSpan.FromMilliseconds(500);
            _fallbackBufferTimer.Tick += FallbackBufferTimer_Tick;

            LoadWhitelist();
            LoadConfig();

            StatusText.Text = "Korean IME Fixer — 실행 중 (대기)";
            _hook.KeyPressed += Hook_KeyPressed;

            _analyzer.FirstLetterBugDetected += (s, e) =>
            {
                if ((DateTime.UtcNow - _lastRestore) < TimeSpan.FromMilliseconds(500))
                {
                    LogService.Write("First-letter IME bug detected but ignored due to debounce.");
                    return;
                }

                LogService.Write("First-letter IME bug detected. (async fallback restore scheduled)");
                System.Threading.Tasks.Task.Run(async () =>
                {
                    await System.Threading.Tasks.Task.Delay(30);
                    try
                    {
                        _restorer.ForceToKorean();
                        _analyzer.EndTypingExpectingKorean();
                        _lastRestore = DateTime.UtcNow;
                        LogService.Write("Async fallback restore executed.");
                    }
                    catch (Exception ex)
                    {
                        LogService.Write("Async fallback restore failed: " + ex.Message);
                    }
                });
            };

            // show initial log
            UpdateLogView();
        }

        private void Hook_KeyPressed(object? sender, Input.KeyPressedEventArgs e)
        {
            try
            {
                // 물리적 키 기록
                try { Input.InputState.ReportPhysicalKey((int)e.Key); } catch { }
                
                // 버퍼링된 자모 조합 처리
                try
                {
                    if (_bufferedFallbackJamo.HasValue && Input.InputAnalyzer.TryGetDubeolsikJamo(e.Key, out char second))
                    {
                        char first = _bufferedFallbackJamo.Value;
                        _bufferedFallbackJamo = null;
                        _fallbackBufferTimer.Stop();
                        bool ok = _restorer.ComposeAndPastePair(first, second);
                        LogService.Write($"Buffered compose '{first}'+'{second}' success={ok}");
                        if (ok)
                        {
                            e.Suppress = true;
                            _analyzer.EndTypingExpectingKorean();
                            _lastRestore = DateTime.UtcNow;
                            _aggressiveMode = false;
                            UpdateLogView();
                            return;
                        }
                        else
                        {
                            bool sentFirst = _restorer.SendUnicodeChar(first);
                            LogService.Write($"Buffered first char '{first}' sent={sentFirst}");
                            if (sentFirst)
                            {
                                _analyzer.EndTypingExpectingKorean();
                                _lastRestore = DateTime.UtcNow;
                                _aggressiveMode = false;
                                UpdateLogView();
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogService.Write("Buffered compose failed: " + ex.Message);
                }

                // ========== ENTER 키 처리 ==========
                if (e.Key == Keys.Enter || e.Key == Keys.Return)
                {
                    _lastRestore = DateTime.UtcNow;
                    _aggressiveMode = true; // 다음 3개 키까지 공격적 체크
                    
                    LogService.Write("━━━━━ ENTER pressed ━━━━━");
                    
                    // 즉시 한글 모드 기대 상태로 설정
                    _analyzer.BeginTypingExpectingKorean();
                    
                    IntPtr hwnd = IntPtr.Zero;
                    try { hwnd = GetForegroundWindow(); } catch { }
                    _ = GetWindowThreadProcessId(hwnd, out uint threadId);
                    
                    // 포그라운드 프로세스 확인
                    string fgProc = string.Empty;
                    bool isNotion = false;
                    try 
                    { 
                        _ = GetWindowThreadProcessId(hwnd, out uint fgPid);
                        using Process p = Process.GetProcessById((int)fgPid); 
                        fgProc = p.ProcessName.ToLowerInvariant();
                        isNotion = fgProc.Contains("notion");
                    } 
                    catch { }
                    
                    LogService.Write($"Foreground process: {fgProc}, isNotion={isNotion}");
                    
                    // 노션인 경우 더 짧은 캐럿 대기 (빠른 타이핑 대응) — 더 빠르게 동작하도록 축소
                    int caretTimeout = isNotion ? 60 : 200;
                    
                    // 비동기로 캐럿 체크
                    System.Threading.Tasks.Task.Run(() =>
                    {
                        bool caret = WaitForCaret(threadId, caretTimeout);
                        LogService.Write($"WaitForCaret({caretTimeout}ms) result={caret}");
                        
                        if (!caret)
                        {
                            // 캐럿 없음 - 더블 엔터 또는 fallback 준비
                            try
                            {
                                // Log AutomationElement at caret/focus to detect if we're in a table cell (Notion)
                                try
                                {
                                    LogFocusedAutomationElementInfo();
                                }
                                catch { }

                                IntPtr fgw = GetForegroundWindow();
                                _ = GetWindowThreadProcessId(fgw, out uint fgPid);
                                string currentProc = string.Empty;
                                try { using Process p = Process.GetProcessById((int)fgPid); currentProc = p.ProcessName.ToLowerInvariant(); } catch { }

                                bool isWhitelisted = _processWhitelist.Any(w => !string.IsNullOrEmpty(currentProc) && currentProc.Contains(w, StringComparison.OrdinalIgnoreCase));
                                
                                if (isWhitelisted && _doubleEnterFixEnabled)
                                {
                                    // run a quick heuristic to check whether the focused area looks like a Notion table cell
                                    bool likelyTable = false;
                                    try
                                    {
                                        likelyTable = IsLikelyNotionTable();
                                        LogService.Write($"Notion table-detection heuristic: likelyTable={likelyTable}");
                                    }
                                    catch (Exception hx)
                                    {
                                        LogService.Write("Table-detection heuristic failed: " + hx.Message);
                                        likelyTable = false;
                                    }

                                    if (!likelyTable)
                                    {
                                        LogService.Write("No caret + whitelisted but not a detected table → using fallback mode instead of double-Enter");
                                        // Activate fallback mode rather than aggressive double-Enter to avoid false positives
                                        _forceNextKeyFallback = true;
                                        return;
                                    }

                                    LogService.Write("No caret + whitelisted → Double-Enter (fast) fix");
                                    _restorer.ForceToKorean();
                                    // very short pause to let layout switch take effect
                                    System.Threading.Thread.Sleep(8);

                                    // Prefer scancode SendInput (more likely accepted by Electron)
                                    bool sc = _restorer.SendEnterScancode();
                                    LogService.Write($"Double-Enter fast: scancode sent={sc}");
                                    if (!sc)
                                    {
                                        // quick virtual-key attempt
                                        bool vk = _restorer.SendVirtualKey((ushort)Keys.Enter);
                                        LogService.Write($"Double-Enter fast: virtual-key sent={vk}");
                                        if (!vk)
                                        {
                                            // last-resort: PostMessage to focused control
                                            try
                                            {
                                                IntPtr fg = GetForegroundWindow();
                                                uint pid; uint tid = GetWindowThreadProcessId(fg, out pid);
                                                IntPtr target = fg;
                                                try
                                                {
                                                    var gti = new GUITHREADINFO();
                                                    gti.cbSize = System.Runtime.InteropServices.Marshal.SizeOf(typeof(GUITHREADINFO));
                                                    if (tid != 0 && GetGUIThreadInfo(tid, ref gti))
                                                    {
                                                        if (gti.hwndFocus != IntPtr.Zero) target = gti.hwndFocus;
                                                    }
                                                }
                                                catch { }

                                                bool posted = _restorer.TryPostEnterToWindow(target);
                                                LogService.Write($"Double-Enter fast: PostMessage fallback success={posted} to HWND=0x{target.ToInt64():X}");
                                            }
                                            catch (Exception ex)
                                            {
                                                LogService.Write("Double-Enter fast: final PostMessage fallback failed: " + ex.Message);
                                            }
                                        }
                                    }

                                    System.Threading.Thread.Sleep(6);
                                    return;
                                }
                                
                                // fallback 모드 활성화
                                if (isWhitelisted)
                                {
                                    _forceNextKeyFallback = true;
                                    LogService.Write("Fallback mode activated for: " + currentProc);
                                }
                            }
                            catch (Exception ex)
                            {
                                LogService.Write("Post-Enter processing error: " + ex.Message);
                            }
                        }
                        else
                        {
                            LogService.Write("Caret detected - normal typing expected");
                        }
                    });
                    
                    return;
                }

                // ========== 공격적 모드: Enter 후 3초 이내 모든 영문 키 체크 ==========
                bool inAggressiveWindow = (DateTime.UtcNow - _lastRestore
                ).TotalMilliseconds < 3000;
                
                if (_aggressiveMode && inAggressiveWindow)
                {
                    bool isLetterKey = Input.InputAnalyzer.IsVirtualKeyLetter(e.Key);
                    
                    if (isLetterKey)
                    {
                        // IME 상태 실시간 체크
                        bool isKoreanMode = false;
                        try
                        {
                            var monitor = new IME.IMEStateMonitor();
                            isKoreanMode = monitor.IsKoreanMode();
                        }
                        catch { }
                        
                        LogService.Write($"AggressiveMode: key={(int)e.Key}, isKoreanMode={isKoreanMode}");
                        
                        // 영문 모드면 즉시 수정
                        if (!isKoreanMode)
                        {
                            // Debounce
                            if ((DateTime.UtcNow - _lastRestore) < TimeSpan.FromMilliseconds(300))
                            {
                                LogService.Write("Aggressive fix skipped (debounce)");
                                return;
                            }
                            
                            LogService.Write($"★★★ FIXING FIRST-LETTER BUG ★★★ key={(int)e.Key}");
                            
                            e.Suppress = true; // 원본 키 억제
                            
                            // 자모 매핑 시도
                            if (Input.InputAnalyzer.TryGetDubeolsikJamo(e.Key, out char jamo))
                            {
                                LogService.Write($"Mapped to jamo: '{jamo}'");
                                
                                // 방법 1: IME 강제 전환 + 키 재주입 (가장 안정적)
                                _restorer.ForceToKorean();
                                System.Threading.Thread.Sleep(60);
                                
                                bool sent = _restorer.SendVirtualKey((ushort)e.Key);
                                LogService.Write($"Method 1 (VK reinject): sent={sent}");
                                
                                if (!sent)
                                {
                                    // 방법 2: Unicode 직접 전송 (fallback)
                                    System.Threading.Thread.Sleep(30);
                                    bool sentUnicode = _restorer.SendUnicodeChar(jamo);
                                    LogService.Write($"Method 2 (Unicode): sent={sentUnicode}");
                                }
                            }
                            else
                            {
                                // 매핑 실패 - 그냥 IME 전환 후 키 재주입
                                LogService.Write("No jamo mapping - forcing Korean and reinject");
                                _restorer.ForceToKorean();
                                System.Threading.Thread.Sleep(60);
                                _restorer.SendVirtualKey((ushort)e.Key);
                            }
                            
                            _lastRestore = DateTime.UtcNow;
                            _aggressiveMode = false; // 한 번 수정하면 끔
                            UpdateLogView();
                            return;
                        }
                        else
                        {
                            // 한글 모드 확인됨 - 공격적 모드 해제
                            LogService.Write("Korean mode confirmed - aggressive mode OFF");
                            _aggressiveMode = false;
                        }
                    }
                }
                
                // 공격적 모드 타임아웃
                if (_aggressiveMode && !inAggressiveWindow)
                {
                    LogService.Write("Aggressive mode timeout");
                    _aggressiveMode = false;
                }

                // ========== Fallback 모드 처리 ==========
                if (_forceNextKeyFallback)
                {
                    // 화이트리스트 재확인
                    try
                    {
                        IntPtr fgw = GetForegroundWindow();
                        _ = GetWindowThreadProcessId(fgw, out uint fgPid);
                        string fgProcCheck = string.Empty;
                        try { using Process p = Process.GetProcessById((int)fgPid); fgProcCheck = p.ProcessName.ToLowerInvariant(); } catch { }
                        bool stillMatch = _processWhitelist.Any(w => !string.IsNullOrEmpty(fgProcCheck) && fgProcCheck.Contains(w, StringComparison.OrdinalIgnoreCase));
                        if (!stillMatch)
                        {
                            LogService.Write($"Fallback cancelled - not whitelisted: {fgProcCheck}");
                            _forceNextKeyFallback = false;
                        }
                    }
                    catch { }

                    if (_forceNextKeyFallback)
                    {
                        _forceNextKeyFallback = false;
                        LogService.Write($"Fallback mode handling key={(int)e.Key}");
                        
                        if (Input.InputAnalyzer.TryGetDubeolsikJamo(e.Key, out char jamo))
                        {
                            if (!_bufferedFallbackJamo.HasValue)
                            {
                                _bufferedFallbackJamo = jamo;
                                _fallbackBufferTimer.Stop();
                                _fallbackBufferTimer.Start();
                                LogService.Write($"Buffered jamo '{jamo}' for composition");
                                e.Suppress = true;
                                return;
                            }
                            else
                            {
                                char first = _bufferedFallbackJamo.Value;
                                _bufferedFallbackJamo = null;
                                _fallbackBufferTimer.Stop();
                                bool ok = _restorer.ComposeAndPastePair(first, jamo);
                                LogService.Write($"Compose '{first}'+'{jamo}' success={ok}");
                                if (ok)
                                {
                                    e.Suppress = true;
                                    _analyzer.EndTypingExpectingKorean();
                                    _lastRestore = DateTime.UtcNow;
                                    UpdateLogView();
                                    return;
                                }

                                bool sentFirst = _restorer.SendUnicodeChar(first);
                                if (sentFirst)
                                {
                                    _analyzer.EndTypingExpectingKorean();
                                    _lastRestore = DateTime.UtcNow;
                                    UpdateLogView();
                                }
                            }
                        }
                    }
                }

                // ========== 기존 분석기 로직 ==========
                if (_analyzer.IsFirstLetterBugCandidate(e.Key))
                {
                    if ((DateTime.UtcNow - _lastRestore) < TimeSpan.FromMilliseconds(300))
                    {
                        LogService.Write("Analyzer candidate ignored (debounce)");
                        return;
                    }

                    LogService.Write($"Analyzer detected candidate: key={(int)e.Key}");

                    IntPtr hwnd = IntPtr.Zero;
                    try { hwnd = GetForegroundWindow(); } catch { }
                    _ = GetWindowThreadProcessId(hwnd, out uint threadId);

                    bool caretAppeared = WaitForCaret(threadId, 200);
                    LogService.Write($"Analyzer WaitForCaret={caretAppeared}");

                    if (!caretAppeared)
                    {
                        LogService.Write("No caret - Unicode fallback");

                        if (Input.InputAnalyzer.TryGetDubeolsikJamo(e.Key, out char jamo))
                        {
                            bool sentJ = _restorer.SendUnicodeChar(jamo);
                            LogService.Write($"Unicode fallback '{jamo}' sent={sentJ}");
                            if (sentJ)
                            {
                                e.Suppress = true;
                                _analyzer.EndTypingExpectingKorean();
                                _lastRestore = DateTime.UtcNow;
                                UpdateLogView();
                                return;
                            }
                        }
                        return;
                    }

                    LogService.Write($"Suppressing key={(int)e.Key}, forcing Korean");
                    e.Suppress = true;

                    _restorer.ForceToKorean();
                    System.Threading.Thread.Sleep(60);
                    bool sent = _restorer.SendVirtualKey((ushort)e.Key);
                    LogService.Write($"Reinjected key={(int)e.Key}, sent={sent}");

                    _analyzer.EndTypingExpectingKorean();
                    _lastRestore = DateTime.UtcNow;
                    UpdateLogView();
                }
                else
                {
                    _analyzer.OnKeyPressed(e.Key);
                    UpdateLogView();
                }
            }
            catch (Exception ex)
            {
                LogService.Write("Hook_KeyPressed exception: " + ex.Message);
            }
        }

        private void MonitorToggle_Checked(object sender, RoutedEventArgs e)
        {
            _analyzer.BeginTypingExpectingKorean();
            StatusText.Text = "Korean IME Fixer — 모니터링: 켜짐";
            try { SaveConfig(_processWhitelist, _doubleEnterFixEnabled, true); } catch { }
            LogService.Write("Monitoring enabled via UI toggle.");
        }

        private void MonitorToggle_Unchecked(object sender, RoutedEventArgs e)
        {
            _analyzer.EndTypingExpectingKorean();
            StatusText.Text = "Korean IME Fixer — 모니터링: 꺼짐";
            try { SaveConfig(_processWhitelist, _doubleEnterFixEnabled, false); } catch { }
            LogService.Write("Monitoring disabled via UI toggle.");
        }

        private void TestRestoreButton_Click(object sender, RoutedEventArgs e)
        {
            LogService.Write("Manual restore button clicked.");
            _restorer.ForceToKorean();
            // Also perform UIA PoC to log focused/caret AutomationElement info for debugging Notion table detection
            try
            {
                LogFocusedAutomationElementInfo();
            }
            catch (Exception ex)
            {
                LogService.Write("LogFocusedAutomationElementInfo failed: " + ex.Message);
            }
            UpdateLogView();
        }

        private void OpenConfigFolder_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string dir = Path.GetDirectoryName(_configPath)!;
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                Process.Start(new ProcessStartInfo("explorer.exe", dir) { UseShellExecute = true });
            }
            catch { }
        }

        private void SaveWhitelistButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string text = WhitelistText.Text ?? string.Empty;
                var parts = text.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(s => s.Trim()).Where(s => s.Length > 0).ToArray();
                if (parts.Length > 0)
                {
                    _processWhitelist = parts;
                    SaveConfig(parts, _doubleEnterFixEnabled, MonitorToggle.IsChecked == true);
                    LogService.Write("Whitelist updated via UI: " + string.Join(",", parts));
                }
            }
            catch (Exception ex)
            {
                LogService.Write("SaveWhitelistButton_Click failed: " + ex.Message);
            }
        }

        private void FocusTimer_Tick(object? sender, EventArgs e)
        {
            try
            {
                IntPtr hwnd = GetForegroundWindow();
                if (hwnd == IntPtr.Zero)
                {
                    SetAutoEnabled(false);
                    return;
                }

                _ = GetWindowThreadProcessId(hwnd, out uint pid);
                string procName;
                try
                {
                    using Process p = Process.GetProcessById((int)pid);
                    procName = p.ProcessName.ToLowerInvariant();
                }
                catch
                {
                    procName = string.Empty;
                }

                bool match = _processWhitelist.Any(w => !string.IsNullOrEmpty(procName) && procName.Contains(w, StringComparison.OrdinalIgnoreCase));
                SetAutoEnabled(match);
            }
            catch { }
        }

        private void SetAutoEnabled(bool enable)
        {
            bool userChecked = MonitorToggle.IsChecked == true;
            if (userChecked) return;

            if (enable && !_autoEnabled)
            {
                _analyzer.BeginTypingExpectingKorean();
                _autoEnabled = true;
                StatusText.Text = "Korean IME Fixer — 모니터링: 자동(켜짐)";
                LogService.Write("Monitoring auto-enabled for foreground process.");
            }
            else if (!enable && _autoEnabled)
            {
                _analyzer.EndTypingExpectingKorean();
                _autoEnabled = false;
                StatusText.Text = "Korean IME Fixer — 모니터링: 자동(꺼짐)";
                LogService.Write("Monitoring auto-disabled as foreground process changed.");
            }
            UpdateLogView();
        }

        private void UpdateLogView()
        {
            try
            {
                string all = LogService.ReadAll() ?? string.Empty;
                var lines = all.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
                LogView.ItemsSource = lines;
            }
            catch { }
        }

        private void FallbackBufferTimer_Tick(object? sender, EventArgs e)
        {
            try
            {
                _fallbackBufferTimer.Stop();
                if (_bufferedFallbackJamo.HasValue)
                {
                    char j = _bufferedFallbackJamo.Value;
                    _bufferedFallbackJamo = null;
                    DateTime now = DateTime.UtcNow;
                    LogService.Write($"FallbackBufferTimer expired at {now:O} for jamo '{j}'. Trying clipboard-paste preferred path.");

                    // Prefer clipboard-based paste (with Ctrl+V fallback) on buffer expiry to improve delivery to apps like Notion.
                    bool sent = false;
                    try
                    {
                        sent = _restorer.PasteStringToForeground(j.ToString());
                        LogService.Write($"PasteStringToForeground result={sent}");
                    }
                    catch (Exception ex)
                    {
                        LogService.Write("PasteStringToForeground threw: " + ex.Message);
                        sent = false;
                    }

                    // If clipboard-paste path failed, fall back to existing Unicode SendInput path
                    if (!sent)
                    {
                        sent = _restorer.SendUnicodeChar(j);
                        LogService.Write($"FallbackBuffer: SendUnicodeChar fallback for '{j}' success={sent}");
                    }
                    if (sent)
                    {
                        _analyzer.EndTypingExpectingKorean();
                        _lastRestore = DateTime.UtcNow;
                        UpdateLogView();
                    }
                }
            }
            catch (Exception ex)
            {
                LogService.Write("FallbackBufferTimer_Tick failed: " + ex.Message);
            }
        }

        private void LoadWhitelist()
        {
            try
            {
                string dir = Path.GetDirectoryName(_configPath)!;
                if (File.Exists(_configPath))
                {
                    string json = File.ReadAllText(_configPath);
                    var doc = JsonSerializer.Deserialize<ConfigModel>(json);
                    if (doc != null && doc.Whitelist != null && doc.Whitelist.Length > 0)
                    {
                        _processWhitelist = doc.Whitelist;
                    }
                }
                WhitelistText.Text = string.Join(", ", _processWhitelist);
            }
            catch (Exception ex)
            {
                LogService.Write("LoadWhitelist failed: " + ex.Message);
            }
        }

        private void SaveWhitelist(string[] list)
        {
            try
            {
                string dir = Path.GetDirectoryName(_configPath)!;
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                // Persist whitelist together with other config flags
                SaveConfig(list, _doubleEnterFixEnabled);
                LogService.Write("Whitelist saved.");
            }
            catch (Exception ex)
            {
                LogService.Write("SaveWhitelist failed: " + ex.Message);
            }
        }

        private class ConfigModel { public string[]? Whitelist { get; set; } public bool DoubleEnterFixEnabled { get; set; } public bool MonitorEnabled { get; set; } }

        private void SaveConfig(string[]? whitelist, bool doubleEnterFix, bool monitorEnabled = false)
        {
            try
            {
                string dir = Path.GetDirectoryName(_configPath)!;
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                var model = new ConfigModel { Whitelist = whitelist ?? _processWhitelist, DoubleEnterFixEnabled = doubleEnterFix, MonitorEnabled = monitorEnabled };
                string json = JsonSerializer.Serialize(model);
                File.WriteAllText(_configPath, json);
                _doubleEnterFixEnabled = doubleEnterFix;
                DoubleEnterToggle.IsChecked = doubleEnterFix;
                MonitorToggle.IsChecked = monitorEnabled;
                LogService.Write("Config saved.");
            }
            catch (Exception ex)
            {
                LogService.Write("SaveConfig failed: " + ex.Message);
            }
        }

        private void LoadConfig()
        {
            try
            {
                if (File.Exists(_configPath))
                {
                    string json = File.ReadAllText(_configPath);
                    var doc = JsonSerializer.Deserialize<ConfigModel>(json);
                    if (doc != null)
                    {
                        if (doc.Whitelist != null && doc.Whitelist.Length > 0) _processWhitelist = doc.Whitelist;
                        _doubleEnterFixEnabled = doc.DoubleEnterFixEnabled;
                        bool monitor = doc.MonitorEnabled;
                        // reflect monitor state
                        MonitorToggle.IsChecked = monitor;
                        if (monitor)
                        {
                            _analyzer.BeginTypingExpectingKorean();
                        }
                    }
                }
                // update UI
                WhitelistText.Text = string.Join(", ", _processWhitelist);
                DoubleEnterToggle.IsChecked = _doubleEnterFixEnabled;
            }
            catch (Exception ex)
            {
                LogService.Write("LoadConfig failed: " + ex.Message);
            }
        }

        private void DoubleEnterToggle_Checked(object sender, RoutedEventArgs e)
        {
            _doubleEnterFixEnabled = true;
            try { SaveConfig(_processWhitelist, _doubleEnterFixEnabled); } catch { }
            LogService.Write("DoubleEnterFix enabled via UI.");
        }

        private void DoubleEnterToggle_Unchecked(object sender, RoutedEventArgs e)
        {
            _doubleEnterFixEnabled = false;
            try { SaveConfig(_processWhitelist, _doubleEnterFixEnabled); } catch { }
            LogService.Write("DoubleEnterFix disabled via UI.");
        }

        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll")]
        private static extern bool GetGUIThreadInfo(uint idThread, ref GUITHREADINFO lpgui);

        [StructLayout(LayoutKind.Sequential)]
        private struct GUITHREADINFO
        {
            public int cbSize;
            public uint flags;
            public IntPtr hwndActive;
            public IntPtr hwndFocus;
            public IntPtr hwndCapture;
            public IntPtr hwndMenuOwner;
            public IntPtr hwndMoveSize;
            public IntPtr hwndCaret;
            public RECT rcCaret;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        private bool WaitForCaret(uint threadId, int timeoutMs = 400)
        {
            try
            {
                var info = new GUITHREADINFO();
                info.cbSize = Marshal.SizeOf(typeof(GUITHREADINFO));
                int waited = 0;
                const int step = 30;
                while (waited < timeoutMs)
                {
                    bool ok = GetGUIThreadInfo(threadId, ref info);
                    if (ok)
                    {
                        if (info.hwndCaret != IntPtr.Zero)
                        {
                            int w = info.rcCaret.Right - info.rcCaret.Left;
                            int h = info.rcCaret.Bottom - info.rcCaret.Top;
                            if (w > 0 || h > 0) return true;
                        }
                    }

                    System.Threading.Thread.Sleep(step);
                    waited += step;
                }
            }
            catch { }
            return false;
        }

        // Use helpers for AutomationElement logging and table-detection to keep MainWindow small
        private void LogFocusedAutomationElementInfo()
        {
            try { AutomationHelpers.LogFocusedAutomationElementInfo(); } catch (Exception ex) { LogService.Write("LogFocusedAutomationElementInfo wrapper failed: " + ex.Message); }
        }

        private bool IsLikelyNotionTable()
        {
            try { return AutomationHelpers.IsLikelyNotionTable(); } catch (Exception ex) { LogService.Write("IsLikelyNotionTable wrapper failed: " + ex.Message); return false; }
        }
    }
}