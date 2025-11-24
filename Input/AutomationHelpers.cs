using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Diagnostics;
using System.Windows.Automation;
using KoreanIMEFixer.Logging;

namespace KoreanIMEFixer
{
    internal static class AutomationHelpers
    {
        // Use centralized native helpers in NativeTypes.cs to avoid duplicate P/Invoke/structs

        // PoC: log AutomationElement details at caret or focused element to detect table/cell contexts (Notion)
        public static void LogFocusedAutomationElementInfo()
        {
            try
            {
                IntPtr fg = NativeMethods.GetForegroundWindow();
                if (fg == IntPtr.Zero)
                {
                    LogService.Write("LogFocusedAutomationElementInfo: no foreground window.");
                    return;
                }
                _ = NativeMethods.GetWindowThreadProcessId(fg, out uint tid);
                AutomationElement? el = null;

                // Try caret point
                try
                {
                    var gti = new GUITHREADINFO();
                    gti.cbSize = Marshal.SizeOf(typeof(GUITHREADINFO));
                    if (tid != 0 && NativeMethods.GetGUIThreadInfo(tid, ref gti))
                    {
                        var rect = gti.rcCaret;
                        if (rect.Right > rect.Left && rect.Bottom > rect.Top)
                        {
                            int cx = rect.Left + (rect.Right - rect.Left) / 2;
                            int cy = rect.Top + (rect.Bottom - rect.Top) / 2;
                            var pt = new System.Windows.Point(cx, cy);
                            try { el = AutomationElement.FromPoint(pt); }
                            catch { el = null; }
                            LogService.Write($"LogFocusedAutomationElementInfo: caret point = ({cx},{cy}), element={(el==null?"null":"found")}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogService.Write("LogFocusedAutomationElementInfo: caret lookup failed: " + ex.Message);
                }

                if (el == null)
                {
                    try { el = AutomationElement.FocusedElement; } catch { el = null; }
                }

                if (el == null)
                {
                    LogService.Write("LogFocusedAutomationElementInfo: no AutomationElement available.");
                    return;
                }

                void LogEl(AutomationElement a, string tag)
                {
                    try
                    {
                        string ctrl = a.Current.ControlType?.ProgrammaticName ?? "?";
                        string loc = a.Current.LocalizedControlType ?? "?";
                        string name = a.Current.Name ?? "";
                        string cls = a.Current.ClassName ?? "";
                        string aid = a.Current.AutomationId ?? "";
                        var rect = a.Current.BoundingRectangle;
                        string rectStr = rect.IsEmpty ? "empty" : $"L={rect.Left:0},T={rect.Top:0},W={rect.Width:0},H={rect.Height:0}";
                        LogService.Write($"AE {tag}: ControlType={ctrl}, LocalizedType={loc}, Name='{name}', ClassName='{cls}', AutomationId='{aid}', Bounding={rectStr}");
                        bool hasTable = a.TryGetCurrentPattern(TablePattern.Pattern, out _);
                        bool hasGrid = a.TryGetCurrentPattern(GridPattern.Pattern, out _);
                        bool hasText = a.TryGetCurrentPattern(TextPattern.Pattern, out _);
                        bool hasValue = a.TryGetCurrentPattern(ValuePattern.Pattern, out _);
                        LogService.Write($"AE {tag}: patterns: Table={hasTable}, Grid={hasGrid}, Text={hasText}, Value={hasValue}");

                        try
                        {
                            var children = a.FindAll(TreeScope.Children, System.Windows.Automation.Condition.TrueCondition);
                            if (children != null)
                            {
                                int cc = children.Count;
                                LogService.Write($"AE {tag}: childCount={cc}");
                                int limit = Math.Min(cc, 8);
                                for (int i = 0; i < limit; i++)
                                {
                                    try
                                    {
                                        var c = children[i];
                                        var r = c.Current.BoundingRectangle;
                                        string rs = r.IsEmpty ? "empty" : $"L={r.Left:0},T={r.Top:0},W={r.Width:0},H={r.Height:0}";
                                        LogService.Write($"AE {tag}: child[{i}] Ctrl={c.Current.ControlType?.ProgrammaticName ?? "?"} Name='{c.Current.Name}' Bounding={rs}");
                                    }
                                    catch { }
                                }
                            }
                        }
                        catch (Exception exChildren)
                        {
                            LogService.Write($"AE {tag}: child enumeration failed: {exChildren.Message}");
                        }
                    }
                    catch (Exception ex)
                    {
                        LogService.Write("AE LogEl failed: " + ex.Message);
                    }
                }

                AutomationElement? cur = el;
                for (int i = 0; i < 7 && cur != null; i++)
                {
                    LogEl(cur, "depth" + i);
                    try { cur = TreeWalker.ControlViewWalker.GetParent(cur); } catch { cur = null; }
                }
            }
            catch (Exception ex)
            {
                LogService.Write("LogFocusedAutomationElementInfo outer failed: " + ex.Message);
            }
        }

        // Heuristic: return true when the focused/caret AutomationElement appears to contain a table-like layout
        // This uses descendant bounding-rectangle clustering (columns/rows) to guess a table in hosts that don't expose TablePattern (e.g., Notion)
        public static bool IsLikelyNotionTable()
        {
            var sw = Stopwatch.StartNew();
            try
            {
                AutomationElement? el = null;
                try
                {
                    IntPtr fg = NativeMethods.GetForegroundWindow();
                    if (fg == IntPtr.Zero) return false;
                    _ = NativeMethods.GetWindowThreadProcessId(fg, out uint tid);
                    var gti = new GUITHREADINFO();
                    gti.cbSize = Marshal.SizeOf(typeof(GUITHREADINFO));
                    if (tid != 0 && NativeMethods.GetGUIThreadInfo(tid, ref gti))
                    {
                        var rect = gti.rcCaret;
                        if (rect.Right > rect.Left && rect.Bottom > rect.Top)
                        {
                            int cx = rect.Left + (rect.Right - rect.Left) / 2;
                            int cy = rect.Top + (rect.Bottom - rect.Top) / 2;
                            try { el = AutomationElement.FromPoint(new System.Windows.Point(cx, cy)); } catch { el = null; }
                        }
                    }
                }
                catch { el = null; }

                if (el == null)
                {
                    try { el = AutomationElement.FocusedElement; } catch { el = null; }
                }

                if (el == null) return false;

                var docRect = el!.Current.BoundingRectangle;
                if (docRect.IsEmpty)
                {
                    try
                    {
                        var p = TreeWalker.ControlViewWalker.GetParent(el!);
                        int depth = 0;
                        while (p != null && depth < 6)
                        {
                            if (!p.Current.BoundingRectangle.IsEmpty)
                            {
                                docRect = p.Current.BoundingRectangle;
                                break;
                            }
                            p = TreeWalker.ControlViewWalker.GetParent(p);
                            depth++;
                        }
                    }
                    catch { }
                }

                if (docRect.IsEmpty) return false;

                // Bounded traversal: sample descendant bounding-rectangles without enumerating all
                var rects = new List<System.Windows.Rect>();
                const int MaxSample = 100; // how many candidate rects to collect
                const int MaxVisited = 1000; // cap how many nodes we traverse
                try
                {
                    var walker = TreeWalker.ControlViewWalker;
                    var q = new Queue<AutomationElement>();
                    int visited = 0;
                    try
                    {
                        var first = walker.GetFirstChild(el!);
                        if (first != null) q.Enqueue(first);
                        while (q.Count > 0 && rects.Count < MaxSample && visited < MaxVisited)
                        {
                            var node = q.Dequeue();
                            visited++;
                            try
                            {
                                var r = node.Current.BoundingRectangle;
                                if (!r.IsEmpty && r.Width > 2 && r.Height > 2)
                                {
                                    if (!(Math.Abs(r.Width - docRect.Width) < 4 && Math.Abs(r.Height - docRect.Height) < 4))
                                    {
                                        rects.Add(r);
                                    }
                                }
                            }
                            catch { }

                            try
                            {
                                var child = walker.GetFirstChild(node);
                                while (child != null)
                                {
                                    q.Enqueue(child);
                                    child = walker.GetNextSibling(child);
                                }
                            }
                            catch { }
                        }
                    }
                    catch (Exception exTrv)
                    {
                        LogService.Write("IsLikelyNotionTable: traversal failed: " + exTrv.Message);
                    }
                }
                catch (Exception exTraverse)
                {
                    LogService.Write("IsLikelyNotionTable: sampling failed: " + exTraverse.Message);
                }

                LogService.Write($"IsLikelyNotionTable: docRect L={docRect.Left:0},T={docRect.Top:0},W={docRect.Width:0},H={docRect.Height:0}; descendantsConsidered={rects.Count}");

                if (rects.Count < 6) return false;

                double docW = Math.Max(1.0, docRect.Width);

                var candidates = rects.Where(r => r.Width < docW * 0.95 && r.Width > docW * 0.05 && r.Height < 240).ToList();
                LogService.Write($"IsLikelyNotionTable: candidates={candidates.Count}");
                if (candidates.Count < 6) return false;

                int round = 16;
                var leftGroups = new Dictionary<int, List<System.Windows.Rect>>();
                var topGroups = new Dictionary<int, List<System.Windows.Rect>>();
                foreach (var r in candidates)
                {
                    int l = (int)(Math.Round(r.Left / round) * round);
                    int t = (int)(Math.Round(r.Top / round) * round);
                    if (!leftGroups.TryGetValue(l, out var ll)) { ll = new List<System.Windows.Rect>(); leftGroups[l] = ll; }
                    ll.Add(r);
                    if (!topGroups.TryGetValue(t, out var tt)) { tt = new List<System.Windows.Rect>(); topGroups[t] = tt; }
                    tt.Add(r);
                }

                int columns = leftGroups.Keys.Count;
                int rows = topGroups.Keys.Count;
                LogService.Write($"IsLikelyNotionTable: found columns={columns}, rows={rows}");

                if (columns >= 2 && rows >= 2 && candidates.Count >= Math.Max(6, columns * 2))
                {
                    int multiCols = leftGroups.Values.Count(g => g.Count >= 2);
                    if (multiCols >= 2) return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                LogService.Write("IsLikelyNotionTable exception: " + ex.Message);
                return false;
            }
        }
    }
}
