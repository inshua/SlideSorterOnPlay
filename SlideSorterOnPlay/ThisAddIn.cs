using Microsoft.Office.Interop.PowerPoint;
using SlideSorterOnPlay.Properties;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing.Imaging;
using System.Linq;
using System.Reflection.Emit;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace SlideSorterOnPlay
{
    public partial class ThisAddIn
    {

        private PpViewType _originalView = PpViewType.ppViewNormal;
        private SlideShowWindow _slideShowWin;
        private LowLevelKeyboardProc _proc;
        private IntPtr _hookID = IntPtr.Zero;

        // ===== Win32 API for keyboard hook =====
        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private const int WM_KEYUP = 0x0101;
        private const int WM_CHAR = 0x0102;

        [DllImport("user32.dll")]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn,
            IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll")]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode,
            IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll")]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        private delegate IntPtr HookProc(int code, IntPtr wParam, IntPtr lParam);

        private const int WH_GETMESSAGE = 3;

        [StructLayout(LayoutKind.Sequential)]
        private struct MSG
        {
            public IntPtr hwnd;
            public uint message;
            public IntPtr wParam;
            public IntPtr lParam;
            public uint time;
            public int pt_x;
            public int pt_y;
        }

        [DllImport("user32.dll")]
        private static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn,
            IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, IntPtr lpdwProcessId);

        private void InstallThreadHook()
        {
            Debug.WriteLine("Installing thread hook");

            _proc = HookCallback;
            IntPtr hwnd = new IntPtr(this.Application.HWND);
            uint threadId = GetWindowThreadProcessId(hwnd, IntPtr.Zero);

            _hookID = SetWindowsHookEx(WH_GETMESSAGE, _proc, IntPtr.Zero, threadId);
        }

        private void RemoveThreadHook()
        {
            if (_hookID != IntPtr.Zero)
            {
                UnhookWindowsHookEx(_hookID);
                _hookID = IntPtr.Zero;
            }
        }


        private long lastKeyTime = 0;
        private DocumentWindow edtWin;

        private IntPtr HookCallback(int code, IntPtr wParam, IntPtr lParam)
        {
            System.Diagnostics.Debug.WriteLine("Hook Callback: {0:X} {1:X} {2:X}", code, wParam, lParam);
            if (code >= 0 && (int)wParam > 0)
            {
                MSG msg = (MSG)Marshal.PtrToStructure(lParam, typeof(MSG));
                if (_slideShowWin != null)
                {
                    if(!Settings.Default.Enabled)
                    {
                        return CallNextHookEx(_hookID, code, wParam, lParam);
                    }

                    if (msg.message == WM_KEYDOWN)
                    {
                        Keys key = (Keys)(int)msg.wParam;
                        Debug.WriteLine("Captured key: " + key);

                        if (key == Keys.Escape)
                        {
                            _slideShowWin.View.Exit();
                            return (IntPtr)1;
                        }
                    } 
                    //else if(msg.message == WM_CHAR){
                    //    var now = DateTimeOffset.Now.ToUnixTimeMilliseconds();
                    //    Debug.WriteLine("Char message time: " + now + ", lastKeyTime: " + lastKeyTime);
                    //    if (now - lastKeyTime > 250)
                    //    {
                    //        // 防止按键连发
                    //        lastKeyTime = now;
                    //        char ch = (char)(int)msg.wParam;
                    //        Debug.WriteLine("Captured char: " + ch);
                    //        //if (ch == ' ')
                    //        {
                    //            GotoNextSlide();
                    //            return (IntPtr)1;
                    //        }
                    //    }
                    //}
                }
            }
            return CallNextHookEx(_hookID, code, wParam, lParam);
        }

        private void GotoNextSlide()
        {
            Debug.WriteLine("GotoNextSlide called");
            var win = Globals.ThisAddIn.Application.ActiveWindow;
            //if (win.ViewType == PowerPoint.PpViewType.ppViewSlideSorter)
            {
                var currentSlide = win.View.Slide;
                if (currentSlide != null)
                {
                    int nextIndex = currentSlide.SlideIndex + 1;
                    if (nextIndex <= Globals.ThisAddIn.Application.ActivePresentation.Slides.Count)
                    {
                        var nextSlide = Globals.ThisAddIn.Application.ActivePresentation.Slides[nextIndex];
                        nextSlide.Select();
                    }
                }
            }
        }


        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            if (Settings.Default.Enabled)
            {
                InstallEvents();
            }
        }

        public void InstallEvents()
        {
            this.Application.SlideShowBegin += Application_SlideShowBegin;
            this.Application.SlideShowEnd += Application_SlideShowEnd;
            this.Application.WindowSelectionChange += Application_WindowSelectionChange;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            if (Settings.Default.Enabled)
            {
                UninstallEvents();
            }
        }

        public void UninstallEvents()
        {
            this.Application.SlideShowBegin -= Application_SlideShowBegin;
            this.Application.SlideShowEnd -= Application_SlideShowEnd;
            this.Application.WindowSelectionChange -= Application_WindowSelectionChange;
        }

        private void Application_SlideShowBegin(SlideShowWindow Wn)
        {
            try
            {
                _slideShowWin = Wn;
                InstallThreadHook();

                if (this.Application.ActivePresentation.Windows.Count > 0)
                {
                    var windows = this.Application.ActivePresentation.Windows;
                    for(var i= 1; i <= windows.Count; i++)
                    {
                        var editWin = windows[i];
                        try
                        {
                            _originalView = editWin.ViewType;
                            editWin.ViewType = PpViewType.ppViewSlideSorter;
                            this.edtWin = editWin;
                        }
                        catch { }
                    }
                    
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("SlideShowBegin error: " + ex.Message);
            }
        }

        private void Application_SlideShowEnd(Presentation Pres)
        {
            try
            {
                _slideShowWin = null;
                RemoveThreadHook();

                try
                {
                    edtWin.ViewType = _originalView;
                }
                catch { 
                    if (Pres.Windows.Count > 0)
                    {
                        var editWin = Pres.Windows[1];
                        editWin.ViewType = _originalView;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("SlideShowEnd error: " + ex.Message);
            }
        }

        private void Application_WindowSelectionChange(Selection Sel)
        {
            try
            {
                if (_slideShowWin != null && Sel.Type == PpSelectionType.ppSelectionSlides)
                {
                    var w = _slideShowWin;
                    // 取第一个选中的幻灯片
                    var slide = Sel.SlideRange[1];
                    if (slide != null)
                    {

                        var slideId = slide.SlideID;
                        if (w.View.Slide.SlideID == slideId) return;

                        // 放映窗口跳转到该幻灯片
                        if (w != null) w.View.GotoSlide(slide.SlideIndex);
                        this.edtWin = Application.ActiveWindow;
                        if (hasVideo(slide)) {
                            w.Activate();
                            StartFocusRestoreTimer(slide);
                            //while(w.Active != Office.MsoTriState.msoFalse) {
                            //    System.Windows.Forms.Application.DoEvents();
                            //    if(w.View.State == PpSlideShowState.ppSlideShowDone || w.View.Slide.SlideID != slideId)
                            //    {
                            //        break;
                            //    }
                            //}
                            //edtWin.Activate();
                            //Debug.WriteLine("Activated edit window after video");
                        } else if(edtWin != null && edtWin.Active == Office.MsoTriState.msoFalse) {
                            StopTimer();
                            edtWin.Activate();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("SelectionChange error: " + ex.Message);
            }
        }

        private static bool hasVideo(Slide slide)
        {
            foreach(Shape shape in slide.Shapes)
            {
                if (shape.Type == Office.MsoShapeType.msoMedia && shape.MediaType == PpMediaType.ppMediaTypeMovie)
                {
                    return true;
                }
            }
            return false;
        }

        private System.Windows.Forms.Timer _focusTimer;

        private void StartFocusRestoreTimer(Slide slide)
        {
            if (_focusTimer != null)
            {
                _focusTimer.Stop();
                _focusTimer.Dispose();
            }

            _focusTimer = new System.Windows.Forms.Timer();
            _focusTimer.Interval = 200; // 300ms 轮询

            _focusTimer.Tick += (s, e) =>
            {
                try
                {
                    var view = _slideShowWin?.View;
                    if (view == null)
                    {
                        StopTimer();
                        return;
                    }

                    // 判断是否还有动画在播放
                    bool stillPlaying = view.Slide == slide && view.State == PpSlideShowState.ppSlideShowRunning
                                        && view.CurrentShowPosition == slide.SlideIndex;

                    // 👉 关键：检测动画是否结束
                    if (!stillPlaying || !HasRunningAnimation(slide))
                    {
                        RestoreFocus();
                        StopTimer();
                    }
                }
                catch
                {
                    StopTimer();
                }
            };

            _focusTimer.Start();
        }

        private void RestoreFocus()
        {
            try
            {
                if(this.edtWin != null) {
                    Trace.WriteLine("Restoring focus to edit window " + this.edtWin.Caption);
                }
                this.edtWin?.Activate();
            }
            catch { }
        }

        private bool HasRunningAnimation(Slide slide)
        {
            var timeline = slide.TimeLine;
            if (timeline.MainSequence.Count == 0)
                return false;

            foreach (Effect effect in timeline.MainSequence)
            {
                if (effect.Timing.Duration > 0)
                {
                    return true;
                }
            }

            return false;
        }

        private void StopTimer()
        {
            if (_focusTimer != null)
            {
                _focusTimer.Stop();
                _focusTimer.Dispose();
                _focusTimer = null;
            }
        }

        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
    }

}
