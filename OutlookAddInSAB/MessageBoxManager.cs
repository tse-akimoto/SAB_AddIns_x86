using System;
using System.Text;
using System.Runtime.InteropServices;
using System.Security.Permissions;

#pragma warning disable 0618
[assembly: SecurityPermission(SecurityAction.RequestMinimum, UnmanagedCode = true)]
namespace System.Windows.Forms
{
    public class MessageBoxManager
    {
        #region 定義

        private delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);
        private delegate bool EnumChildProc(IntPtr hWnd, IntPtr lParam);

        private const int WH_CALLWNDPROCRET = 12;
        private const int WM_DESTROY = 0x0002;
        private const int WM_INITDIALOG = 0x0110;
        private const int WM_TIMER = 0x0113;
        private const int WM_USER = 0x400;
        private const int DM_GETDEFID = WM_USER + 0;

        private const int MBOK = 1;
        private const int MBCancel = 2;
        private const int MBAbort = 3;
        private const int MBRetry = 4;
        private const int MBIgnore = 5;
        private const int MBYes = 6;
        private const int MBNo = 7;


        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hInstance, int threadId);

        [DllImport("user32.dll")]
        private static extern int UnhookWindowsHookEx(IntPtr idHook);

        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(IntPtr idHook, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", EntryPoint = "GetWindowTextLengthW", CharSet = CharSet.Unicode)]
        private static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll", EntryPoint = "GetWindowTextW", CharSet = CharSet.Unicode)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int maxLength);

        [DllImport("user32.dll")]
        private static extern int EndDialog(IntPtr hDlg, IntPtr nResult);

        [DllImport("user32.dll")]
        private static extern bool EnumChildWindows(IntPtr hWndParent, EnumChildProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll", EntryPoint = "GetClassNameW", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        private static extern int GetDlgCtrlID(IntPtr hwndCtl);

        [DllImport("user32.dll")]
        private static extern IntPtr GetDlgItem(IntPtr hDlg, int nIDDlgItem);

        [DllImport("user32.dll", EntryPoint = "SetWindowTextW", CharSet = CharSet.Unicode)]
        private static extern bool SetWindowText(IntPtr hWnd, string lpString);


        [StructLayout(LayoutKind.Sequential)]
        public struct CWPRETSTRUCT
        {
            public IntPtr lResult;
            public IntPtr lParam;
            public IntPtr wParam;
            public uint message;
            public IntPtr hwnd;
        };

        private static HookProc hookProc;
        private static EnumChildProc enumProc;
        [ThreadStatic]
        private static IntPtr hHook;
        [ThreadStatic]
        private static int nButton;

        /// <summary>
        /// OK text
        /// </summary>
        public static string DefaultOK = "";
        /// <summary>
        /// Cancel text
        /// </summary>
        public static string DefaultCancel = "";
        /// <summary>
        /// Abort text
        /// </summary>
        public static string DefaultAbort = "";
        /// <summary>
        /// Retry text
        /// </summary>
        public static string DefaultRetry = "";
        /// <summary>
        /// Ignore text
        /// </summary>
        public static string DefaultIgnore = "";
        /// <summary>
        /// Yes text
        /// </summary>
        public static string DefaultYes = "";
        /// <summary>
        /// No text
        /// </summary>
        public static string DefaultNo = "";



        /// <summary>
        /// OK text
        /// </summary>
        public static string OK = "&OK";
        /// <summary>
        /// Cancel text
        /// </summary>
        public static string Cancel = "&Cancel";
        /// <summary>
        /// Abort text
        /// </summary>
        public static string Abort = "&Abort";
        /// <summary>
        /// Retry text
        /// </summary>
        public static string Retry = "&Retry";
        /// <summary>
        /// Ignore text
        /// </summary>
        public static string Ignore = "&Ignore";
        /// <summary>
        /// Yes text
        /// </summary>
        public static string Yes = "&Yes";
        /// <summary>
        /// No text
        /// </summary>
        public static string No = "&No";

        #endregion


        static MessageBoxManager()
        {
            hookProc = new HookProc(MessageBoxHookProc);
            enumProc = new EnumChildProc(MessageBoxEnumProc);
            hHook = IntPtr.Zero;
        }

        #region メソッド

        /// <summary>
        /// Enables MessageBoxManager functionality
        /// </summary>
        /// <remarks>
        /// MessageBoxManager functionality is enabled on current thread only.
        /// Each thread that needs MessageBoxManager functionality has to call this method.
        /// </remarks>
        public static void Register()
        {
            if (hHook != IntPtr.Zero)
                throw new NotSupportedException("One hook per thread allowed.");
            hHook = SetWindowsHookEx(WH_CALLWNDPROCRET, hookProc, IntPtr.Zero, AppDomain.GetCurrentThreadId());
        }

        /// <summary>
        /// ダイアログに表示するボタンの初期化
        /// </summary>
        public static void InitMessageBoxManager()
        {
            OK = "";
            Cancel = "";
            Abort = "";
            Retry = "";
            Ignore = "";
            Yes = "";
            No = "";

            Unregister();
        }

        /// <summary>
        /// ダイアログに表示するボタン
        /// </summary>
        public static void ResetText()
        {
            if (string.IsNullOrEmpty(DefaultOK) == false)
            {
                OK = DefaultOK;
            }

            if (string.IsNullOrEmpty(DefaultCancel) == false)
            {
                Cancel = DefaultCancel;
            }

            if (string.IsNullOrEmpty(DefaultAbort) == false)
            {
                Abort = DefaultAbort;
            }

            if (string.IsNullOrEmpty(DefaultRetry) == false)
            {
                Retry = DefaultRetry;
            }

            if (string.IsNullOrEmpty(DefaultIgnore) == false)
            {
                Ignore = DefaultIgnore;
            }

            if (string.IsNullOrEmpty(DefaultYes) == false)
            {
                Yes = DefaultYes;
            }

            if (string.IsNullOrEmpty(DefaultNo) == false)
            {
                No = DefaultNo;
            }

            // 再登録
            Unregister();
            Register();

            // 終了
            Unregister();
        }

        /// <summary>
        /// Disables MessageBoxManager functionality
        /// </summary>
        /// <remarks>
        /// Disables MessageBoxManager functionality on current thread only.
        /// </remarks>
        public static void Unregister()
        {
            if (hHook != IntPtr.Zero)
            {
                UnhookWindowsHookEx(hHook);
                hHook = IntPtr.Zero;
            }
        }

        private static IntPtr MessageBoxHookProc(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode < 0)
                return CallNextHookEx(hHook, nCode, wParam, lParam);

            CWPRETSTRUCT msg = (CWPRETSTRUCT)Marshal.PtrToStructure(lParam, typeof(CWPRETSTRUCT));
            IntPtr hook = hHook;

            if (msg.message == WM_INITDIALOG)
            {
                int nLength = GetWindowTextLength(msg.hwnd);
                StringBuilder className = new StringBuilder(10);
                GetClassName(msg.hwnd, className, className.Capacity);
                if (className.ToString() == "#32770")
                {
                    nButton = 0;
                    EnumChildWindows(msg.hwnd, enumProc, IntPtr.Zero);
                    if (nButton == 1)
                    {
                        IntPtr hButton = GetDlgItem(msg.hwnd, MBCancel);
                        if (hButton != IntPtr.Zero)
                            SetWindowText(hButton, OK);
                    }
                }
            }

            return CallNextHookEx(hook, nCode, wParam, lParam);
        }

        private static bool MessageBoxEnumProc(IntPtr hWnd, IntPtr lParam)
        {
            StringBuilder className = new StringBuilder(10);
            GetClassName(hWnd, className, className.Capacity);
            if (className.ToString() == "Button")
            {

                int ctlId = GetDlgCtrlID(hWnd);

                int textLen = GetWindowTextLength(hWnd);
                StringBuilder tsb = new StringBuilder(textLen + 1);
                switch (ctlId)
                {
                    case MBOK:
                        GetWindowText(hWnd, tsb, tsb.Capacity);
                        if (string.IsNullOrEmpty(DefaultOK) != false)
                        {
                            DefaultOK = tsb.ToString();
                        }

                        if (string.IsNullOrEmpty(OK) != false)
                        {
                            OK = DefaultOK;
                        }

                        SetWindowText(hWnd, OK);
                        break;
                    case MBCancel:
                        GetWindowText(hWnd, tsb, tsb.Capacity);
                        if (string.IsNullOrEmpty(DefaultCancel) != false)
                        {
                            DefaultCancel = tsb.ToString();
                        }

                        if (string.IsNullOrEmpty(Cancel) != false)
                        {
                            Cancel = DefaultCancel;
                        }

                        SetWindowText(hWnd, Cancel);
                        break;
                    case MBAbort:
                        GetWindowText(hWnd, tsb, tsb.Capacity);
                        if (string.IsNullOrEmpty(DefaultAbort) != false)
                        {
                            DefaultAbort = tsb.ToString();
                        }

                        if (string.IsNullOrEmpty(Abort) != false)
                        {
                            Abort = DefaultAbort;
                        }

                        SetWindowText(hWnd, Abort);
                        break;
                    case MBRetry:
                        GetWindowText(hWnd, tsb, tsb.Capacity);
                        if (string.IsNullOrEmpty(DefaultRetry) != false)
                        {
                            DefaultRetry = tsb.ToString();
                        }

                        if (string.IsNullOrEmpty(Retry) != false)
                        {
                            Retry = DefaultRetry;
                        }

                        SetWindowText(hWnd, Retry);
                        break;
                    case MBIgnore:
                        GetWindowText(hWnd, tsb, tsb.Capacity);
                        if (string.IsNullOrEmpty(DefaultIgnore) != false)
                        {
                            DefaultIgnore = tsb.ToString();
                        }

                        if (string.IsNullOrEmpty(Ignore) != false)
                        {
                            Ignore = DefaultIgnore;
                        }

                        SetWindowText(hWnd, Ignore);
                        break;
                    case MBYes:
                        GetWindowText(hWnd, tsb, tsb.Capacity);
                        if (string.IsNullOrEmpty(DefaultYes) != false)
                        {
                            DefaultYes = tsb.ToString();
                        }

                        if (string.IsNullOrEmpty(Yes) != false)
                        {
                            Yes = DefaultYes;
                        }

                        SetWindowText(hWnd, Yes);
                        break;
                    case MBNo:
                        GetWindowText(hWnd, tsb, tsb.Capacity);
                        if (string.IsNullOrEmpty(DefaultNo) != false)
                        {
                            DefaultNo = tsb.ToString();
                        }

                        if (string.IsNullOrEmpty(No) != false)
                        {
                            No = DefaultNo;
                        }

                        SetWindowText(hWnd, No);
                        break;

                }
                nButton++;
            }

            return true;
        }

        #endregion

        #region クラス：Messagebox

        public class Messagebox
        {
            public Messagebox()
            {

            }

            /// <summary>
            /// ダイアログに表示するメッセージ一覧
            /// </summary>
            public DialogResult Show_MessageBox(int num)
            {
                DialogResult ret = DialogResult.None;
                switch (num)
                {
                    case 1:
                        ret = MessageBox.Show(AddInsLibrary.Properties.Resources.msg_err_not_stamp, AddInsLibrary.Properties.Resources.msgError, MessageBoxButtons.OK, MessageBoxIcon.Error);
                        break;
                    case 2:
                        ret = MessageBox.Show(AddInsLibrary.Properties.Resources.msg_err_includingS, AddInsLibrary.Properties.Resources.msgError, MessageBoxButtons.OK, MessageBoxIcon.Error);
                        break;
                    case 3:
                        ret = MessageBox.Show(AddInsLibrary.Properties.Resources.msg_qst_send_authorization, AddInsLibrary.Properties.Resources.msgConfirm, MessageBoxButtons.YesNo, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
                        break;
                    case 4:
                        ret = MessageBox.Show(AddInsLibrary.Properties.Resources.msg_err_not_authorization, AddInsLibrary.Properties.Resources.msgError, MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
                        break;
                    case 5:
                        ret = MessageBox.Show(AddInsLibrary.Properties.Resources.msg_qst_send_check, AddInsLibrary.Properties.Resources.msgError, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Exclamation);
                        break;
                    case 6:
                        ret = MessageBox.Show(AddInsLibrary.Properties.Resources.msg_qst_create_zip, AddInsLibrary.Properties.Resources.msgConfirm, MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                        break;
                    case 7:
                        ret = MessageBox.Show(AddInsLibrary.Properties.Resources.msg_info_zip_send, AddInsLibrary.Properties.Resources.msgConfirm, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        break;
                }

                return ret;
            }
        }

        #endregion
    }
}