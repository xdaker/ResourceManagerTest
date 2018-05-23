using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ClassLibrary1.FIleShell
{
    #region API 导入

    public class API
    {


        public const int MAX_PATH = 260;
        public const int S_OK = 0;
        public const int S_FALSE = 1;
        public const uint CMD_FIRST = 1;
        public const uint CMD_LAST = 30000;
        public const int LVM_FIRST = 0x1000;
        public const int LVM_SETIMAGELIST = LVM_FIRST + 3;
        public const int LVSIL_NORMAL = 0;
        public const uint SHGFI_LARGEICON = 0x0; //大图标 32×32

        [DllImport("shell32.dll")]
        public static extern Int32 SHGetDesktopFolder(out IntPtr ppshf);

        [DllImport("Shlwapi.Dll", CharSet = CharSet.Auto)]
        public static extern Int32 StrRetToBuf(IntPtr pstr, IntPtr pidl, StringBuilder pszBuf, int cchBuf);

        [DllImport("shell32.dll")]
        public static extern int SHGetSpecialFolderLocation(IntPtr handle, CSIDL nFolder, out IntPtr ppidl);



        [DllImport("shell32", EntryPoint = "SHGetFileInfo", ExactSpelling = false, CharSet = CharSet.Auto, SetLastError = true)]
        public static extern IntPtr SHGetFileInfo(IntPtr ppidl, FILE_ATTRIBUTE dwFileAttributes, ref SHFILEINFO sfi, int cbFileInfo, SHGFI uFlags);

        [DllImport("Shell32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr SHGetFileInfo(string Path, FILE_ATTRIBUTE fileAttributes, out SHFILEINFO sfi, int cbFileInfo, SHGFI flags);

        [DllImport("user32",
            SetLastError = true,
            CharSet = CharSet.Auto)]
        public static extern IntPtr CreatePopupMenu();

        [DllImport("user32.dll",
            ExactSpelling = true,
            CharSet = CharSet.Auto)]
        public static extern uint TrackPopupMenuEx(
            IntPtr hmenu,
            TPM flags,
            int x,
            int y,
            IntPtr hwnd,
            IntPtr lptpm);

        [DllImport("Shell32.Dll")]
        private static extern bool SHGetSpecialFolderPath(
            IntPtr hwndOwner,
            StringBuilder lpszPath,
            ShellSpecialFolders nFolder,
            bool fCreate);

        [DllImport("user32.dll", EntryPoint = "SendMessage", SetLastError = true, CallingConvention = CallingConvention.StdCall)]
        public static extern int SendMessage(IntPtr hwnd, uint Msg, int wParam, IntPtr lParam);


        /// <summary>
        /// 获得桌面 Shell
        /// </summary>
        public static IShellFolder GetDesktopFolder(out IntPtr ppshf)
        {
            SHGetDesktopFolder(out ppshf);
            Object obj = Marshal.GetObjectForIUnknown(ppshf);
            return (IShellFolder)obj;
        }

        /// <summary>
        /// 获取路径
        /// </summary>
        public static string GetPathByIShell(IShellFolder Root, IntPtr pidlSub)
        {
            IntPtr strr = Marshal.AllocCoTaskMem(MAX_PATH * 2 + 4);
            Marshal.WriteInt32(strr, 0, 0);
            StringBuilder buf = new StringBuilder(MAX_PATH);
            Root.GetDisplayNameOf(pidlSub, SHGNO.FORADDRESSBAR | SHGNO.FORPARSING, strr);
            API.StrRetToBuf(strr, pidlSub, buf, MAX_PATH);
            Marshal.FreeCoTaskMem(strr);
            return buf.ToString();
        }

        /// <summary>
        /// 获取显示名称
        /// </summary>
        public static string GetNameByIShell(IShellFolder Root, IntPtr pidlSub)
        {
            IntPtr strr = Marshal.AllocCoTaskMem(MAX_PATH * 2 + 4);
            Marshal.WriteInt32(strr, 0, 0);
            StringBuilder buf = new StringBuilder(MAX_PATH);
            Root.GetDisplayNameOf(pidlSub, SHGNO.INFOLDER, strr);
            API.StrRetToBuf(strr, pidlSub, buf, MAX_PATH);
            Marshal.FreeCoTaskMem(strr);
            return buf.ToString();
        }

        /// <summary>
        /// 根据 PIDL 获取显示名称
        /// </summary>
        public static string GetNameByPIDL(IntPtr pidl)
        {
            SHFILEINFO info = new SHFILEINFO();
            API.SHGetFileInfo(pidl, 0, ref info, Marshal.SizeOf(typeof(SHFILEINFO)),
                SHGFI.PIDL | SHGFI.DISPLAYNAME | SHGFI.TYPENAME);
            return info.szDisplayName;
        }

        /// <summary>
        /// 获取特殊文件夹的路径
        /// </summary>
        public static string GetSpecialFolderPath(IntPtr hwnd, ShellSpecialFolders nFolder)
        {
            StringBuilder sb = new StringBuilder(MAX_PATH);
            SHGetSpecialFolderPath(hwnd, sb, nFolder, false);
            return sb.ToString();
        }

        /// <summary>
        /// 根据路径获取 IShellFolder 和 PIDL
        /// </summary>
        public static IShellFolder GetShellFolder(IShellFolder desktop, string path, out IntPtr Pidl)
        {
            IShellFolder IFolder;
            uint i, j = 0;
            desktop.ParseDisplayName(IntPtr.Zero, IntPtr.Zero, path, out i, out Pidl, ref j);
            desktop.BindToObject(Pidl, IntPtr.Zero, ref Guids.IID_IShellFolder, out IFolder);
            return IFolder;
        }

        /// <summary>
        /// 根据路径获取 IShellFolder
        /// </summary>
        public static IShellFolder GetShellFolder(IShellFolder desktop, string path)
        {
            IntPtr Pidl;
            return GetShellFolder(desktop, path, out Pidl);
        }
        /// <summary>
        /// 获取大图标索引
        /// </summary>
        public static IntPtr GetLargeIconIndex(IntPtr ipIDList)
        {
            SHFILEINFO psfi = new SHFILEINFO();
            SHGFI dwflag = SHGFI.ICON | SHGFI.LARGEICON | SHGFI.PIDL | SHGFI.ICON;
            IntPtr ipIcon = SHGetFileInfo(ipIDList, 0, ref psfi, Marshal.SizeOf(psfi), dwflag);
            return psfi.hIcon;
        }
        private static IntPtr m_ipSmallSystemImageList;




    }
    #endregion
}
