using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;

namespace ClassLibrary1.FIleShell
{
   
    class WinShellHelper
    {
        public IntPtr DeskTopPtr;
        public IShellFolder DeskTop;//桌面
        public ShellItem ShellItem;
        public IntPtr EnumPtr = IntPtr.Zero;
        public IEnumIDList Enum;
        public  List<IShellFolder> ShellFolder = new List<IShellFolder>();
        public List<ShellItem> AllShellItem = new List<ShellItem>();//当前文件夹下的文件夹集合
        


        /// <summary>
        /// 获取外壳空间的根路径
        /// </summary>
        /// <returns></returns>
        public IShellFolder GetDeskTopFolder()
        {
            DeskTop = API.GetDesktopFolder(out DeskTopPtr);
            return DeskTop;
        }

        /// <summary>
        /// 获取大图标
        /// </summary>
        /// <param name="pidlSub">根文件夹下的索引</param>
        /// <param name="sItem">根文件夹</param>
        /// <returns></returns>
        public Icon GetLargeIcon(IntPtr pidlSub, ShellItem sItem)
        {
            ShellItem shellItem = new ShellItem(pidlSub, null, sItem);
            IntPtr imgIntPtr = API.GetLargeIconIndex(shellItem.PIDLFull.Ptr);
            if (imgIntPtr == IntPtr.Zero) return null;
            Icon systemIcon = Icon.FromHandle(imgIntPtr);
            return systemIcon;
        }

        /// <summary>
        /// 获取文件名称
        /// </summary>
        /// <param name="root"></param>
        /// <param name="pidlSub"></param>
        /// <returns></returns>
        public string GetFileName(IShellFolder root, IntPtr pidlSub)
        {
            string name;
            name = API.GetNameByIShell(root, pidlSub);
            return name;
        }

        /// <summary>
        /// 获取文件路径
        /// </summary>
        /// <param name="root"></param>
        /// <param name="pidlSub"></param>
        /// <returns></returns>
        public string GetFilePath(IShellFolder root, IntPtr pidlSub)
        {
            string path;
            path = API.GetPathByIShell(root, pidlSub);
            return path;
        }

       
        /// <summary>
        /// 获取当前父节点
        /// </summary>
        /// <param name="sItem"></param>
        /// <returns></returns>
        public IShellFolder GetShellFolder(ShellItem sItem)
        {
            IShellFolder root = sItem.ShellFolder;
            return root;
        }

        /// <summary>
        /// 获取所有文件夹
        /// </summary>
        /// <param name="root">根目录</param>
        /// <param name="handle">当前窗口句柄</param>
        /// <param name="sItem">当前项</param>
        public List<ShellItem> GetAllFolders(IShellFolder root, IntPtr handle, ShellItem sItem)
        {
            var list = new List<ShellItem>();
            list.AddRange(GetFolders(root , handle , sItem));
            list.AddRange(GetNonFolders(root, handle, sItem));
            return list;
        }

        /// <summary>
        /// 获取文件夹
        /// </summary>
        /// <param name="root">根目录</param>
        /// <param name="handle">当前窗口句柄</param>
        /// <param name="sItem">当前项</param>
        public List<ShellItem> GetFolders(IShellFolder root, IntPtr handle, ShellItem sItem)
        {
            List<ShellItem> allFolderShellItem = new List<ShellItem>();//文件夹的集合
            if (root.EnumObjects(handle, SHCONTF.FOLDERS | SHCONTF.INCLUDEHIDDEN, out EnumPtr) ==
               API.S_OK)
            {
                Enum = (IEnumIDList)Marshal.GetObjectForIUnknown(EnumPtr);
                int celtFetched;
                IntPtr pidlSub;
                while (Enum.Next(1, out pidlSub, out celtFetched) == 0 && celtFetched == API.S_FALSE)
                {
                    var shellItem = new ShellItem(pidlSub, null, sItem);
                    allFolderShellItem.Add(shellItem);
                }
            }
          
            return allFolderShellItem;
        }

        /// <summary>
        /// 获取非文件夹的文件
        /// </summary>
        /// <param name="root">根目录</param>
        /// <param name="handle">当前窗口句柄</param>
        /// <param name="sItem">当前项</param>
        /// <returns></returns>
        public List<ShellItem> GetNonFolders(IShellFolder root, IntPtr handle, ShellItem sItem)
        {
            List<ShellItem> allFileShellItem = new List<ShellItem>();//文件的集合
            if (root.EnumObjects(handle, SHCONTF.NONFOLDERS | SHCONTF.INCLUDEHIDDEN, out EnumPtr) ==
              API.S_OK)
            {
                Enum = (IEnumIDList)Marshal.GetObjectForIUnknown(EnumPtr);
                IntPtr pidlSub;
                int celtFetched;
                while (Enum.Next(1, out pidlSub, out celtFetched) == 0 && celtFetched == API.S_FALSE)
                {
                    var shellItem = new ShellItem(pidlSub, null, sItem);
                    allFileShellItem.Add(shellItem);
                }
            }
            return allFileShellItem;
        }
    }
}
