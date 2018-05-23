using System;
using System.Runtime.InteropServices;
using System.Text;

namespace ClassLibrary1.FIleShell
{
        #region 右键菜单
        [ComImport(), InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("000214e4-0000-0000-c000-000000000046")]
        public interface IContextMenu
        {
            [PreserveSig()]
            Int32 QueryContextMenu(
                IntPtr hmenu,
                uint iMenu,
                uint idCmdFirst,
                uint idCmdLast,
                CMF uFlags);

            [PreserveSig()]
            Int32 InvokeCommand(
                ref CMINVOKECOMMANDINFOEX info);

            [PreserveSig()]
            void GetCommandString(
                int idcmd,
                GetCommandStringInformations uflags,
                int reserved,
                StringBuilder commandstring,
                int cch);
        }
        #endregion

        #region IEnumIDList
        [ComImport(),
       Guid("000214F2-0000-0000-C000-000000000046"),
       InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IEnumIDList
        {
            [PreserveSig()]
            uint Next(
                uint celt,
                out IntPtr rgelt,
                out int pceltFetched);

            void Skip(
                uint celt);

            void Reset();

            IEnumIDList Clone();
        }
        #endregion

        #region ShellItem
        public class ShellItem
        {
            public ShellItem(IntPtr PIDL, IShellFolder ShellFolder, ShellItem ParentItem)
            {
                m_PIDL = PIDL;
                m_ShellFolder = ShellFolder;
                m_ParentItem = ParentItem;
            }
            public ShellItem(IntPtr PIDL, IShellFolder ShellFolder)
            {
                m_PIDL = PIDL;
                m_ShellFolder = ShellFolder;
            }

            private IntPtr m_PIDL;

            public IntPtr PIDL
            {
                get { return m_PIDL; }
                set { m_PIDL = value; }
            }


            private IShellFolder m_ShellFolder;

            public IShellFolder ShellFolder
            {
                get { return m_ShellFolder; }
                set { m_ShellFolder = value; }
            }
            private ShellItem m_ParentItem;

            public ShellItem ParentItem
            {
                get { return m_ParentItem; }
                set { m_ParentItem = value; }
            }

            /// <summary>
            /// 绝对 PIDL
            /// </summary>
            public PIDL PIDLFull
            {
                get
                {
                    PIDL pidlFull = new PIDL(PIDL, true);
                    ShellItem current = ParentItem;
                    while (current != null)
                    {
                        pidlFull.Insert(current.PIDL);
                        current = current.ParentItem;
                    }
                    return pidlFull;
                }
            }
        }
        #endregion


        #region IShellFolder
        [ComImport]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("000214E6-0000-0000-C000-000000000046")]
        public interface IShellFolder
        {
            void ParseDisplayName(
                IntPtr hwnd,
                IntPtr pbc,
                [MarshalAs(UnmanagedType.LPWStr)] string pszDisplayName,
                out uint pchEaten,
                out IntPtr ppidl,
                ref uint pdwAttributes);

            [PreserveSig]
            int EnumObjects(IntPtr hWnd, SHCONTF flags, out IntPtr enumIDList);

            void BindToObject(
                IntPtr pidl,
                IntPtr pbc,
                [In()] ref Guid riid,
                out IShellFolder ppv);

            void BindToStorage(
                IntPtr pidl,
                IntPtr pbc,
                [In()] ref Guid riid,
                [MarshalAs(UnmanagedType.Interface)] out object ppv);

            [PreserveSig()]
            uint CompareIDs(
                int lParam,
                IntPtr pidl1,
                IntPtr pidl2);

            void CreateViewObject(
                IntPtr hwndOwner,
                [In()] ref Guid riid,
                [MarshalAs(UnmanagedType.Interface)] out object ppv);

            void GetAttributesOf(
                uint cidl,
                [In(), MarshalAs(UnmanagedType.LPArray)] IntPtr[] apidl,
               ref SFGAO rgfInOut);

            IntPtr GetUIObjectOf(
                IntPtr hwndOwner,
                uint cidl,
                [MarshalAs(UnmanagedType.LPArray)] IntPtr[] apidl,
                [In()] ref Guid riid,
                out IntPtr rgfReserved);

            void GetDisplayNameOf(
                IntPtr pidl,
                SHGNO uFlags,
                IntPtr lpName);

            IntPtr SetNameOf(
                IntPtr hwnd,
                IntPtr pidl,
                [MarshalAs(UnmanagedType.LPWStr)] string pszName,
               SHGNO uFlags);
        }
        #endregion



        public class PIDL
        {
            private IntPtr pidl = IntPtr.Zero;

            public PIDL(IntPtr pidl, bool clone)
            {
                if (clone)
                    this.pidl = ILClone(pidl);
                else
                    this.pidl = pidl;
            }

            public PIDL(PIDL pidl, bool clone)
            {
                if (clone)
                    this.pidl = ILClone(pidl.Ptr);
                else
                    this.pidl = pidl.Ptr;
            }


            public IntPtr Ptr
            {
                get { return pidl; }
            }

            public void Insert(IntPtr insertPidl)
            {
                IntPtr newPidl = ILCombine(insertPidl, pidl);

                Marshal.FreeCoTaskMem(pidl);
                pidl = newPidl;
            }

            public void Free()
            {
                if (pidl != IntPtr.Zero)
                {
                    Marshal.FreeCoTaskMem(pidl);
                    pidl = IntPtr.Zero;
                }
            }


            private static int ItemIDSize(IntPtr pidl)
            {
                if (!pidl.Equals(IntPtr.Zero))
                {
                    byte[] buffer = new byte[2];
                    Marshal.Copy(pidl, buffer, 0, 2);
                    return buffer[1] * 256 + buffer[0];
                }
                else
                    return 0;
            }

            private static int ItemIDListSize(IntPtr pidl)
            {
                if (pidl.Equals(IntPtr.Zero))
                    return 0;
                else
                {
                    int size = ItemIDSize(pidl);
                    int nextSize = Marshal.ReadByte(pidl, size) + (Marshal.ReadByte(pidl, size + 1) * 256);
                    while (nextSize > 0)
                    {
                        size += nextSize;
                        nextSize = Marshal.ReadByte(pidl, size) + (Marshal.ReadByte(pidl, size + 1) * 256);
                    }

                    return size;
                }
            }

            public static IntPtr ILClone(IntPtr pidl)
            {
                int size = ItemIDListSize(pidl);

                byte[] bytes = new byte[size + 2];
                Marshal.Copy(pidl, bytes, 0, size);

                IntPtr newPidl = Marshal.AllocCoTaskMem(size + 2);
                Marshal.Copy(bytes, 0, newPidl, size + 2);

                return newPidl;
            }

            public static IntPtr ILCombine(IntPtr pidl1, IntPtr pidl2)
            {
                int size1 = ItemIDListSize(pidl1);
                int size2 = ItemIDListSize(pidl2);

                IntPtr newPidl = Marshal.AllocCoTaskMem(size1 + size2 + 2);
                byte[] bytes = new byte[size1 + size2 + 2];

                Marshal.Copy(pidl1, bytes, 0, size1);
                Marshal.Copy(pidl2, bytes, size1, size2);

                Marshal.Copy(bytes, 0, newPidl, bytes.Length);

                return newPidl;
            }
        }

    }


