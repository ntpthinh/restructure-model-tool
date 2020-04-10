#region Copyright
//      .NET Sample
//
//      Copyright (c) 2012 by Autodesk, Inc.
//
//      Permission to use, copy, modify, and distribute this software
//      for any purpose and without fee is hereby granted, provided
//      that the above copyright notice appears in all copies and
//      that both that copyright notice and the limited warranty and
//      restricted rights notice below appear in all supporting
//      documentation.
//
//      AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS.
//      AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
//      MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE.  AUTODESK, INC.
//      DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
//      UNINTERRUPTED OR ERROR FREE.
//
//      Use, duplication, or disclosure by the U.S. Government is subject to
//      restrictions set forth in FAR 52.227-19 (Commercial Computer
//      Software - Restricted Rights) and DFAR 252.227-7013(c)(1)(ii)
//      (Rights in Technical Data and Computer Software), as applicable.
//
#endregion

#region imports
using System;
using System.Windows;
using System.Runtime.InteropServices;
using System.Diagnostics;

using UiViewModels.Actions;
using Autodesk.Max;
using RestructureModelTool.View;
using System.Drawing;
using System.Runtime.InteropServices.ComTypes;
using Microsoft.VisualStudio.OLE.Interop;
using STATSTG = Microsoft.VisualStudio.OLE.Interop.STATSTG;
using IStream = Microsoft.VisualStudio.OLE.Interop.IStream;
#endregion

namespace ADNMenuSample
{
    public class MenuItemStrings
    {
        // just convienence for globals strings. Normally strings would probably be loaded from resources
        public static string actionWindow = "Restructure Model";
        public static string actionCategory = "Restructure Model";
    }
   
    public class RestructureModelWindow : CuiActionCommandAdapter
    {
        public bool isopen = false;
        MainWindow dialog = null;
        public override void Execute(object parameter)
        {
            if (dialog == null)
            {
                dialog = new MainWindow();
            }
            try
            {
                // Setup a dialog that is connected to the parent window. In this case
                // it is 3ds Max, so we use some of the utility APIs from ManagedServices
                // to make that connection.
                if (dialog.Visibility == Visibility.Hidden)
                {
                    isopen = false;
                }
                if (!isopen)
                {
                    isopen = true;
                    /*Bitmap a = GetMaxPreviewBitmapFromFile(@"E:\Model\Animal\754638.584807811659e\max\PUG2011.max");
                    a.Save(@"D:\test.jpg");*/
                    dialog.Closed += new EventHandler(SearchSimilar_Closed);
                    System.Windows.Interop.WindowInteropHelper windowHandle = new System.Windows.Interop.WindowInteropHelper(dialog);
                    windowHandle.Owner = ManagedServices.AppSDK.GetMaxHWND();
                    ManagedServices.AppSDK.ConfigureWindowForMax(dialog);
                    dialog.Show();
                }
                else
                {
                    isopen = false;
                    dialog.Hide();
                }

            }
            catch (System.Exception ex)
            {
                Debug.Print("Exception occurred: " + ex.Message);
            }

        }
        void SearchSimilar_Closed(object sender, EventArgs e)
        {
            isopen = false;
        }
        public override string InternalActionText
        {
            get { return MenuItemStrings.actionWindow; }
        }

        public override string InternalCategory
        {
            get { return MenuItemStrings.actionCategory; }
        }

        public override string ActionText
        {
            get { return InternalActionText; }
        }

        public override string Category
        {
            get { return InternalCategory; }
        }
        [Flags]
        public enum STGM : int
        {
            DIRECT = 0x00000000,
            TRANSACTED = 0x00010000,
            SIMPLE = 0x08000000,
            READ = 0x00000000,
            WRITE = 0x00000001,
            READWRITE = 0x00000002,
            SHARE_DENY_NONE = 0x00000040,
            SHARE_DENY_READ = 0x00000030,
            SHARE_DENY_WRITE = 0x00000020,
            SHARE_EXCLUSIVE = 0x00000010,
            PRIORITY = 0x00040000,
            DELETEONRELEASE = 0x04000000,
            NOSCRATCH = 0x00100000,
            CREATE = 0x00001000,
            CONVERT = 0x00020000,
            FAILIFTHERE = 0x00000000,
            NOSNAPSHOT = 0x00200000,
            DIRECT_SWMR = 0x00400000,
        }

        enum ulKind : uint
        {
            PRSPEC_LPWSTR = 0,
            PRSPEC_PROPID = 1
        }

        enum SumInfoProperty : uint
        {
            PIDSI_TITLE = 0x00000002,
            PIDSI_SUBJECT = 0x00000003,
            PIDSI_AUTHOR = 0x00000004,
            PIDSI_KEYWORDS = 0x00000005,
            PIDSI_COMMENTS = 0x00000006,
            PIDSI_TEMPLATE = 0x00000007,
            PIDSI_LASTAUTHOR = 0x00000008,
            PIDSI_REVNUMBER = 0x00000009,
            PIDSI_EDITTIME = 0x0000000A,
            PIDSI_LASTPRINTED = 0x0000000B,
            PIDSI_CREATE_DTM = 0x0000000C,
            PIDSI_LASTSAVE_DTM = 0x0000000D,
            PIDSI_PAGECOUNT = 0x0000000E,
            PIDSI_WORDCOUNT = 0x0000000F,
            PIDSI_CHARCOUNT = 0x00000010,
            PIDSI_THUMBNAIL = 0x00000011,
            PIDSI_APPNAME = 0x00000012,
            PIDSI_SECURITY = 0x00000013
        }

        public enum VARTYPE : short
        {
            VT_BSTR = 8,
            VT_FILETIME = 0x40,
            VT_LPSTR = 30,
            VT_CF = 71
        }

        [StructLayout(LayoutKind.Explicit)]
        public struct PROPVARIANTunion
        {
            [FieldOffset(0)]
            public sbyte cVal;
            [FieldOffset(0)]
            public byte bVal;
            [FieldOffset(0)]
            public short iVal;
            [FieldOffset(0)]
            public ushort uiVal;
            [FieldOffset(0)]
            public int lVal;
            [FieldOffset(0)]
            public uint ulVal;
            [FieldOffset(0)]
            public int intVal;
            [FieldOffset(0)]
            public uint uintVal;
            [FieldOffset(0)]
            public long hVal;
            [FieldOffset(0)]
            public ulong uhVal;
            [FieldOffset(0)]
            public float fltVal;
            [FieldOffset(0)]
            public double dblVal;
            [FieldOffset(0)]
            public short boolVal;
            [FieldOffset(0)]
            public int scode;
            [FieldOffset(0)]
            public long cyVal;
            [FieldOffset(0)]
            public double date;
            [FieldOffset(0)]
            public long filetime;
            [FieldOffset(0)]
            public IntPtr bstrVal;
            [FieldOffset(0)]
            public IntPtr pszVal;
            [FieldOffset(0)]
            public IntPtr pwszVal;
            [FieldOffset(0)]
            public IntPtr punkVal;
            [FieldOffset(0)]
            public IntPtr pdispVal;
        }

        struct PACKEDMETA
        {
            public ushort mm, xExt, yExt, reserved;
        }

        [DllImport("ole32.dll")]
        static extern int StgOpenStorage(
            [MarshalAs(UnmanagedType.LPWStr)]string pwcsName, IStorage pstgPriority,
            int grfMode, IntPtr snbExclude, uint reserved, out IStorage ppstgOpen);

        [DllImport("ole32.dll")]
        static extern int StgCreatePropSetStg(IStorage pStorage, uint reserved,
           out IPropertySetStorage ppPropSetStg);

        [DllImport("ole32.dll")]
        private extern static int PropVariantClear(ref PROPVARIANT pvar);


        [ComImport]
        [Guid("00000138-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IPropertyStorage
        {
            [PreserveSig]
            int ReadMultiple(uint cpspec,
                [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] [In] PropertySpec[] rgpspec,
                [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] [Out] PropertyVariant[] rgpropvar);

            [PreserveSig]
            void WriteMultiple(uint cpspec,
                [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] [In] PropertySpec[] rgpspec,
                [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] [In] PropertyVariant[] rgpropvar,
                uint propidNameFirst);

            [PreserveSig]
            uint DeleteMultiple(uint cpspec,
                [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] [In] PropertySpec[] rgpspec);
            [PreserveSig]
            uint ReadPropertyNames(uint cpropid,
                [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] [In] uint[] rgpropid,
                [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr, SizeParamIndex = 0)] [Out] string[] rglpwstrName);
            [PreserveSig]
            uint NotDeclared1();
            [PreserveSig]
            uint NotDeclared2();
            [PreserveSig]
            uint Commit(uint grfCommitFlags);
            [PreserveSig]
            uint NotDeclared3();
            [PreserveSig]
            uint Enum(out IEnumSTATPROPSTG ppenum);
        }

        [ComImport]
        [Guid("0000013A-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IPropertySetStorage
        {
            [PreserveSig]
            uint Create(ref Guid rfmtid, ref Guid pclsid, uint grfFlags, STGM grfMode, out IPropertyStorage ppprstg);
            [PreserveSig]
            uint Open(ref Guid rfmtid, STGM grfMode, out IPropertyStorage ppprstg);
            [PreserveSig]
            uint NotDeclared3();
            [PreserveSig]
            uint Enum(out IEnumSTATPROPSETSTG ppenum);
        }

        public enum PropertySpecKind
        {
            Lpwstr,
            PropId
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct PropertySpec
        {
            public PropertySpecKind kind;
            public PropertySpecData data;
        }

        [StructLayout(LayoutKind.Explicit)]
        public struct PropertySpecData
        {
            [FieldOffset(0)]
            public uint propertyId;
            [FieldOffset(0)]
            public IntPtr name;
        }

        public struct PropertyVariant
        {
            public VARTYPE vt;
            public ushort wReserved1;
            public ushort wReserved2;
            public ushort wReserved3;
            public PROPVARIANTunion unionmember;
        }

        static public Bitmap GetMaxPreviewBitmapFromFile(string path)
        {
            var FMTID_SummaryInformation = new Guid("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}");

            Bitmap bitmap = null;

            IStorage Is;
            if (StgOpenStorage(path, null, (int)(STGM.SHARE_EXCLUSIVE | STGM.READWRITE), IntPtr.Zero, 0, out Is) == 0 && Is != null)
            {
                IPropertySetStorage pss;
                if (StgCreatePropSetStg(Is, 0, out pss) == 0)
                {
                    IPropertyStorage ps;
                    pss.Open(ref FMTID_SummaryInformation, (STGM.SHARE_EXCLUSIVE | STGM.READ), out ps);
                    if (ps != null)
                    {
                        var propSpec = new PropertySpec[1];
                        var propVariant = new PropertyVariant[1];

                        propSpec[0].kind = PropertySpecKind.PropId;
                        propSpec[0].data.propertyId = (uint)SumInfoProperty.PIDSI_THUMBNAIL;

                        System.UInt32 n = 1;
                        ps.ReadMultiple(n, propSpec, propVariant);

                        var clipData =
                            (CLIPDATA)Marshal.PtrToStructure(propVariant[0].unionmember.pszVal, typeof(CLIPDATA));

                        var pb = clipData.pClipData;
                        pb += sizeof(uint);

                        var packedMeta = (PACKEDMETA)Marshal.PtrToStructure(pb, typeof(PACKEDMETA));
                        pb += Marshal.SizeOf(packedMeta);

                        var magicNumber = 3 * 29;
                        pb += magicNumber;

                        var pformat = System.Drawing.Imaging.PixelFormat.Format24bppRgb;
                        int bitsPerPixel = ((int)pformat & 0xff00) >> 8;
                        int bytesPerPixel = (bitsPerPixel + 7) / 8;
                        int stride = 4 * ((packedMeta.xExt * bytesPerPixel + 3) / 4);

                        unsafe
                        {
                            byte* ptr = (byte*)pb;
                            for (int y = 0; y < packedMeta.yExt; y++)
                                for (int x = 0; x < packedMeta.xExt; x++)
                                {
                                    var i = (x * 3) + y * stride;

                                    var r = ptr[i];
                                    var g = ptr[i + 1];
                                    var b = ptr[i + 2];

                                    ptr[i] = b;
                                    ptr[i + 1] = r;
                                    ptr[i + 2] = g;

                                }

                            bitmap = new Bitmap(packedMeta.xExt, packedMeta.yExt, stride, pformat, (IntPtr)ptr);

                            bitmap.RotateFlip(RotateFlipType.Rotate180FlipX);
                        }

                        //PropVariantClear(ref propVariant[0]);
                        Marshal.FinalReleaseComObject(ps);
                        ps = null;
                    }
                    else
                    {
                        Console.WriteLine("Could not open property storage");
                    }

                    Marshal.FinalReleaseComObject(pss);
                    pss = null;
                }
                else
                {
                    Console.WriteLine("Could not create property set storage");
                }

                Marshal.FinalReleaseComObject(Is);
                Is = null;
            }
            else
            {
                Console.WriteLine("File does not contain a structured storage");
            }

            GC.Collect();

            return bitmap;
        }
    }
}
