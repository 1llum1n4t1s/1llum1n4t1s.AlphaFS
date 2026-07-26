/*  Copyright (C) 2008-2018 Peter Palotas, Jeffrey Jangli, Alexandr Normuradov
 *  
 *  Permission is hereby granted, free of charge, to any person obtaining a copy 
 *  of this software and associated documentation files (the "Software"), to deal 
 *  in the Software without restriction, including without limitation the rights 
 *  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell 
 *  copies of the Software, and to permit persons to whom the Software is 
 *  furnished to do so, subject to the following conditions:
 *  
 *  The above copyright notice and this permission notice shall be included in 
 *  all copies or substantial portions of the Software.
 *  
 *  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR 
 *  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
 *  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE 
 *  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER 
 *  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, 
 *  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN 
 *  THE SOFTWARE. 
 */

using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using System.Security;

namespace Alphaleonis.Win32
{
   /// <summary>アセンブリが実行されているオペレーティングシステムに関する情報へのアクセスを提供する静的クラス。</summary>
   public static class OperatingSystem
   {
      /// <summary>名前付き Windows バージョンを記述するフラグのセット。</summary>
      /// <remarks>列挙の値は順序付けられています。より後にリリースされたオペレーティングシステムバージョンはより大きな数値を持つため、名前付きバージョン間の比較は意味があります。</remarks>
      [SuppressMessage("Microsoft.Naming", "CA1709:IdentifiersShouldBeCasedCorrectly", MessageId = "Os")]
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Os")]
      public enum EnumOsName
      {
         /// <summary>Windows 2000 より前の Windows バージョン。</summary>
         Earlier = -1,

         /// <summary>Windows 2000 (Server または Professional)。</summary>
         Windows2000 = 0,

         /// <summary>Windows XP.</summary>
         WindowsXP = 1,

         /// <summary>Windows Server 2003.</summary>
         WindowsServer2003 = 2,

         /// <summary>Windows Vista.</summary>
         WindowsVista = 3,

         /// <summary>Windows Server 2008.</summary>
         WindowsServer2008 = 4,

         /// <summary>Windows 7.</summary>
         Windows7 = 5,

         /// <summary>Windows Server 2008 R2.</summary>
         WindowsServer2008R2 = 6,

         /// <summary>Windows 8.</summary>
         Windows8 = 7,

         /// <summary>Windows Server 2012.</summary>
         WindowsServer2012 = 8,

         /// <summary>Windows 8.1.</summary>
         Windows81 = 9,

         /// <summary>Windows Server 2012 R2</summary>
         WindowsServer2012R2 = 10,

         /// <summary>Windows 10</summary>
         Windows10 = 11,

         /// <summary>Windows Server 2016</summary>
         WindowsServer2016 = 12,

         /// <summary>Windows Server 2019 (build 17763+)。</summary>
         WindowsServer2019 = 13,

         /// <summary>Windows 11 (build 22000+)。</summary>
         Windows11 = 14,

         /// <summary>Windows Server 2022 (build 20348+)。</summary>
         WindowsServer2022 = 15,

         /// <summary>Windows Server 2025 (build 26100+)。</summary>
         WindowsServer2025 = 16,

         /// <summary>現在インストールされているものより後のバージョンの Windows。</summary>
         /// <remarks>既知の OS には専用の値を割り当てるため、この値は「このライブラリが知らない将来の OS」だけを表す。</remarks>
         Later = 65535
      }


      /// <summary>オペレーティングシステムが対象とし実行されている現在のプロセッサアーキテクチャを示すフラグのセット。</summary>
      [SuppressMessage("Microsoft.Naming", "CA1709:IdentifiersShouldBeCasedCorrectly", MessageId = "Pa")]
      [SuppressMessage("Microsoft.Design", "CA1028:EnumStorageShouldBeInt32")]      
      public enum EnumProcessorArchitecture 
      {
         /// <summary>PROCESSOR_ARCHITECTURE_INTEL
         /// <para>システムは 32 ビット版の Windows を実行しています。</para>
         /// </summary>
         X86 = 0,

         /// <summary>PROCESSOR_ARCHITECTURE_IA64
         /// <para>システムは Itanium プロセッサ上で実行されています。</para>
         /// </summary>
         [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Ia")]
         [SuppressMessage("Microsoft.Naming", "CA1709:IdentifiersShouldBeCasedCorrectly", MessageId = "Ia")]
         IA64 = 6,

         /// <summary>PROCESSOR_ARCHITECTURE_AMD64
         /// <para>システムは 64 ビット版の Windows を実行しています。</para>
         /// </summary>
         X64 = 9,

         /// <summary>PROCESSOR_ARCHITECTURE_UNKNOWN
         /// <para>不明なアーキテクチャ。</para>
         /// </summary>
         Unknown = 65535
      }


      
      
      #region Properties

      private static bool _isServer;
      /// <summary>オペレーティングシステムがサーバーオペレーティングシステムであるかどうかを示す値を取得します。</summary>
      /// <value>現在のオペレーティングシステムがサーバーオペレーティングシステムの場合は <c>true</c>、それ以外の場合は <c>false</c>。</value>
      public static bool IsServer
      {
         get
         {
            if (null == _servicePackVersion)
            {
               UpdateData();
            }

            return _isServer;
         }
      }


      private static bool? _isWow64Process;
      /// <summary>現在のプロセスが WOW64 で実行されているかどうかを示す値を取得します。</summary>
      /// <value>現在のプロセスが WOW64 で実行されている場合は <c>true</c>、それ以外の場合は <c>false</c>。</value>
      public static bool IsWow64Process
      {
         get
         {
            if (null == _isWow64Process)
            {
               // Process インスタンスを変数に保持せずに .Handle だけ取り出すと、ネイティブ呼び出しの前に
               // Process が回収・ファイナライズされてハンドルが閉じられ、ERROR_INVALID_HANDLE になり得る。
               // using で呼び出しが終わるまで確実に生存させる。
               bool value;

               using (var process = Process.GetCurrentProcess())
               {
                  if (!NativeMethods.IsWow64Process(process.Handle, out value))
                  {
                     throw new Win32Exception(Marshal.GetLastWin32Error());
                  }
               }

               // プロセスが WOW64 で実行されている場合に TRUE に設定される値へのポインター。
               // プロセスが 32 ビット Windows で実行されている場合、値は FALSE に設定される。
               // プロセスが 64 ビット Windows で実行されている 64 ビットアプリケーションの場合も、値は FALSE に設定される。

               _isWow64Process = value;
            }

            return (bool) _isWow64Process;
         }
      }


      private static Version _osVersion;
      /// <summary>オペレーティングシステムの数値バージョンを取得します。</summary>
      /// <value>オペレーティングシステムの数値バージョン。</value>
      public static Version OSVersion
      {
         get
         {
            if (null == _osVersion)
            {
               UpdateData();
            }

            return _osVersion;
         }
      }


      private static EnumOsName _enumOsName = EnumOsName.Later;
      /// <summary>オペレーティングシステムの名前付きバージョンを取得します。</summary>
      /// <value>オペレーティングシステムの名前付きバージョン。</value>
      public static EnumOsName VersionName
      {
         get
         {
            if (null == _servicePackVersion)
            {
               UpdateData();
            }

            return _enumOsName;
         }
      }


      private static EnumProcessorArchitecture _processorArchitecture;
      /// <summary>オペレーティングシステムが対象とするプロセッサアーキテクチャを取得します。</summary>
      /// <value>オペレーティングシステムが対象とするプロセッサアーキテクチャ。</value>
      /// <remarks>WOW64 で実行されている場合、32 ビットプロセッサが返されます。これに該当するかどうかを判定するには <see cref="IsWow64Process"/> を使用してください。</remarks>
      public static EnumProcessorArchitecture ProcessorArchitecture
      {
         get
         {
            if (null == _servicePackVersion)
            {
               UpdateData();
            }

            return _processorArchitecture;
         }
      }


      private static Version _servicePackVersion;
      /// <summary>オペレーティングシステムに現在インストールされているサービスパックのバージョンを取得します。</summary>
      /// <value>オペレーティングシステムに現在インストールされているサービスパックのバージョン。</value>
      /// <remarks><see cref="System.Version.Major"/> および <see cref="System.Version.Minor"/> フィールドのみが使用されます。</remarks>
      public static Version ServicePackVersion
      {
         get
         {
            if (null == _servicePackVersion)
            {
               UpdateData();
            }

            return _servicePackVersion;
         }
      }

      #endregion // Properties


      #region Methods

      /// <summary>オペレーティングシステムが指定されたバージョン以降であるかどうかを判定します。</summary>
      /// <returns>オペレーティングシステムが指定された <paramref name="version"/> 以降の場合は <c>true</c>、それ以外の場合は <c>false</c>。</returns>
      /// <param name="version">true を返す最小バージョン。</param>
      public static bool IsAtLeast(EnumOsName version)
      {
         return VersionName >= version;
      }

      
      /// <summary>オペレーティングシステムが指定されたバージョン以降であるかどうかを判定し、最小バージョンにインストールされている必要がある最小サービスパックの指定を可能にします。</summary>
      /// <returns>オペレーティングシステムが指定されたサービスパックを持つ指定された <paramref name="version"/> に一致する場合、またはオペレーティングシステムがそれ以降のバージョンの場合は <c>true</c>、それ以外の場合は <c>false</c>。</returns>
      /// <param name="version">必要な最小バージョン。</param>
      /// <param name="servicePackVersion">true を返すために必要な最小バージョンにインストールされている必要があるサービスパックのメジャーバージョン。サービスパックが不要であることを示すには 0 を指定できます。</param>
      public static bool IsAtLeast(EnumOsName version, int servicePackVersion)
      {
         return VersionName > version || (VersionName >= version && ServicePackVersion.Major >= servicePackVersion);
      }
      
      #endregion // Methods


      #region Private

      [SuppressMessage("Microsoft.Naming", "CA2204:Literals should be spelled correctly", MessageId = "RtlGetVersion")]
      private static void UpdateData()
      {
         var verInfo = new NativeMethods.RTL_OSVERSIONINFOEXW();

         // 次のエラーを防ぐために必要: System.Runtime.InteropServices.COMException:
         // システムコールに渡されたデータ領域が小さすぎます。(HRESULT からの例外: 0x8007007A)
         verInfo.dwOSVersionInfoSize = Marshal.SizeOf(verInfo);

         var sysInfo = new NativeMethods.SYSTEM_INFO();
         NativeMethods.GetNativeSystemInfo(ref sysInfo);


         // RtlGetVersion は STATUS_SUCCESS (0) を返す。
         var success = !NativeMethods.RtlGetVersion(ref verInfo);
         
         var lastError = Marshal.GetLastWin32Error();
         if (!success)
         {
            throw new Win32Exception(lastError, "Function RtlGetVersion() failed to retrieve the operating system information.");
         }


         _osVersion = new Version(verInfo.dwMajorVersion, verInfo.dwMinorVersion, verInfo.dwBuildNumber);

         _processorArchitecture = sysInfo.wProcessorArchitecture;
         _servicePackVersion = new Version(verInfo.wServicePackMajor, verInfo.wServicePackMinor);
         _isServer = verInfo.wProductType == NativeMethods.VER_NT_DOMAIN_CONTROLLER || verInfo.wProductType == NativeMethods.VER_NT_SERVER;


         // RtlGetVersion: https://msdn.microsoft.com/en-us/library/windows/hardware/ff561910%28v=vs.85%29.aspx
         // Operating System Version: https://msdn.microsoft.com/en-us/library/windows/desktop/ms724832(v=vs.85).aspx

         // The following table summarizes the most recent operating system version numbers.
         //    Operating system	            Version number    Other
         // ================================================================================
         //    Windows 10                    10.0              OSVERSIONINFOEX.wProductType == VER_NT_WORKSTATION
         //    Windows Server 2016           10.0              OSVERSIONINFOEX.wProductType != VER_NT_WORKSTATION
         //    Windows 8.1                   6.3               OSVERSIONINFOEX.wProductType == VER_NT_WORKSTATION
         //    Windows Server 2012 R2        6.3               OSVERSIONINFOEX.wProductType != VER_NT_WORKSTATION
         //    Windows 8	                  6.2               OSVERSIONINFOEX.wProductType == VER_NT_WORKSTATION
         //    Windows Server 2012	         6.2               OSVERSIONINFOEX.wProductType != VER_NT_WORKSTATION
         //    Windows 7	                  6.1               OSVERSIONINFOEX.wProductType == VER_NT_WORKSTATION
         //    Windows Server 2008 R2	      6.1               OSVERSIONINFOEX.wProductType != VER_NT_WORKSTATION
         //    Windows Server 2008	         6.0               OSVERSIONINFOEX.wProductType != VER_NT_WORKSTATION  
         //    Windows Vista	               6.0               OSVERSIONINFOEX.wProductType == VER_NT_WORKSTATION
         //    Windows Server 2003 R2	      5.2               GetSystemMetrics(SM_SERVERR2) != 0
         //    Windows Server 2003           5.2               GetSystemMetrics(SM_SERVERR2) == 0
         //    Windows XP 64-Bit Edition     5.2               (OSVERSIONINFOEX.wProductType == VER_NT_WORKSTATION) && (sysInfo.PaName == PaName.X64)
         //    Windows XP	                  5.1               Not applicable
         //    Windows 2000	               5.0               Not applicable


         // 2026-05-17: Windows 11 / Server 2022 以降も dwMajorVersion=10 を返す（Build 番号で識別する必要がある）。
         // Later は「このライブラリが知らない将来の OS」専用。既知の OS は必ず専用の EnumOsName に解決する。
         if (verInfo.dwMajorVersion > 10)
         {
            _enumOsName = EnumOsName.Later;
         }

         else
         {
            switch (verInfo.dwMajorVersion)
            {
               #region Version 10

               case 10:

                  // major=10 の系列は Build 番号でしか世代を区別できない。
                  if (verInfo.wProductType == NativeMethods.VER_NT_WORKSTATION)
                  {
                     // Build 22000+ = Windows 11 (Server 系は Server 区分なのでここには来ない)
                     _enumOsName = verInfo.dwBuildNumber >= 22000 ? EnumOsName.Windows11 : EnumOsName.Windows10;
                  }

                  else
                  {
                     // 26100+ = Server 2025, 20348+ = Server 2022, 17763+ = Server 2019, それ未満 = Server 2016
                     _enumOsName =
                        verInfo.dwBuildNumber >= 26100 ? EnumOsName.WindowsServer2025 :
                        verInfo.dwBuildNumber >= 20348 ? EnumOsName.WindowsServer2022 :
                        verInfo.dwBuildNumber >= 17763 ? EnumOsName.WindowsServer2019 :
                                                         EnumOsName.WindowsServer2016;
                  }

                  break;


               #endregion // Version 10


               #region Version 6

               case 6:
                  switch (verInfo.dwMinorVersion)
                  {
                     // Windows 8.1 or Windows Server 2012 R2
                     case 3:
                        _enumOsName = verInfo.wProductType == NativeMethods.VER_NT_WORKSTATION
                           ? EnumOsName.Windows81
                           : EnumOsName.WindowsServer2012R2;
                        break;


                     // Windows 8 or Windows Server 2012
                     case 2:
                        _enumOsName = verInfo.wProductType == NativeMethods.VER_NT_WORKSTATION
                           ? EnumOsName.Windows8
                           : EnumOsName.WindowsServer2012;
                        break;


                     // Windows 7 or Windows Server 2008 R2
                     case 1:
                        _enumOsName = verInfo.wProductType == NativeMethods.VER_NT_WORKSTATION
                           ? EnumOsName.Windows7
                           : EnumOsName.WindowsServer2008R2;
                        break;


                     // Windows Vista or Windows Server 2008
                     case 0:
                        _enumOsName = verInfo.wProductType == NativeMethods.VER_NT_WORKSTATION
                           ? EnumOsName.WindowsVista
                           : EnumOsName.WindowsServer2008;
                        break;
                        

                     default:
                        _enumOsName = EnumOsName.Later;
                        break;
                  }

                  break;

               #endregion // Version 6


               #region Version 5

               case 5:
                  switch (verInfo.dwMinorVersion)
                  {
                     case 2:
                        _enumOsName = verInfo.wProductType == NativeMethods.VER_NT_WORKSTATION && _processorArchitecture == EnumProcessorArchitecture.X64
                           ? EnumOsName.WindowsXP
                           : verInfo.wProductType != NativeMethods.VER_NT_WORKSTATION ? EnumOsName.WindowsServer2003 : EnumOsName.Later;
                        break;


                     case 1:
                        _enumOsName = EnumOsName.WindowsXP;
                        break;


                     case 0:
                        _enumOsName = EnumOsName.Windows2000;
                        break;


                     default:
                        _enumOsName = EnumOsName.Later;
                        break;
                  }
                  break;

               #endregion // Version 5


               default:
                  _enumOsName = EnumOsName.Earlier;
                  break;
            }
         }
      }
      

      private static class NativeMethods
      {
         internal const short VER_NT_WORKSTATION = 1;
         internal const short VER_NT_DOMAIN_CONTROLLER = 2;
         internal const short VER_NT_SERVER = 3;


         [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
         internal struct RTL_OSVERSIONINFOEXW
         {
            public int dwOSVersionInfoSize;
            public readonly int dwMajorVersion;
            public readonly int dwMinorVersion;
            public readonly int dwBuildNumber;
            public readonly int dwPlatformId;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
            public readonly string szCSDVersion;
            public readonly ushort wServicePackMajor;
            public readonly ushort wServicePackMinor;
            public readonly ushort wSuiteMask;
            public readonly byte wProductType;
            public readonly byte wReserved;
         }


         [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
         internal struct SYSTEM_INFO
         {
            public readonly EnumProcessorArchitecture wProcessorArchitecture;
            private readonly ushort wReserved;
            public readonly uint dwPageSize;
            public readonly IntPtr lpMinimumApplicationAddress;
            public readonly IntPtr lpMaximumApplicationAddress;
            public readonly IntPtr dwActiveProcessorMask;
            public readonly uint dwNumberOfProcessors;
            public readonly uint dwProcessorType;
            public readonly uint dwAllocationGranularity;
            public readonly ushort wProcessorLevel;
            public readonly ushort wProcessorRevision;
         }


         /// <summary>RtlGetVersion ルーチンは、現在実行中のオペレーティングシステムに関するバージョン情報を返します。</summary>
         /// <returns>RtlGetVersion は STATUS_SUCCESS を返します。</returns>
         /// <remarks>Windows 2000 以降で利用可能です。</remarks>
         [SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule"), DllImport("ntdll.dll", SetLastError = true, CharSet = CharSet.Unicode)]
         [return: MarshalAs(UnmanagedType.Bool)]
         internal static extern bool RtlGetVersion([MarshalAs(UnmanagedType.Struct)] ref RTL_OSVERSIONINFOEXW lpVersionInformation);


         /// <summary>WOW64 で実行されているアプリケーションに現在のシステムに関する情報を取得します。
         /// 64 ビットアプリケーションから呼び出された場合、GetSystemInfo 関数と同等です。
         /// </summary>
         /// <returns>この関数は値を返しません。</returns>
         /// <remarks>Win32 ベースのアプリケーションが WOW64 で実行されているかどうかを判定するには、<see cref="IsWow64Process"/> 関数を呼び出してください。</remarks>
         /// <remarks>サポートされる最小クライアント: Windows XP [デスクトップアプリ | Windows ストアアプリ]</remarks>
         /// <remarks>サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリ | Windows ストアアプリ]</remarks>
         [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
         [DllImport("kernel32.dll", SetLastError = false, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
         internal static extern void GetNativeSystemInfo([MarshalAs(UnmanagedType.Struct)] ref SYSTEM_INFO lpSystemInfo);


         /// <summary>指定されたプロセスが WOW64 で実行されているかどうかを判定します。</summary>
         /// <returns>
         /// 関数が成功した場合、戻り値はゼロ以外の値です。
         /// 関数が失敗した場合、戻り値はゼロです。拡張エラー情報を取得するには、GetLastError を呼び出してください。
         /// </returns>
         /// <remarks>サポートされる最小クライアント: Windows Vista、Windows XP SP2 [デスクトップアプリのみ]</remarks>
         /// <remarks>サポートされる最小サーバー: Windows Server 2008、Windows Server 2003 SP1 [デスクトップアプリのみ]</remarks>
         [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
         [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
         [return: MarshalAs(UnmanagedType.Bool)]
         internal static extern bool IsWow64Process([In] IntPtr hProcess, [Out, MarshalAs(UnmanagedType.Bool)] out bool lpSystemInfo);
      }

      #endregion // Private
   }
}
