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

using Microsoft.Win32.SafeHandles;
using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using System.Security;

namespace Alphaleonis.Win32.Filesystem
{
   internal static partial class NativeMethods
   {
      #region CM_Xxx

      /// <summary>CM_Connect_Machine 関数はリモートマシンへの接続を作成します。</summary>
      /// <remarks>
      ///   <para>Windows 8 および Windows Server 2012 以降、リモートマシンへのアクセス機能は削除されました。</para>
      ///   <para>これらのバージョンの Windows で実行している場合、リモートマシンにアクセスできません。</para>
      ///   <para>Microsoft Windows 2000 以降のバージョンの Windows で利用可能です。</para>
      /// </remarks>
      /// <param name="uncServerName">UNC サーバー名。</param>
      /// <param name="phMachine">[out] マシンハンドル。</param>
      /// <returns>
      ///   <para>操作が成功した場合、関数は CR_SUCCESS を返します。</para>
      ///   <para>それ以外の場合、Cfgmgr32.h で定義された CR_ プレフィックス付きのエラーコードの1つを返します。</para>
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("setupapi.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "CM_Connect_MachineW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.I4)]
      public static extern int CM_Connect_Machine([MarshalAs(UnmanagedType.LPWStr)] string uncServerName, out SafeCmConnectMachineHandle phMachine);

      /// <summary>
      ///   CM_Get_Device_ID_Ex 関数は、ローカルまたはリモートマシン上の指定されたデバイスインスタンスのデバイスインスタンス ID を取得します。
      /// </summary>
      /// <remarks>
      ///   <para>Windows 8 および Windows Server 2012 以降、リモートマシンへのアクセス機能は削除されました。</para>
      ///   <para>これらのバージョンの Windows で実行している場合、リモートマシンにアクセスできません。</para>
      ///   <para>&#160;</para>
      ///   <para>Microsoft Windows 2000 以降のバージョンの Windows で利用可能です。</para>
      /// </remarks>
      /// <param name="dnDevInst">デバイスインスタンス。</param>
      /// <param name="buffer">バッファ。</param>
      /// <param name="bufferLen">バッファの長さ。</param>
      /// <param name="ulFlags">フラグ。</param>
      /// <param name="hMachine">マシン。</param>
      /// <returns>
      ///   <para>操作が成功した場合、関数は CR_SUCCESS を返します。</para>
      ///   <para>それ以外の場合、Cfgmgr32.h で定義された CR_ プレフィックス付きのエラーコードの1つを返します。</para>
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("setupapi.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "CM_Get_Device_ID_ExW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.I4)]
      public static extern int CM_Get_Device_ID_Ex([MarshalAs(UnmanagedType.U4)] uint dnDevInst, SafeGlobalMemoryBufferHandle buffer, [MarshalAs(UnmanagedType.U4)] uint bufferLen, [MarshalAs(UnmanagedType.U4)] uint ulFlags, SafeCmConnectMachineHandle hMachine);

      /// <summary>
      ///   CM_Disconnect_Machine 関数はリモートマシンへの接続を削除します。
      /// </summary>
      /// <remarks>
      ///   <para>Windows 8 および Windows Server 2012 以降、リモートマシンへのアクセス機能は削除されました。</para>
      ///   <para>これらのバージョンの Windows で実行している場合、リモートマシンにアクセスできません。</para>
      ///   <para>SetLastError は <c>false</c> に設定されています。</para>
      ///   <para>Microsoft Windows 2000 以降のバージョンの Windows で利用可能です。</para>
      /// </remarks>
      /// <param name="hMachine">マシン。</param>
      /// <returns>
      ///   <para>操作が成功した場合、関数は CR_SUCCESS を返します。</para>
      ///   <para>それ以外の場合、Cfgmgr32.h で定義された CR_ プレフィックス付きのエラーコードの1つを返します。</para>
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("setupapi.dll", SetLastError = false, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.I4)]
      internal static extern int CM_Disconnect_Machine(IntPtr hMachine);

      /// <summary>
      ///   CM_Get_Parent_Ex 関数は、ローカルまたはリモートマシンのデバイスツリー内の指定されたデバイスノード (devnode) の
      ///   親ノードへのデバイスインスタンスハンドルを取得します。
      /// </summary>
      /// <remarks>
      ///   <para>Windows 8 および Windows Server 2012 以降、リモートマシンへのアクセス機能は削除されました。</para>
      ///   <para>これらのバージョンの Windows で実行している場合、リモートマシンにアクセスできません。</para>
      ///   <para>Microsoft Windows 2000 以降のバージョンの Windows で利用可能です。</para>
      /// </remarks>
      /// <param name="pdnDevInst">[out] 親デバイスインスタンス。</param>
      /// <param name="dnDevInst">デバイスインスタンス。</param>
      /// <param name="ulFlags">フラグ。</param>
      /// <param name="hMachine">マシン。</param>
      /// <returns>
      ///   <para>操作が成功した場合、関数は CR_SUCCESS を返します。</para>
      ///   <para>それ以外の場合、Cfgmgr32.h で定義された CR_ プレフィックス付きのエラーコードの1つを返します。</para>
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("setupapi.dll", SetLastError = true, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.I4)]
      internal static extern int CM_Get_Parent_Ex([MarshalAs(UnmanagedType.U4)] out uint pdnDevInst, [MarshalAs(UnmanagedType.U4)] uint dnDevInst, [MarshalAs(UnmanagedType.U4)] uint ulFlags, SafeCmConnectMachineHandle hMachine);

      #endregion // CM_Xxx

      #region DeviceIoControl

      /// <summary>指定されたデバイスドライバに直接制御コードを送信し、対応するデバイスに対応する操作を実行させます。</summary>
      /// <returns>
      ///   <para>操作が正常に完了した場合、戻り値はゼロ以外です。</para>
      ///   <para>操作が失敗したか保留中の場合、戻り値はゼロです。拡張エラー情報を取得するには GetLastError を呼び出してください。</para>
      /// </returns>
      /// <remarks>
      ///   <para>デバイスへのハンドルを取得するには、デバイスの名前またはデバイスに関連付けられたドライバの名前を指定して
      ///   <see cref="CreateFile"/> 関数を呼び出す必要があります。</para>
      ///   <para>デバイス名を指定するには、次の形式を使用します: <c>\\.\DeviceName</c></para>
      ///   <para>サポートされる最小クライアント: Windows XP</para>
      ///   <para>サポートされる最小サーバー: Windows Server 2003</para>
      /// </remarks>
      /// <param name="hDevice">デバイス。</param>
      /// <param name="dwIoControlCode">I/O 制御コード。</param>
      /// <param name="lpInBuffer">入力データ用バッファ。</param>
      /// <param name="nInBufferSize">入力バッファのサイズ。</param>
      /// <param name="lpOutBuffer">出力データ用バッファ。</param>
      /// <param name="nOutBufferSize">出力バッファのサイズ。</param>
      /// <param name="lpBytesReturned">[out] 返されたバイト数。</param>
      /// <param name="lpOverlapped">オーバーラップ構造体。</param>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool DeviceIoControl(SafeFileHandle hDevice, [MarshalAs(UnmanagedType.U4)] uint dwIoControlCode, IntPtr lpInBuffer, [MarshalAs(UnmanagedType.U4)] uint nInBufferSize, SafeGlobalMemoryBufferHandle lpOutBuffer, [MarshalAs(UnmanagedType.U4)] uint nOutBufferSize, [MarshalAs(UnmanagedType.U4)] out uint lpBytesReturned, IntPtr lpOverlapped);

      /// <summary>指定されたデバイスドライバに直接制御コードを送信し、対応するデバイスに対応する操作を実行させます。</summary>
      /// <returns>
      ///   <para>操作が正常に完了した場合、戻り値はゼロ以外です。</para>
      ///   <para>操作が失敗したか保留中の場合、戻り値はゼロです。拡張エラー情報を取得するには GetLastError を呼び出してください。</para>
      /// </returns>
      /// <remarks>
      ///   <para>デバイスへのハンドルを取得するには、デバイスの名前またはデバイスに関連付けられたドライバの名前を指定して
      ///   <see cref="CreateFile"/> 関数を呼び出す必要があります。</para>
      ///   <para>デバイス名を指定するには、次の形式を使用します: <c>\\.\DeviceName</c></para>
      ///   <para>サポートされる最小クライアント: Windows XP</para>
      ///   <para>サポートされる最小サーバー: Windows Server 2003</para>
      /// </remarks>
      /// <param name="hDevice">デバイス。</param>
      /// <param name="dwIoControlCode">I/O 制御コード。</param>
      /// <param name="lpInBuffer">入力データ用バッファ。</param>
      /// <param name="nInBufferSize">入力バッファのサイズ。</param>
      /// <param name="lpOutBuffer">出力データ用バッファ。</param>
      /// <param name="nOutBufferSize">出力バッファのサイズ。</param>
      /// <param name="lpBytesReturned">[out] 返されたバイト数。</param>
      /// <param name="lpOverlapped">オーバーラップ構造体。</param>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "DeviceIoControl"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool DeviceIoControl2(SafeFileHandle hDevice, [MarshalAs(UnmanagedType.U4)] uint dwIoControlCode, SafeGlobalMemoryBufferHandle lpInBuffer, [MarshalAs(UnmanagedType.U4)] uint nInBufferSize, IntPtr lpOutBuffer, [MarshalAs(UnmanagedType.U4)] uint nOutBufferSize, [MarshalAs(UnmanagedType.U4)] out uint lpBytesReturned, IntPtr lpOverlapped);

      /// <summary>指定されたデバイスドライバに直接制御コードを送信し、対応するデバイスに対応する操作を実行させます。</summary>
      /// <returns>
      ///   <para>操作が正常に完了した場合、戻り値はゼロ以外です。</para>
      ///   <para>操作が失敗したか保留中の場合、戻り値はゼロです。拡張エラー情報を取得するには GetLastError を呼び出してください。</para>
      /// </returns>
      /// <remarks>
      ///   <para>デバイスへのハンドルを取得するには、デバイスの名前またはデバイスに関連付けられたドライバの名前を指定して
      ///   <see cref="CreateFile"/> 関数を呼び出す必要があります。</para>
      ///   <para>デバイス名を指定するには、次の形式を使用します: <c>\\.\DeviceName</c></para>
      ///   <para>サポートされる最小クライアント: Windows XP</para>
      ///   <para>サポートされる最小サーバー: Windows Server 2003</para>
      /// </remarks>
      /// <param name="hDevice">デバイス。</param>
      /// <param name="dwIoControlCode">I/O 制御コード。</param>
      /// <param name="lpInBuffer">入力データ用バッファ。</param>
      /// <param name="nInBufferSize">入力バッファのサイズ。</param>
      /// <param name="lpOutBuffer">出力データ用バッファ。</param>
      /// <param name="nOutBufferSize">出力バッファのサイズ。</param>
      /// <param name="lpBytesReturned">[out] 返されたバイト数。</param>
      /// <param name="lpOverlapped">オーバーラップ構造体。</param>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "DeviceIoControl"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool DeviceIoControlUnknownSize(SafeFileHandle hDevice, [MarshalAs(UnmanagedType.U4)] uint dwIoControlCode, IntPtr lpInBuffer, [MarshalAs(UnmanagedType.U4)] uint nInBufferSize, IntPtr lpOutBuffer, [MarshalAs(UnmanagedType.U4)] uint nOutBufferSize, [MarshalAs(UnmanagedType.U4)] out uint lpBytesReturned, IntPtr lpOverlapped);

      #endregion // DeviceIoControl

      #region SetupDiXxx

      /// <summary>
      ///   SetupDiDestroyDeviceInfoList 関数はデバイス情報セットを削除し、関連するすべてのメモリを解放します。
      /// </summary>
      /// <remarks>
      ///   <para>SetLastError は <c>false</c> に設定されています。</para>
      ///   <para>Microsoft Windows 2000 以降のバージョンの Windows で利用可能です。</para>
      /// </remarks>
      /// <param name="hDevInfo">デバイス情報。</param>
      /// <returns>
      ///   <para>関数が成功した場合、TRUE を返します。</para>
      ///   <para>それ以外の場合、FALSE を返し、記録されたエラーは GetLastError の呼び出しで取得できます。</para>
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("setupapi.dll", SetLastError = false, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      private static extern bool SetupDiDestroyDeviceInfoList(IntPtr hDevInfo);

      /// <summary>
      ///   SetupDiEnumDeviceInterfaces 関数はデバイス情報セットに含まれるデバイスインターフェイスを列挙します。
      /// </summary>
      /// <remarks>
      ///   <para>この関数を繰り返し呼び出すと、異なるデバイスインターフェイスの <see cref="SP_DEVICE_INTERFACE_DATA"/> 構造体が返されます。</para>
      ///   <para>この関数を繰り返し呼び出すことで、デバイス情報セット内の特定のデバイス情報要素に関連付けられたインターフェイス、</para>
      ///   <para>またはすべてのデバイス情報要素に関連付けられたインターフェイスに関する情報を取得できます。</para>
      ///   <para>Microsoft Windows 2000 以降のバージョンの Windows で利用可能です。</para>
      /// </remarks>
      /// <param name="hDevInfo">デバイス情報。</param>
      /// <param name="devInfo">デバイス情報。</param>
      /// <param name="interfaceClassGuid">[in,out] インターフェイスクラスの一意識別子。</param>
      /// <param name="memberIndex">メンバーのゼロベースインデックス。</param>
      /// <param name="deviceInterfaceData">[in,out] デバイスインターフェイス情報。</param>
      /// <returns>
      ///   <para>関数がエラーなしで完了した場合、SetupDiEnumDeviceInterfaces は TRUE を返します。</para>
      ///   <para>関数がエラーで完了した場合、FALSE が返され、GetLastError を呼び出すことで失敗のエラーコードを取得できます。</para>
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("setupapi.dll", SetLastError = true, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool SetupDiEnumDeviceInterfaces(SafeHandle hDevInfo, IntPtr devInfo, ref Guid interfaceClassGuid, [MarshalAs(UnmanagedType.U4)] uint memberIndex, ref SP_DEVICE_INTERFACE_DATA deviceInterfaceData);

      /// <summary>
      ///   SetupDiGetClassDevsEx 関数は、ローカルまたはリモートコンピュータの要求されたデバイス情報要素を含む
      ///   デバイス情報セットへのハンドルを返します。
      /// </summary>
      /// <remarks>
      ///   <para>SetupDiGetClassDevsEx の呼び出し元は、不要になった返されたデバイス情報セットを
      ///   <see cref="SetupDiDestroyDeviceInfoList"/> を呼び出して削除する必要があります。</para>
      ///   <para>Microsoft Windows 2000 以降のバージョンの Windows で利用可能です。</para>
      /// </remarks>
      /// <param name="classGuid">[in,out] クラスの一意識別子。</param>
      /// <param name="enumerator">列挙子。</param>
      /// <param name="hwndParent">親ウィンドウ。</param>
      /// <param name="devsExFlags">デバイス拡張フラグ。</param>
      /// <param name="deviceInfoSet">デバイス情報が属するセット。</param>
      /// <param name="machineName">マシン名。</param>
      /// <param name="reserved">予約済み。</param>
      /// <returns>
      ///   <para>操作が成功した場合、SetupDiGetClassDevsEx は指定されたパラメータに一致するすべてのインストール済みデバイスを含む
      ///   デバイス情報セットへのハンドルを返します。</para>
      ///   <para>操作が失敗した場合、関数は INVALID_HANDLE_VALUE を返します。拡張エラー情報を取得するには
      ///   GetLastError を呼び出してください。</para>
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("setupapi.dll", SetLastError = true, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      internal static extern SafeSetupDiClassDevsExHandle SetupDiGetClassDevsEx(ref Guid classGuid, IntPtr enumerator, IntPtr hwndParent, [MarshalAs(UnmanagedType.U4)] SetupDiGetClassDevsExFlags devsExFlags, IntPtr deviceInfoSet, [MarshalAs(UnmanagedType.LPWStr)] string machineName, IntPtr reserved);

      /// <summary>
      ///   SetupDiGetDeviceInterfaceDetail 関数はデバイスインターフェイスの詳細を返します。
      /// </summary>
      /// <remarks>
      ///   <para>この関数が返すインターフェイスの詳細は、CreateFile などの Win32 関数に渡すことができるデバイスパスで構成されます。</para>
      ///   <para>デバイスパスのシンボリック名を解析しようとしないでください。デバイスパスはシステムの再起動をまたいで再利用できます。</para>
      ///   <para>Microsoft Windows 2000 以降のバージョンの Windows で利用可能です。</para>
      /// </remarks>
      /// <param name="hDevInfo">デバイス情報。</param>
      /// <param name="deviceInterfaceData">[in,out] デバイスインターフェイス情報。</param>
      /// <param name="deviceInterfaceDetailData">[in,out] デバイスインターフェイスの詳細情報。</param>
      /// <param name="deviceInterfaceDetailDataSize">デバイスインターフェイスの詳細データのサイズ。</param>
      /// <param name="requiredSize">必要なサイズ。</param>
      /// <param name="deviceInfoData">[in,out] デバイス情報データ。</param>
      /// <returns>
      ///   <para>関数がエラーなしで完了した場合、SetupDiGetDeviceInterfaceDetail は TRUE を返します。</para>
      ///   <para>関数がエラーで完了した場合、FALSE が返され、GetLastError を呼び出すことで失敗のエラーコードを取得できます。</para>
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("setupapi.dll", SetLastError = true, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool SetupDiGetDeviceInterfaceDetail(SafeHandle hDevInfo, ref SP_DEVICE_INTERFACE_DATA deviceInterfaceData, ref SP_DEVICE_INTERFACE_DETAIL_DATA deviceInterfaceDetailData, [MarshalAs(UnmanagedType.U4)] uint deviceInterfaceDetailDataSize, IntPtr requiredSize, ref SP_DEVINFO_DATA deviceInfoData);

      /// <summary>
      ///   SetupDiGetDeviceRegistryProperty 関数は指定されたプラグアンドプレイデバイスプロパティを取得します。
      /// </summary>
      /// <remarks><para>Microsoft Windows 2000 以降のバージョンの Windows で利用可能です。</para></remarks>
      /// <param name="deviceInfoSet">デバイス情報が属するセット。</param>
      /// <param name="deviceInfoData">[in,out] デバイス情報データ。</param>
      /// <param name="property">プロパティ。</param>
      /// <param name="propertyRegDataType">[out] プロパティレジストリデータの型。</param>
      /// <param name="propertyBuffer">プロパティデータ用バッファ。</param>
      /// <param name="propertyBufferSize">プロパティバッファのサイズ。</param>
      /// <param name="requiredSize">必要なサイズ。</param>
      /// <returns>
      ///   <para>呼び出しが成功した場合、SetupDiGetDeviceRegistryProperty は TRUE を返します。</para>
      ///   <para>それ以外の場合、FALSE を返し、記録されたエラーは GetLastError の呼び出しで取得できます。</para>
      ///   <para>要求されたプロパティがデバイスに存在しない場合、またはプロパティデータが有効でない場合、
      ///   SetupDiGetDeviceRegistryProperty は ERROR_INVALID_DATA エラーコードを返します。</para>
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("setupapi.dll", SetLastError = true, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool SetupDiGetDeviceRegistryProperty(SafeHandle deviceInfoSet, ref SP_DEVINFO_DATA deviceInfoData, SetupDiGetDeviceRegistryPropertyEnum property, IntPtr propertyRegDataType, SafeGlobalMemoryBufferHandle propertyBuffer, [MarshalAs(UnmanagedType.U4)] uint propertyBufferSize, IntPtr requiredSize);

      #endregion // SetupDiXxx
   }
}
