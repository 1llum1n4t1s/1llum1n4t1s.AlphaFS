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
using System.IO;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;

namespace Alphaleonis.Win32.Filesystem
{
   internal static partial class NativeMethods
   {  
      /// <summary>MS-DOS デバイス名を定義、再定義、または削除します。</summary>
      /// <returns>
      /// 関数が成功した場合、戻り値はゼロ以外です。
      /// 関数が失敗した場合、戻り値はゼロです。拡張エラー情報を取得するには GetLastError を呼び出してください。
      /// </returns>
      /// <remarks>サポートされる最小クライアント: Windows XP [デスクトップアプリのみ]</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリのみ]</remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "DefineDosDeviceW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool DefineDosDevice(DosDeviceAttributes dwFlags, [MarshalAs(UnmanagedType.LPWStr)] string lpDeviceName, [MarshalAs(UnmanagedType.LPWStr)] string lpTargetPath);

      /// <summary>ドライブ文字またはマウントされたフォルダを削除します。</summary>
      /// <returns>
      /// 関数が成功した場合、戻り値はゼロ以外です。
      /// 関数が失敗した場合、戻り値はゼロです。拡張エラー情報を取得するには GetLastError を呼び出してください。
      /// </returns>
      /// <remarks>サポートされる最小クライアント: Windows XP [デスクトップアプリのみ]</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリのみ]</remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "DeleteVolumeMountPointW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool DeleteVolumeMountPoint([MarshalAs(UnmanagedType.LPWStr)] string lpszVolumeMountPoint);

      /// <summary>コンピュータ上のボリュームの名前を取得します。FindFirstVolume はコンピュータのボリュームのスキャンを開始するために使用されます。</summary>
      /// <returns>
      /// 関数が成功した場合、戻り値は後続の FindNextVolume および FindVolumeClose 関数の呼び出しで使用される検索ハンドルです。
      /// 関数がボリュームを見つけられなかった場合、戻り値は INVALID_HANDLE_VALUE エラーコードです。拡張エラー情報を取得するには GetLastError を呼び出してください。
      /// </returns>
      /// <remarks>サポートされる最小クライアント: Windows XP [デスクトップアプリのみ]</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリのみ]</remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "FindFirstVolumeW"), SuppressUnmanagedCodeSecurity]
      internal static extern SafeFindVolumeHandle FindFirstVolume(StringBuilder lpszVolumeName, [MarshalAs(UnmanagedType.U4)] uint cchBufferLength);

      /// <summary>指定されたボリューム上のマウントされたフォルダの名前を取得します。FindFirstVolumeMountPoint はボリューム上のマウントされたフォルダのスキャンを開始するために使用されます。</summary>
      /// <returns>
      /// 関数が成功した場合、戻り値は後続の FindNextVolumeMountPoint および FindVolumeMountPointClose 関数の呼び出しで使用される検索ハンドルです。
      /// 関数がボリューム上のマウントされたフォルダを見つけられなかった場合、戻り値は INVALID_HANDLE_VALUE エラーコードです。
      /// </returns>
      /// <remarks>サポートされる最小クライアント: Windows XP [デスクトップアプリのみ]</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリのみ]</remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "FindFirstVolumeMountPointW"), SuppressUnmanagedCodeSecurity]
      internal static extern SafeFindVolumeMountPointHandle FindFirstVolumeMountPoint([MarshalAs(UnmanagedType.LPWStr)] string lpszRootPathName, StringBuilder lpszVolumeMountPoint, [MarshalAs(UnmanagedType.U4)] uint cchBufferLength);

      /// <summary>FindFirstVolume 関数の呼び出しで開始されたボリューム検索を続行します。FindNextVolume は呼び出しごとに1つのボリュームを検索します。</summary>
      /// <returns>
      /// 関数が成功した場合、戻り値はゼロ以外です。
      /// 関数が失敗した場合、戻り値はゼロです。拡張エラー情報を取得するには GetLastError を呼び出してください。
      /// </returns>
      /// <remarks>サポートされる最小クライアント: Windows XP [デスクトップアプリのみ]</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリのみ]</remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "FindNextVolumeW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool FindNextVolume(SafeFindVolumeHandle hFindVolume, StringBuilder lpszVolumeName, [MarshalAs(UnmanagedType.U4)] uint cchBufferLength);

      /// <summary>FindFirstVolumeMountPoint 関数の呼び出しで開始されたマウントされたフォルダの検索を続行します。FindNextVolumeMountPoint は呼び出しごとに1つのマウントされたフォルダを検索します。</summary>
      /// <returns>
      /// 関数が成功した場合、戻り値はゼロ以外です。
      /// 関数が失敗した場合、戻り値はゼロです。拡張エラー情報を取得するには GetLastError を呼び出してください。マウントされたフォルダがこれ以上見つからない場合、GetLastError 関数は ERROR_NO_MORE_FILES エラーコードを返します。
      /// その場合は、FindVolumeMountPointClose 関数で検索を閉じてください。
      /// </returns>
      /// <remarks>サポートされる最小クライアント: Windows XP</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2003</remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "FindNextVolumeMountPointW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool FindNextVolumeMountPoint(SafeFindVolumeMountPointHandle hFindVolume, StringBuilder lpszVolumeName, [MarshalAs(UnmanagedType.U4)] uint cchBufferLength);

      /// <summary>指定されたボリューム検索ハンドルを閉じます。</summary>
      /// <remarks>
      ///   <para>SetLastError は <c>false</c> に設定されています。</para>
      ///   サポートされる最小クライアント: Windows XP [デスクトップアプリのみ]。サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリのみ]。
      /// </remarks>
      /// <returns>
      ///   関数が成功した場合、戻り値はゼロ以外です。関数が失敗した場合、戻り値はゼロです。拡張エラー情報を取得するには
      ///   GetLastError を呼び出してください。
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = false, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool FindVolumeClose(IntPtr hFindVolume);

      /// <summary>指定されたマウントされたフォルダの検索ハンドルを閉じます。</summary>
      /// <remarks>
      ///   <para>SetLastError は <c>false</c> に設定されています。</para>
      ///   <para>サポートされる最小クライアント: Windows XP</para>
      ///   <para>サポートされる最小サーバー: Windows Server 2003</para>
      /// </remarks>
      /// <returns>
      ///   関数が成功した場合、戻り値はゼロ以外です。関数が失敗した場合、戻り値はゼロです。拡張エラー情報を取得するには
      ///   GetLastError を呼び出してください。
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = false, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool FindVolumeMountPointClose(IntPtr hFindVolume);

      /// <summary>
      ///   ディスクドライブがリムーバブル、固定、CD-ROM、RAM ディスク、またはネットワークドライブかどうかを判断します。
      ///   <para>ドライブが USB タイプのドライブかどうかを判断するには、<see cref="SetupDiGetDeviceRegistryProperty"/> を呼び出して
      ///   SPDRP_REMOVAL_POLICY プロパティを指定します。</para>
      /// </summary>
      /// <remarks>
      ///   <para>SMB はボリューム管理関数をサポートしていません。</para>
      ///   <para>サポートされる最小クライアント: Windows XP [デスクトップアプリのみ]</para>
      ///   <para>サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリのみ]</para>
      /// </remarks>
      /// <param name="lpRootPathName">ルートファイルの完全パス名。</param>
      /// <returns>
      ///   <para>戻り値はドライブの種類を指定します。<see cref="DriveType"/> を参照してください。</para>
      ///   <para>関数が失敗した場合、戻り値はゼロです。拡張エラー情報を取得するには GetLastError を呼び出してください。</para>
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "GetDriveTypeW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.U4)]
      internal static extern DriveType GetDriveType([MarshalAs(UnmanagedType.LPWStr)] string lpRootPathName);

      /// <summary>
      ///   現在利用可能なディスクドライブを表すビットマスクを取得します。
      /// </summary>
      /// <remarks>
      ///   <para>SMB はボリューム管理関数をサポートしていません。</para>
      ///   <para>サポートされる最小クライアント: Windows XP [デスクトップアプリのみ]</para>
      ///   <para>サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリのみ]</para>
      /// </remarks>
      /// <returns>
      ///   <para>関数が成功した場合、戻り値は現在利用可能なディスクドライブを表すビットマスクです。</para>
      ///   <para>ビット位置 0 (最下位ビット) はドライブ A、ビット位置 1 はドライブ B、ビット位置 2 はドライブ C、以下同様です。</para>
      ///   <para>関数が失敗した場合、戻り値はゼロです。拡張エラー情報を取得するには GetLastError を呼び出してください。</para>
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.U4)]
      internal static extern uint GetLogicalDrives();

      /// <summary>指定されたルートディレクトリに関連付けられたファイルシステムとボリュームに関する情報を取得します。</summary>
      /// <returns>
      /// 要求されたすべての情報が取得された場合、戻り値はゼロ以外です。
      /// 要求されたすべての情報が取得されなかった場合、戻り値はゼロです。
      /// </returns>
      /// <remarks>サポートされる最小クライアント: Windows XP [デスクトップアプリのみ]</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリのみ]</remarks>
      /// <remarks>"lpRootPathName" は末尾のバックスラッシュで終わる必要があります。</remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "GetVolumeInformationW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool GetVolumeInformation([MarshalAs(UnmanagedType.LPWStr)] string lpRootPathName, StringBuilder lpVolumeNameBuffer, [MarshalAs(UnmanagedType.U4)] uint nVolumeNameSize, [MarshalAs(UnmanagedType.U4)] out uint lpVolumeSerialNumber, [MarshalAs(UnmanagedType.U4)] out int lpMaximumComponentLength, [MarshalAs(UnmanagedType.U4)] out VOLUME_INFO_FLAGS lpFileSystemAttributes, StringBuilder lpFileSystemNameBuffer, [MarshalAs(UnmanagedType.U4)] uint nFileSystemNameSize);

      /// <summary>指定されたファイルに関連付けられたファイルシステムとボリュームに関する情報を取得します。</summary>
      /// <returns>
      /// 要求されたすべての情報が取得された場合、戻り値はゼロ以外です。
      /// 要求されたすべての情報が取得されなかった場合、戻り値はゼロです。拡張エラー情報を取得するには GetLastError を呼び出してください。
      /// </returns>
      /// <remarks>ファイルまたはディレクトリの現在の圧縮状態を取得するには、FSCTL_GET_COMPRESSION を使用してください。</remarks>
      /// <remarks>SMB はボリューム管理関数をサポートしていません。</remarks>
      /// <remarks>サポートされる最小クライアント: Windows Vista [デスクトップアプリのみ]</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2008 [デスクトップアプリのみ]</remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "GetVolumeInformationByHandleW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool GetVolumeInformationByHandle(SafeFileHandle hFile, StringBuilder lpVolumeNameBuffer, [MarshalAs(UnmanagedType.U4)] uint nVolumeNameSize, [MarshalAs(UnmanagedType.U4)] out uint lpVolumeSerialNumber, [MarshalAs(UnmanagedType.U4)] out int lpMaximumComponentLength, out VOLUME_INFO_FLAGS lpFileSystemAttributes, StringBuilder lpFileSystemNameBuffer, [MarshalAs(UnmanagedType.U4)] uint nFileSystemNameSize);

      /// <summary>指定されたボリュームマウントポイント (ドライブ文字、ボリューム GUID パス、またはマウントされたフォルダ) に関連付けられたボリュームのボリューム GUID パスを取得します。</summary>
      /// <returns>
      /// 関数が成功した場合、戻り値はゼロ以外です。
      /// 関数が失敗した場合、戻り値はゼロです。拡張エラー情報を取得するには GetLastError を呼び出してください。
      /// </returns>
      /// <remarks>入力パラメータとしてボリューム GUID パスを必要とする SetVolumeMountPoint や FindFirstVolumeMountPoint などの関数で使用するボリューム GUID パスを取得するには、GetVolumeNameForVolumeMountPoint を使用してください。</remarks>
      /// <remarks>SMB はボリューム管理関数をサポートしていません。</remarks>
      /// <remarks>マウントポイントは ReFS ボリュームではサポートされていません。</remarks>
      /// <remarks>サポートされる最小クライアント: Windows XP [デスクトップアプリのみ]</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリのみ]</remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "GetVolumeNameForVolumeMountPointW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool GetVolumeNameForVolumeMountPoint([MarshalAs(UnmanagedType.LPWStr)] string lpszVolumeMountPoint, StringBuilder lpszVolumeName, [MarshalAs(UnmanagedType.U4)] uint cchBufferLength);

      /// <summary>指定されたパスがマウントされているボリュームマウントポイントを取得します。</summary>
      /// <remarks>
      ///   <para>指定されたパスが渡された場合、GetVolumePathName はボリュームマウントポイントへのパスを返します。
      ///   これは、指定されたパスのエンドポイントが存在するボリュームのルートを返すことを意味します。</para>
      ///   <para>例えば、ボリューム D が C:\Mnt\Ddrive にマウントされ、ボリューム E が "C:\Mnt\Ddrive\Mnt\Edrive" にマウントされているとします。
      ///   また、パス "E:\Dir\Subdir\MyFile" のファイルがあるとします。</para>
      ///   <para>"C:\Mnt\Ddrive\Mnt\Edrive\Dir\Subdir\MyFile" を GetVolumePathName に渡すと、パス "C:\Mnt\Ddrive\Mnt\Edrive\" が返されます。</para>
      ///   <para>ネットワーク共有が指定された場合、GetVolumePathName は GetDriveType が DRIVE_REMOTE を返す最短パスを返します。
      ///   これは、パスが現在のユーザーがアクセスできる存在するリモートドライブとして検証されることを意味します。</para>
      ///   <para>サポートされる最小クライアント: Windows XP [デスクトップアプリのみ]</para>
      ///   <para>サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリのみ]</para>
      /// </remarks>
      /// <returns>
      ///   <para>関数が成功した場合、戻り値はゼロ以外です。</para>
      ///   <para>関数が失敗した場合、戻り値はゼロです。拡張エラー情報を取得するには GetLastError を呼び出してください。</para>
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "GetVolumePathNameW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool GetVolumePathName([MarshalAs(UnmanagedType.LPWStr)] string lpszFileName, StringBuilder lpszVolumePathName, [MarshalAs(UnmanagedType.U4)] uint cchBufferLength);

      /// <summary>指定されたボリュームのドライブ文字とマウントされたフォルダパスの一覧を取得します。</summary>
      /// <remarks>サポートされる最小クライアント: Windows XP。</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2003。</remarks>
      /// <returns>
      ///   関数が成功した場合、戻り値はゼロ以外です。関数が失敗した場合、戻り値はゼロです。拡張エラー情報を取得するには
      ///   GetLastError を呼び出してください。
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "GetVolumePathNamesForVolumeNameW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool GetVolumePathNamesForVolumeName([MarshalAs(UnmanagedType.LPWStr)] string lpszVolumeName, char[] lpszVolumePathNames, [MarshalAs(UnmanagedType.U4)] uint cchBuferLength, [MarshalAs(UnmanagedType.U4)] out uint lpcchReturnLength);

      /// <summary>ファイルシステムボリュームのラベルを設定します。</summary>
      /// <remarks>サポートされる最小クライアント: Windows XP [デスクトップアプリのみ]。</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリのみ]。</remarks>
      /// <remarks>"lpRootPathName" は末尾のバックスラッシュで終わる必要があります。</remarks>
      /// <returns>
      ///   関数が成功した場合、戻り値はゼロ以外です。関数が失敗した場合、戻り値はゼロです。拡張エラー情報を取得するには
      ///   GetLastError を呼び出してください。
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "SetVolumeLabelW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool SetVolumeLabel([MarshalAs(UnmanagedType.LPWStr)] string lpRootPathName, [MarshalAs(UnmanagedType.LPWStr)] string lpVolumeName);

      /// <summary>ボリュームをドライブ文字または別のボリューム上のディレクトリに関連付けます。</summary>
      /// <remarks>サポートされる最小クライアント: Windows XP [デスクトップアプリのみ]。</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリのみ]。</remarks>
      /// <returns>
      ///   関数が成功した場合、戻り値はゼロ以外です。関数が失敗した場合、戻り値はゼロです。拡張エラー情報を取得するには
      ///   GetLastError を呼び出してください。
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "SetVolumeMountPointW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool SetVolumeMountPoint([MarshalAs(UnmanagedType.LPWStr)] string lpszVolumeMountPoint, [MarshalAs(UnmanagedType.LPWStr)] string lpszVolumeName);

      /// <summary>MS-DOS デバイス名に関する情報を取得します。</summary>
      /// <remarks>サポートされる最小クライアント: Windows XP [デスクトップアプリのみ]。</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリのみ]。</remarks>
      /// <returns>
      ///   関数が成功した場合、戻り値は lpTargetPath が指すバッファに格納された TCHAR の数です。
      ///   関数が失敗した場合、戻り値はゼロです。拡張エラー情報を取得するには GetLastError を呼び出してください。バッファが小さすぎる場合、
      ///   関数は失敗し、最後のエラーコードは ERROR_INSUFFICIENT_BUFFER です。
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "QueryDosDeviceW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.U4)]
      internal static extern uint QueryDosDevice([MarshalAs(UnmanagedType.LPWStr)] string lpDeviceName, char[] lpTargetPath, [MarshalAs(UnmanagedType.U4)] uint ucchMax);
   }
}