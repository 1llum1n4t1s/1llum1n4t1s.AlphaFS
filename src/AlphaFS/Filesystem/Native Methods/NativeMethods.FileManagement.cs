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
using Alphaleonis.Win32.Security;
using Microsoft.Win32.SafeHandles;
using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Runtime.InteropServices;
using System.Security;
using System.Security.AccessControl;
using System.Text;

namespace Alphaleonis.Win32.Filesystem
{
   internal static partial class NativeMethods
   {
      /// <summary>
      ///   既存のファイルを新しいファイルにコピーし、コールバック関数を通じてアプリケーションに進捗を通知します。
      /// </summary>
      /// <remarks>
      ///   <para>コピー先ファイルが既に存在し、FILE_ATTRIBUTE_HIDDEN または FILE_ATTRIBUTE_READONLY 属性が設定されている場合、
      ///   この関数は ERROR_ACCESS_DENIED で失敗します。</para>
      ///   <para>この関数は拡張属性、OLE 構造化ストレージ、NTFS ファイルシステムの代替データストリーム、セキュリティリソース属性、
      ///   およびファイル属性を保持します。</para>
      ///   <para>Windows 7、Windows Server 2008 R2、Windows Server 2008、Windows Vista、Windows Server 2003、および Windows XP:
      ///   既存ファイルのセキュリティリソース属性 (ATTRIBUTE_SECURITY_INFORMATION) は Windows 8 および Windows Server 2012 まで
      ///   新しいファイルにコピーされません。</para>
      ///   <para>サポートされる最小クライアント: Windows XP [デスクトップアプリのみ]</para>
      ///   <para>サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリのみ]</para>
      /// </remarks>
      /// <param name="lpExistingFileName">既存ファイルのファイル名。</param>
      /// <param name="lpNewFileName">新しいファイルのファイル名。</param>
      /// <param name="lpProgressRoutine">進捗ルーチン。</param>
      /// <param name="lpData">データ。</param>
      /// <param name="pbCancel">[out] キャンセルフラグ。</param>
      /// <param name="dwCopyFlags">コピーフラグ。</param>
      /// <returns>
      ///   <para>関数が成功した場合、戻り値はゼロ以外です。</para>
      ///   <para>関数が失敗した場合、戻り値はゼロです。拡張エラー情報を取得するには GetLastError を呼び出してください。</para>
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "CopyFileExW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool CopyFileEx([MarshalAs(UnmanagedType.LPWStr)] string lpExistingFileName, [MarshalAs(UnmanagedType.LPWStr)] string lpNewFileName, NativeCopyMoveProgressRoutine lpProgressRoutine, IntPtr lpData, [MarshalAs(UnmanagedType.Bool)] out bool pbCancel, CopyOptions dwCopyFlags);

      /// <summary>
      ///   トランザクション操作として、既存のファイルを新しいファイルにコピーし、コールバック関数を通じてアプリケーションに進捗を通知します。
      /// </summary>
      /// <remarks>
      ///   <para>コピー先ファイルが既に存在し、FILE_ATTRIBUTE_HIDDEN または FILE_ATTRIBUTE_READONLY 属性が設定されている場合、
      ///   この関数は ERROR_ACCESS_DENIED で失敗します。</para>
      ///   <para>この関数は拡張属性、OLE 構造化ストレージ、NTFS ファイルシステムの代替データストリーム、セキュリティリソース属性、
      ///   およびファイル属性を保持します。</para>
      ///   <para>Windows 7、Windows Server 2008 R2、Windows Server 2008、Windows Vista、Windows Server 2003、および Windows XP:
      ///   既存ファイルのセキュリティリソース属性 (ATTRIBUTE_SECURITY_INFORMATION) は Windows 8 および Windows Server 2012 まで
      ///   新しいファイルにコピーされません。</para>
      ///   <para>サポートされる最小クライアント: Windows Vista [デスクトップアプリのみ]</para>
      ///   <para>サポートされる最小サーバー: Windows Server 2008 [デスクトップアプリのみ]</para>
      /// </remarks>
      /// <returns>
      ///   <para>関数が成功した場合、戻り値はゼロ以外です。</para>
      ///   <para>関数が失敗した場合、戻り値はゼロです。拡張エラー情報を取得するには GetLastError を呼び出してください。</para>
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "CopyFileTransactedW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool CopyFileTransacted([MarshalAs(UnmanagedType.LPWStr)] string lpExistingFileName, [MarshalAs(UnmanagedType.LPWStr)] string lpNewFileName, NativeCopyMoveProgressRoutine lpProgressRoutine, IntPtr lpData, [MarshalAs(UnmanagedType.Bool)] out bool pbCancel, CopyOptions dwCopyFlags, SafeHandle hTransaction);

      /// <summary>
      ///   ファイルまたは I/O デバイスを作成または開きます。最も一般的に使用される I/O デバイスは、ファイル、ファイルストリーム、ディレクトリ、
      ///   物理ディスク、ボリューム、コンソールバッファ、テープドライブ、通信リソース、メールスロット、およびパイプです。
      /// </summary>
      /// <remarks>サポートされる最小クライアント: Windows XP。</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2003。</remarks>
      /// <returns>
      ///   関数が成功した場合、戻り値は指定されたファイル、デバイス、名前付きパイプ、またはメールスロットへの開いたハンドルです。
      ///   関数が失敗した場合、戻り値は Win32Errors.ERROR_INVALID_HANDLE です。拡張エラー情報を取得するには GetLastError を呼び出してください。
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "CreateFileW"), SuppressUnmanagedCodeSecurity]
      internal static extern SafeFileHandle CreateFile([MarshalAs(UnmanagedType.LPWStr)] string lpFileName, [MarshalAs(UnmanagedType.U4)] FileSystemRights dwDesiredAccess, [MarshalAs(UnmanagedType.U4)] FileShare dwShareMode, [MarshalAs(UnmanagedType.LPStruct)] Security.NativeMethods.SecurityAttributes lpSecurityAttributes, [MarshalAs(UnmanagedType.U4)] FileMode dwCreationDisposition, [MarshalAs(UnmanagedType.U4)] ExtendedFileAttributes dwFlagsAndAttributes, IntPtr hTemplateFile);

      /// <summary>
      ///   ファイルまたは I/O デバイスを作成または開きます。最も一般的に使用される I/O デバイスは、ファイル、ファイルストリーム、ディレクトリ、
      ///   物理ディスク、ボリューム、コンソールバッファ、テープドライブ、通信リソース、メールスロット、およびパイプです。
      /// </summary>
      /// <remarks>サポートされる最小クライアント: Windows Vista [デスクトップアプリのみ]。</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2008 [デスクトップアプリのみ]。</remarks>
      /// <returns>
      ///   関数が成功した場合、戻り値は指定されたファイル、デバイス、名前付きパイプ、またはメールスロットへの開いたハンドルです。
      ///   関数が失敗した場合、戻り値は Win32Errors.ERROR_INVALID_HANDLE です。拡張エラー情報を取得するには GetLastError を呼び出してください。
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "CreateFileTransactedW"), SuppressUnmanagedCodeSecurity]
      internal static extern SafeFileHandle CreateFileTransacted([MarshalAs(UnmanagedType.LPWStr)] string lpFileName, [MarshalAs(UnmanagedType.U4)] FileSystemRights dwDesiredAccess, [MarshalAs(UnmanagedType.U4)] FileShare dwShareMode, [MarshalAs(UnmanagedType.LPStruct)] Security.NativeMethods.SecurityAttributes lpSecurityAttributes, [MarshalAs(UnmanagedType.U4)] FileMode dwCreationDisposition, [MarshalAs(UnmanagedType.U4)] ExtendedFileAttributes dwFlagsAndAttributes, IntPtr hTemplateFile, SafeHandle hTransaction, IntPtr pusMiniVersion, IntPtr pExtendedParameter);

      /// <summary>指定されたファイルの名前付きまたは名前なしのファイルマッピングオブジェクトを作成または開きます。</summary>
      /// <remarks>サポートされる最小クライアント: Windows XP。</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2003。</remarks>
      /// <returns>
      ///   関数が成功した場合、戻り値は新しく作成されたファイルマッピングオブジェクトへのハンドルです。関数が失敗した場合、
      ///   戻り値は <c>null</c> です。
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = false, CharSet = CharSet.Unicode, EntryPoint = "CreateFileMappingW"), SuppressUnmanagedCodeSecurity]
      internal static extern SafeFileHandle CreateFileMapping(SafeFileHandle hFile, SafeHandle lpSecurityAttributes, [MarshalAs(UnmanagedType.U4)] uint flProtect, [MarshalAs(UnmanagedType.U4)] uint dwMaximumSizeHigh, [MarshalAs(UnmanagedType.U4)] uint dwMaximumSizeLow, [MarshalAs(UnmanagedType.LPWStr)] string lpName);

      /// <summary>既存のファイルと新しいファイルの間にハードリンク (CMD コマンド "MKLINK /H" に類似) を確立します。この関数は NTFS ファイルシステムでのみサポートされ、ファイルのみが対象でディレクトリは対象外です。</summary>
      /// <remarks>サポートされる最小クライアント: Windows XP [デスクトップアプリのみ]。</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリのみ]。</remarks>
      /// <returns>
      ///   If the function succeeds, the return value is nonzero. If the function fails, the return value is zero (0). To get extended error
      ///   information, call GetLastError.
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "CreateHardLinkW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool CreateHardLink([MarshalAs(UnmanagedType.LPWStr)] string lpFileName, [MarshalAs(UnmanagedType.LPWStr)] string lpExistingFileName, IntPtr lpSecurityAttributes);

      /// <summary>トランザクション操作として、既存のファイルと新しいファイルの間にハードリンク (CMD コマンド "MKLINK /H" に類似) を確立します。この関数は NTFS ファイルシステムでのみサポートされ、ファイルのみが対象でディレクトリは対象外です。</summary>
      /// <remarks>サポートされる最小クライアント: Windows Vista [デスクトップアプリのみ]。</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2008 [デスクトップアプリのみ]。</remarks>
      /// <returns>
      ///   If the function succeeds, the return value is nonzero. If the function fails, the return value is zero (0). To get extended error
      ///   information, call GetLastError.
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "CreateHardLinkTransactedW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool CreateHardLinkTransacted([MarshalAs(UnmanagedType.LPWStr)] string lpFileName, [MarshalAs(UnmanagedType.LPWStr)] string lpExistingFileName, IntPtr lpSecurityAttributes, SafeHandle hTransaction);

      /// <summary>シンボリックリンク (CMD コマンド "MKLINK /D" に類似) を作成します。</summary>
      /// <remarks>サポートされる最小クライアント: Windows Vista [デスクトップアプリのみ]。</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2008 [デスクトップアプリのみ]。</remarks>
      /// <remarks>
      /// アンマネージドプロトタイプには return ディレクティブが含まれています。CreateSymbolicLink API 関数は 1 バイトのデータ型である BOOLEAN を返すためです。
      /// bool のデフォルトのマーシャリングは 4 バイトです (BOOL 戻り値とのシームレスな統合を可能にするため)。
      /// BOOLEAN 値のデフォルトのマーシャリングを使用すると、誤った結果が得られる可能性があります。
      /// return ディレクティブにより、PInvoke は戻り値の 1 バイトのみをマーシャリングします。
      /// Source: http://www.informit.com/guides/content.aspx?g=dotnet&amp;seqNum=762&amp;ns=16196
      /// </remarks>
      /// <returns>
      ///   If the function succeeds, the return value is nonzero. If the function fails, the return value is zero (0). To get extended error
      ///   information, call GetLastError.
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "CreateSymbolicLinkW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.I1)]
      internal static extern bool CreateSymbolicLink([MarshalAs(UnmanagedType.LPWStr)] string lpSymlinkFileName, [MarshalAs(UnmanagedType.LPWStr)] string lpTargetFileName, [MarshalAs(UnmanagedType.U4)] SymbolicLinkTarget dwFlags);

      /// <summary>トランザクション操作として、シンボリックリンク (CMD コマンド "MKLINK /D" に類似) を作成します。</summary>
      /// <remarks>サポートされる最小クライアント: Windows Vista [デスクトップアプリのみ]。</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2008 [デスクトップアプリのみ]。</remarks>
      /// <remarks>
      /// アンマネージドプロトタイプには return ディレクティブが含まれています。CreateSymbolicLink API 関数は 1 バイトのデータ型である BOOLEAN を返すためです。
      /// bool のデフォルトのマーシャリングは 4 バイトです (BOOL 戻り値とのシームレスな統合を可能にするため)。
      /// BOOLEAN 値のデフォルトのマーシャリングを使用すると、誤った結果が得られる可能性があります。
      /// return ディレクティブにより、PInvoke は戻り値の 1 バイトのみをマーシャリングします。
      /// Source: http://www.informit.com/guides/content.aspx?g=dotnet&amp;seqNum=762&amp;ns=16196
      /// </remarks>
      /// <returns>
      ///   If the function succeeds, the return value is nonzero. If the function fails, the return value is zero (0). To get extended error
      ///   information, call GetLastError.
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "CreateSymbolicLinkTransactedW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.I1)]
      internal static extern bool CreateSymbolicLinkTransacted([MarshalAs(UnmanagedType.LPWStr)] string lpSymlinkFileName, [MarshalAs(UnmanagedType.LPWStr)] string lpTargetFileName, [MarshalAs(UnmanagedType.U4)] SymbolicLinkTarget dwFlags, SafeHandle hTransaction);

      /// <summary>暗号化されたファイルまたはディレクトリを復号します。</summary>
      /// <remarks>
      ///   DecryptFile 関数は復号対象のファイルへの排他的アクセスを必要とし、別のプロセスがファイルを使用している場合は失敗します。
      ///   ファイルが暗号化されていない場合、DecryptFile は単にゼロ以外の値を返し、成功を示します。lpFileName が読み取り専用ファイルを指定した場合、
      ///   関数は失敗し、GetLastError は ERROR_FILE_READ_ONLY を返します。lpFileName が読み取り専用ファイルを含むディレクトリを指定した場合、
      ///   関数は成功しますがディレクトリは復号されません。
      /// </remarks>
      /// <remarks>サポートされる最小クライアント: Windows XP。</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2003。</remarks>
      /// <returns>
      ///   If the function succeeds, the return value is nonzero. If the function fails, the return value is zero. To get extended error
      ///   information, call GetLastError.
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "DecryptFileW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool DecryptFile([MarshalAs(UnmanagedType.LPWStr)] string lpFileName, [MarshalAs(UnmanagedType.U4)] uint dwReserved);

      /// <summary>既存のファイルを削除します。</summary>
      /// <remarks>
      ///   存在しないファイルを削除しようとすると、DeleteFile 関数は ERROR_FILE_NOT_FOUND で失敗します。
      /// </remarks>
      /// <remarks>ファイルが読み取り専用の場合、関数は ERROR_ACCESS_DENIED で失敗します。</remarks>
      /// <remarks>
      ///   パスがシンボリックリンクを指している場合、ターゲットではなくシンボリックリンクが削除されます。ターゲットを削除するには、
      ///   CreateFile を呼び出して FILE_FLAG_DELETE_ON_CLOSE を指定する必要があります。
      /// </remarks>
      /// <remarks>サポートされる最小クライアント: Windows XP [デスクトップアプリ | Windows ストアアプリ]。</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリ | Windows ストアアプリ]。</remarks>
      /// <returns>
      ///   If the function succeeds, the return value is nonzero. If the function fails, the return value is zero (0). To get extended error
      ///   information, call GetLastError.
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "DeleteFileW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool DeleteFile([MarshalAs(UnmanagedType.LPWStr)] string lpFileName);

      /// <summary>トランザクション操作として、既存のファイルを削除します。</summary>
      /// <remarks>サポートされる最小クライアント: Windows Vista [デスクトップアプリのみ]。</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2008 [デスクトップアプリのみ]。</remarks>
      /// <returns>
      ///   If the function succeeds, the return value is nonzero. If the function fails, the return value is zero (0). To get extended error
      ///   information, call GetLastError.
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "DeleteFileTransactedW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool DeleteFileTransacted([MarshalAs(UnmanagedType.LPWStr)] string lpFileName, SafeHandle hTransaction);

      /// <summary>
      ///   ファイルまたはディレクトリを暗号化します。ファイル内のすべてのデータストリームが暗号化されます。暗号化されたディレクトリ内で
      ///   作成されたすべての新しいファイルも暗号化されます。
      /// </summary>
      /// <remarks>
      ///   EncryptFile 関数は暗号化対象のファイルへの排他的アクセスを必要とし、別のプロセスがファイルを使用している場合は失敗します。
      ///   ファイルが既に暗号化されている場合、EncryptFile は単にゼロ以外の値を返し、成功を示します。ファイルが圧縮されている場合、
      ///   EncryptFile は暗号化前にファイルを展開します。lpFileName が読み取り専用ファイルを指定した場合、関数は失敗し、
      ///   GetLastError は ERROR_FILE_READ_ONLY を返します。lpFileName が読み取り専用ファイルを含むディレクトリを指定した場合、
      ///   関数は成功しますがディレクトリは暗号化されません。
      /// </remarks>
      /// <remarks>サポートされる最小クライアント: Windows XP。</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2003。</remarks>
      /// <returns>
      ///   If the function succeeds, the return value is nonzero. If the function fails, the return value is zero. To get extended error
      ///   information, call GetLastError.
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "EncryptFileW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool EncryptFile([MarshalAs(UnmanagedType.LPWStr)] string lpFileName);

      /// <summary>
      ///   指定されたディレクトリとその中のファイルの暗号化を無効化または有効化します。指定されたディレクトリ配下の
      ///   サブディレクトリの暗号化には影響しません。
      /// </summary>
      /// <remarks>
      ///   EncryptionDisable() はディレクトリとファイルの暗号化を無効にします。FILE_ATTRIBUTE_SYSTEM 属性が設定されたファイルの
      ///   表示には影響しません。このメソッドはファイル "Desktop.ini" を作成/変更し、暗号化の値を設定します:
      ///   "Disable=0|1"。
      /// </remarks>
      /// <remarks>サポートされる最小クライアント: Windows XP Professional [デスクトップアプリのみ]。</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリのみ]。</remarks>
      /// <returns>
      ///   If the function succeeds, the return value is nonzero. If the function fails, the return value is zero. To get extended error
      ///   information, call GetLastError.
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule"), SuppressUnmanagedCodeSecurity]
      [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool EncryptionDisable([MarshalAs(UnmanagedType.LPWStr)] string dirPath, [MarshalAs(UnmanagedType.Bool)] bool disable);

      /// <summary>指定されたファイルの暗号化状態を取得します。</summary>
      /// <remarks>サポートされる最小クライアント: Windows XP Professional [デスクトップアプリのみ]。</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリのみ]。</remarks>
      /// <returns>
      ///   If the function succeeds, the return value is nonzero. If the function fails, the return value is zero. To get extended error
      ///   information, call GetLastError.
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "FileEncryptionStatusW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool FileEncryptionStatus([MarshalAs(UnmanagedType.LPWStr)] string lpFileName, out FileEncryptionStatus lpStatus);

      /// <summary>
      ///   FindFirstFile、FindFirstFileEx、FindFirstFileNameW、FindFirstFileNameTransactedW、FindFirstFileTransacted、
      ///   FindFirstStreamTransactedW、または FindFirstStreamW 関数によって開かれたファイル検索ハンドルを閉じます。
      /// </summary>
      /// <remarks>サポートされる最小クライアント: Windows XP [デスクトップアプリ | Windows ストアアプリ]。</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリ | Windows ストアアプリ]。</remarks>
      /// <returns>
      ///   If the function succeeds, the return value is nonzero. If the function fails, the return value is zero. To get extended error
      ///   information, call GetLastError.
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule"), SuppressUnmanagedCodeSecurity]
      [DllImport("kernel32.dll", SetLastError = false, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool FindClose(IntPtr hFindFile);

      /// <summary>指定された名前と属性に一致するファイルまたはサブディレクトリをディレクトリ内で検索します。</summary>
      /// <remarks>末尾のバックスラッシュは許可されておらず、削除されます。</remarks>
      /// <remarks>サポートされる最小クライアント: Windows XP [デスクトップアプリ | Windows ストアアプリ]。</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリ | Windows ストアアプリ]。</remarks>
      /// <returns>
      ///   If the function succeeds, the return value is a search handle used in a subsequent call to FindNextFile or FindClose, and the
      ///   lpFindFileData parameter contains information about the first file or directory found. If the function fails or fails to locate
      ///   files from the search string in the lpFileName parameter, the return value is INVALID_HANDLE_VALUE and the contents of
      ///   lpFindFileData are indeterminate. To get extended error information, call the GetLastError function.
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "FindFirstFileExW"), SuppressUnmanagedCodeSecurity]
      internal static extern SafeFindFileHandle FindFirstFileEx([MarshalAs(UnmanagedType.LPWStr)] string lpFileName, FINDEX_INFO_LEVELS fInfoLevelId, out WIN32_FIND_DATA lpFindFileData, FINDEX_SEARCH_OPS fSearchOp, IntPtr lpSearchFilter, FIND_FIRST_EX_FLAGS dwAdditionalFlags);

      /// <summary>
      ///   トランザクション操作として、特定の名前に一致するファイルまたはサブディレクトリをディレクトリ内で検索します。
      /// </summary>
      /// <remarks>末尾のバックスラッシュは許可されておらず、削除されます。</remarks>
      /// <remarks>サポートされる最小クライアント: Windows Vista [デスクトップアプリのみ]。</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2008 [デスクトップアプリのみ]。</remarks>
      /// <returns>
      ///   If the function succeeds, the return value is a search handle used in a subsequent call to FindNextFile or FindClose, and the
      ///   lpFindFileData parameter contains information about the first file or directory found. If the function fails or fails to locate
      ///   files from the search string in the lpFileName parameter, the return value is INVALID_HANDLE_VALUE and the contents of
      ///   lpFindFileData are indeterminate. To get extended error information, call the GetLastError function.
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "FindFirstFileTransactedW"), SuppressUnmanagedCodeSecurity]
      internal static extern SafeFindFileHandle FindFirstFileTransacted([MarshalAs(UnmanagedType.LPWStr)] string lpFileName, FINDEX_INFO_LEVELS fInfoLevelId, out WIN32_FIND_DATA lpFindFileData, FINDEX_SEARCH_OPS fSearchOp, IntPtr lpSearchFilter, FIND_FIRST_EX_FLAGS dwAdditionalFlags, SafeHandle hTransaction);

      /// <summary>
      ///   指定されたファイルへのすべてのハードリンクの列挙を作成します。FindFirstFileNameW 関数は、後続の
      ///   FindNextFileNameW 関数の呼び出しで使用できる列挙へのハンドルを返します。
      /// </summary>
      /// <remarks>サポートされる最小クライアント: Windows Vista [デスクトップアプリのみ]。</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2008 [デスクトップアプリのみ]。</remarks>
      /// <returns>
      ///   If the function succeeds, the return value is a search handle that can be used with the FindNextFileNameW function or closed with
      ///   the FindClose function. If the function fails, the return value is INVALID_HANDLE_VALUE (0xffffffff). To get extended error
      ///   information, call the GetLastError function.
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      internal static extern SafeFindFileHandle FindFirstFileNameW([MarshalAs(UnmanagedType.LPWStr)] string lpFileName, [MarshalAs(UnmanagedType.U4)] uint dwFlags, [MarshalAs(UnmanagedType.U4)] out uint stringLength, StringBuilder linkName);

      /// <summary>
      ///   トランザクション操作として、指定されたファイルへのすべてのハードリンクの列挙を作成します。この関数は、後続の
      ///   FindNextFileNameW 関数の呼び出しで使用できる列挙へのハンドルを返します。
      /// </summary>
      /// <remarks>サポートされる最小クライアント: Windows Vista [デスクトップアプリのみ]。</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2008 [デスクトップアプリのみ]。</remarks>
      /// <returns>
      ///   If the function succeeds, the return value is a search handle that can be used with the FindNextFileNameW function or closed with
      ///   the FindClose function. If the function fails, the return value is INVALID_HANDLE_VALUE (0xffffffff). To get extended error
      ///   information, call the GetLastError function.
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      internal static extern SafeFindFileHandle FindFirstFileNameTransactedW([MarshalAs(UnmanagedType.LPWStr)] string lpFileName, [MarshalAs(UnmanagedType.U4)] uint dwFlags, [MarshalAs(UnmanagedType.U4)] out uint stringLength, StringBuilder linkName, SafeHandle hTransaction);

      /// <summary>
      ///   FindFirstFile、FindFirstFileEx、または FindFirstFileTransacted 関数の前回の呼び出しからファイル検索を続行します。
      /// </summary>
      /// <remarks>サポートされる最小クライアント: Windows XP [デスクトップアプリ | Windows ストアアプリ]。</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリ | Windows ストアアプリ]。</remarks>
      /// <returns>
      ///   If the function succeeds, the return value is nonzero and the lpFindFileData parameter contains information about the next file or
      ///   directory found. If the function fails, the return value is zero and the contents of lpFindFileData are indeterminate. To get
      ///   extended error information, call the GetLastError function. If the function fails because no more matching files can be found, the
      ///   GetLastError function returns ERROR_NO_MORE_FILES.
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "FindNextFileW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool FindNextFile(SafeFindFileHandle hFindFile, out WIN32_FIND_DATA lpFindFileData);

      /// <summary>
      ///   FindFirstFileName 関数の正常な呼び出しで返されたハンドルを使用して、ファイルへのハードリンクの列挙を続行します。
      /// </summary>
      /// <remarks>サポートされる最小クライアント: Windows Vista [デスクトップアプリのみ]。</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2008 [デスクトップアプリのみ]。</remarks>
      /// <returns>
      ///   If the function succeeds, the return value is nonzero. If the function fails, the return value is zero (0). To get extended error
      ///   information, call GetLastError. If no matching files can be found, the GetLastError function returns ERROR_HANDLE_EOF.
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool FindNextFileNameW(SafeFindFileHandle hFindStream, [MarshalAs(UnmanagedType.U4)] out uint stringLength, StringBuilder linkName);

      /// <summary>指定されたファイルのバッファをフラッシュし、バッファリングされたすべてのデータをファイルに書き込ませます。</summary>
      /// <remarks>サポートされる最小クライアント: Windows XP [デスクトップアプリ | Windows ストアアプリ]。</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリ | Windows ストアアプリ]。</remarks>
      /// <returns>
      ///   If the function succeeds, the return value is nonzero. If the function fails, the return value is zero. To get extended error
      ///   information, call GetLastError.
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool FlushFileBuffers(SafeFileHandle hFile);

      /// <summary>指定されたファイルの格納に使用されるディスクストレージの実際のバイト数を取得します。</summary>
      /// <remarks>サポートされる最小クライアント: Windows XP [デスクトップアプリのみ]。</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリのみ]。</remarks>
      /// <returns>
      ///   If the function succeeds, the return value is the low-order DWORD of the actual number of bytes of disk storage used to store the
      ///   specified file, and if lpFileSizeHigh is non-NULL, the function puts the high-order DWORD of that actual value into the DWORD
      ///   pointed to by that parameter. This is the compressed file size for compressed files, the actual file size for noncompressed files.
      ///   If the function fails, and lpFileSizeHigh is NULL, the return value is INVALID_FILE_SIZE. To get extended error information, call
      ///   GetLastError. If the return value is INVALID_FILE_SIZE and lpFileSizeHigh is non-NULL, an application must call GetLastError to
      ///   determine whether the function has succeeded (value is NO_ERROR) or failed (value is other than NO_ERROR).
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "GetCompressedFileSizeW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.U4)]
      internal static extern uint GetCompressedFileSize([MarshalAs(UnmanagedType.LPWStr)] string lpFileName, [MarshalAs(UnmanagedType.U4)] out uint lpFileSizeHigh);

      /// <summary>トランザクション操作として、指定されたファイルの格納に使用されるディスクストレージの実際のバイト数を取得します。</summary>
      /// <remarks>サポートされる最小クライアント: Windows Vista [デスクトップアプリのみ]。</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2008 [デスクトップアプリのみ]。</remarks>
      /// <returns>
      ///   If the function succeeds, the return value is the low-order DWORD of the actual number of bytes of disk storage used to store the
      ///   specified file, and if lpFileSizeHigh is non-NULL, the function puts the high-order DWORD of that actual value into the DWORD
      ///   pointed to by that parameter. This is the compressed file size for compressed files, the actual file size for noncompressed files.
      ///   If the function fails, and lpFileSizeHigh is NULL, the return value is INVALID_FILE_SIZE. To get extended error information, call
      ///   GetLastError. If the return value is INVALID_FILE_SIZE and lpFileSizeHigh is non-NULL, an application must call GetLastError to
      ///   determine whether the function has succeeded (value is NO_ERROR) or failed (value is other than NO_ERROR).
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "GetCompressedFileSizeTransactedW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.U4)]
      internal static extern uint GetCompressedFileSizeTransacted([MarshalAs(UnmanagedType.LPWStr)] string lpFileName, [MarshalAs(UnmanagedType.U4)] out uint lpFileSizeHigh, SafeHandle hTransaction);

      /// <summary>
      ///   指定されたファイルまたはディレクトリの属性を取得します。
      /// </summary>
      /// <remarks>
      ///   <para>GetFileAttributes 関数はファイルシステム属性情報を取得します。</para>
      ///   <para>GetFileAttributesEx は他のファイルまたはディレクトリ属性情報セットを取得できます。</para>
      ///   <para>現在、GetFileAttributesEx はファイルシステム属性情報のスーパーセットである標準属性のセットを取得します。
      ///   マウントされたフォルダであるディレクトリに対して GetFileAttributesEx 関数が呼び出されると、マウントされたフォルダがディレクトリに関連付けるボリューム内の
      ///   ルートディレクトリの属性ではなく、ディレクトリの属性を返します。関連付けられたボリュームの属性を取得するには、
      ///   GetVolumeNameForVolumeMountPoint を呼び出して関連付けられたボリュームの名前を取得します。次に、その名前を
      ///   GetFileAttributesEx の呼び出しで使用します。結果は関連付けられたボリューム上のルートディレクトリの属性です。</para>
      ///   <para>シンボリックリンクの動作: パスがシンボリックリンクを指している場合、関数はシンボリックリンクの属性を返します。</para>
      ///   <para>サポートされる最小クライアント: Windows XP [デスクトップアプリのみ]</para>
      ///   <para>サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリのみ]</para>
      /// </remarks>
      /// <returns>
      ///   <para>If the function succeeds, the return value is nonzero.</para>
      ///   <para>If the function fails, the return value is zero. To get extended error information, call GetLastError.</para>
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "GetFileAttributesExW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool GetFileAttributesEx([MarshalAs(UnmanagedType.LPWStr)] string lpFileName, [MarshalAs(UnmanagedType.U4)] GET_FILEEX_INFO_LEVELS fInfoLevelId, out WIN32_FILE_ATTRIBUTE_DATA lpFileInformation);

      /// <summary>指定されたファイルまたはディレクトリの属性を取得します。</summary>
      /// <remarks>
      ///   <para>GetFileAttributes 関数はファイルシステム属性情報を取得します。</para>
      ///   <para>GetFileAttributesEx は他のファイルまたはディレクトリ属性情報セットを取得できます。</para>
      ///   <para>
      ///   現在、GetFileAttributesEx はファイルシステム属性情報のスーパーセットである標準属性のセットを取得します。
      ///   マウントされたフォルダであるディレクトリに対して GetFileAttributesEx 関数が呼び出されると、マウントされたフォルダがディレクトリに関連付けるボリューム内の
      ///   ルートディレクトリの属性ではなく、ディレクトリの属性を返します。関連付けられたボリュームの属性を取得するには、
      ///   GetVolumeNameForVolumeMountPoint を呼び出して関連付けられたボリュームの名前を取得します。次に、その名前を
      ///   GetFileAttributesEx の呼び出しで使用します。結果は関連付けられたボリューム上のルートディレクトリの属性です。</para>
      ///   <para>シンボリックリンクの動作: パスがシンボリックリンクを指している場合、関数はシンボリックリンクの属性を返します。</para>
      ///   <para>トランザクション操作</para>
      ///   <para>トランザクション内でファイルが変更用に開かれている場合、トランザクションがコミットされるまで他のスレッドはそのファイルを変更用に開くことができません。
      ///   逆に、トランザクション外でファイルが変更用に開かれている場合、トランザクション外のハンドルが閉じられるまでトランザクションスレッドはそのファイルを変更用に開くことができません。
      ///   トランザクション外のスレッドがファイルの変更用にハンドルを開いている場合、そのファイルに対する GetFileAttributesTransacted の呼び出しは
      ///   ERROR_TRANSACTIONAL_CONFLICT エラーで失敗します。</para>
      ///   <para>サポートされる最小クライアント: Windows Vista [デスクトップアプリのみ]</para>
      ///   <para>サポートされる最小サーバー: Windows Server 2008 [デスクトップアプリのみ]</para>
      /// </remarks>
      /// <returns>
      ///   <para>If the function succeeds, the return value is nonzero.</para>
      ///   <para>If the function fails, the return value is zero. To get extended error information, call GetLastError.</para>
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "GetFileAttributesTransactedW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool GetFileAttributesTransacted([MarshalAs(UnmanagedType.LPWStr)] string lpFileName, [MarshalAs(UnmanagedType.U4)] GET_FILEEX_INFO_LEVELS fInfoLevelId, out WIN32_FILE_ATTRIBUTE_DATA lpFileInformation, SafeHandle hTransaction);

      /// <summary>指定されたファイルのファイル情報を取得します。</summary>
      /// <returns>
      /// 関数が成功した場合、戻り値はゼロ以外で、ファイル情報データは lpByHandleFileInformation パラメータが指すバッファに格納されます。
      /// 関数が失敗した場合、戻り値はゼロです。拡張エラー情報を取得するには GetLastError を呼び出してください。
      /// </returns>
      /// <remarks>
      /// オペレーティングシステムの基盤となるネットワーク機能と接続先のサーバーの種類によっては、
      /// GetFileInformationByHandle 関数が失敗したり、部分的な情報を返したり、指定されたファイルの完全な情報を返したりする場合があります。
      /// </remarks>
      /// <remarks>サポートされる最小クライアント: Windows XP [デスクトップアプリのみ]</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリのみ]</remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool GetFileInformationByHandle(SafeFileHandle hFile, [MarshalAs(UnmanagedType.Struct)] out BY_HANDLE_FILE_INFORMATION lpByHandleFileInformation);

      /// <summary>指定されたファイルのファイル情報を取得します。</summary>
      /// <remarks>
      ///   <para>Minimum supported client: Windows Vista [desktop apps | Windows Store apps]</para>
      ///   <para>Minimum supported server: Windows Server 2008 [desktop apps | Windows Store apps]</para>
      ///   <para>Redistributable: Windows SDK on Windows Server 2003 and Windows XP.</para>
      /// </remarks>
      /// <param name="hFile">The file.</param>
      /// <param name="fileInfoByHandleClass">The file information by handle class.</param>
      /// <param name="lpFileInformation">Information describing the file.</param>
      /// <param name="dwBufferSize">Size of the buffer.</param>
      /// <returns>
      ///   <para>If the function succeeds, the return value is nonzero and file information data is contained in the buffer pointed to by the
      ///   lpByHandleFileInformation parameter.</para>
      ///   <para>If the function fails, the return value is zero. To get extended error information, call GetLastError.</para>
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool GetFileInformationByHandleEx(SafeFileHandle hFile, [MarshalAs(UnmanagedType.I4)] FILE_INFO_BY_HANDLE_CLASS fileInfoByHandleClass, SafeGlobalMemoryBufferHandle lpFileInformation, [MarshalAs(UnmanagedType.U4)] uint dwBufferSize);

      /// <summary>指定されたファイルのファイル情報を取得します。</summary>
      /// <remarks>
      ///   <para>Minimum supported client: Windows Vista [desktop apps | Windows Store apps]</para>
      ///   <para>Minimum supported server: Windows Server 2008 [desktop apps | Windows Store apps]</para>
      ///   <para>Redistributable: Windows SDK on Windows Server 2003 and Windows XP.</para>
      /// </remarks>
      /// <returns>
      ///   <para>If the function succeeds, the return value is nonzero and file information data is contained in the buffer pointed to by the
      ///   lpByHandleFileInformation parameter.</para>
      ///   <para>If the function fails, the return value is zero. To get extended error information, call GetLastError.</para>
      /// </returns>
      /// <param name="hFile">The file.</param>
      /// <param name="fileInfoByHandleClass">The file information by handle class.</param>
      /// <param name="lpFileInformation">Information describing the file.</param>
      /// <param name="dwBufferSize">Size of the buffer.</param>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "GetFileInformationByHandleEx"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool GetFileInformationByHandleEx_FileBasicInfo(SafeFileHandle hFile, [MarshalAs(UnmanagedType.I4)] FILE_INFO_BY_HANDLE_CLASS fileInfoByHandleClass, [MarshalAs(UnmanagedType.Struct)] out FILE_BASIC_INFO lpFileInformation, [MarshalAs(UnmanagedType.U4)] uint dwBufferSize);

      /// <summary>指定されたファイルのサイズを取得します。</summary>
      /// <remarks>
      ///   <para>Minimum supported client: Windows XP [desktop apps only]</para>
      ///   <para>Minimum supported server: Windows Server 2003 [desktop apps only]</para>
      /// </remarks>
      /// <returns>
      ///   <para>If the function succeeds, the return value is nonzero.</para>
      ///   <para>If the function fails, the return value is zero. To get extended error information, call GetLastError.</para>
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool GetFileSizeEx(SafeFileHandle hFile, out long lpFileSize);

      /// <summary>指定されたファイルの最終パスを取得します。</summary>
      /// <remarks>サポートされる最小クライアント: Windows Vista [デスクトップアプリのみ]。</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2008 [デスクトップアプリのみ]。</remarks>
      /// <returns>
      ///   If the function succeeds, the return value is nonzero. If the function fails, the return value is zero. To get extended error
      ///   information, call GetLastError.
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "GetFinalPathNameByHandleW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.U4)]
      internal static extern uint GetFinalPathNameByHandle(SafeFileHandle hFile, StringBuilder lpszFilePath, [MarshalAs(UnmanagedType.U4)] uint cchFilePath, FinalPathFormats dwFlags);

      /// <summary>
      ///   指定されたアドレスが指定されたプロセスのアドレス空間内のメモリマップドファイル内にあるかどうかを確認します。
      ///   該当する場合、関数はメモリマップドファイルの名前を返します。
      /// </summary>
      /// <remarks>サポートされる最小クライアント: Windows XP。</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2003。</remarks>
      /// <returns>
      ///   If the function succeeds, the return value is nonzero. If the function fails, the return value is zero. To get extended error
      ///   information, call GetLastError.
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("psapi.dll", SetLastError = false, CharSet = CharSet.Unicode, EntryPoint = "GetMappedFileNameW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool GetMappedFileName(IntPtr hProcess, SafeLocalMemoryBufferHandle lpv, StringBuilder lpFilename, [MarshalAs(UnmanagedType.U4)] uint nSize);

      /// <summary>呼び出しプロセスによる排他的アクセスのために指定されたファイルをロックします。</summary>
      /// <remarks>サポートされる最小クライアント: Windows XP。</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2003。</remarks>
      /// <returns>
      ///   If the function succeeds, the return value is nonzero (TRUE). If the function fails, the return value is zero (FALSE). To get
      ///   extended error information, call GetLastError.
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool LockFile(SafeFileHandle hFile, [MarshalAs(UnmanagedType.U4)] uint dwFileOffsetLow, [MarshalAs(UnmanagedType.U4)] uint dwFileOffsetHigh, [MarshalAs(UnmanagedType.U4)] uint nNumberOfBytesToLockLow, [MarshalAs(UnmanagedType.U4)] uint nNumberOfBytesToLockHigh);

      /// <summary>ファイルマッピングのビューを呼び出しプロセスのアドレス空間にマップします。</summary>
      /// <remarks>サポートされる最小クライアント: Windows XP。</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2003。</remarks>
      /// <returns>
      ///   If the function succeeds, the return value is the starting address of the mapped view. If the function fails, the return value is
      ///   <c>null</c>.
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = false, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      internal static extern SafeLocalMemoryBufferHandle MapViewOfFile(SafeFileHandle hFileMappingObject, [MarshalAs(UnmanagedType.U4)] uint dwDesiredAccess, [MarshalAs(UnmanagedType.U4)] uint dwFileOffsetHigh, [MarshalAs(UnmanagedType.U4)] uint dwFileOffsetLow, UIntPtr dwNumberOfBytesToMap);

      /// <summary>
      ///   ファイルまたはディレクトリ（その子要素を含む）を移動します。
      ///   <para>進捗通知を受け取るコールバック関数を指定できます。</para>
      /// </summary>
      /// <remarks>
      ///   <para>MoveFileWithProgress 関数はリンク追跡サービスと連携して動作するため、リンクソースは移動時に追跡できます。</para>
      ///   <para>サポートされる最小クライアント: Windows XP [デスクトップアプリのみ]</para>
      ///   <para>サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリのみ]</para>
      /// </remarks>
      /// <param name="lpExistingFileName">既存ファイルのファイル名。</param>
      /// <param name="lpNewFileName">新しいファイルのファイル名。</param>
      /// <param name="lpProgressRoutine">進捗ルーチン。</param>
      /// <param name="lpData">データ。</param>
      /// <param name="dwFlags">フラグ。</param>
      /// <returns>
      ///   <para>If the function succeeds, the return value is nonzero.</para>
      ///   <para>If the function fails, the return value is zero. To get extended error information, call GetLastError.</para>
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "MoveFileWithProgressW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool MoveFileWithProgress([MarshalAs(UnmanagedType.LPWStr)] string lpExistingFileName, [MarshalAs(UnmanagedType.LPWStr)] string lpNewFileName, NativeCopyMoveProgressRoutine lpProgressRoutine, IntPtr lpData, [MarshalAs(UnmanagedType.U4)] MoveOptions dwFlags);

      /// <summary>
      ///   トランザクション操作として、既存のファイルまたはディレクトリ（その子要素を含む）を移動します。
      ///   <para>進捗通知を受け取るコールバック関数を指定できます。</para>
      /// </summary>
      /// <remarks>
      ///   <para>サポートされる最小クライアント: Windows Vista [デスクトップアプリのみ]</para>
      ///   <para>サポートされる最小サーバー: Windows Server 2008 [デスクトップアプリのみ]</para>
      /// </remarks>     
      /// <returns>
      ///   <para>If the function succeeds, the return value is nonzero.</para>
      ///   <para>If the function fails, the return value is zero. To get extended error information, call GetLastError.</para>
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "MoveFileTransactedW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool MoveFileTransacted([MarshalAs(UnmanagedType.LPWStr)] string lpExistingFileName, [MarshalAs(UnmanagedType.LPWStr)] string lpNewFileName, NativeCopyMoveProgressRoutine lpProgressRoutine, IntPtr lpData, [MarshalAs(UnmanagedType.U4)] MoveOptions dwCopyFlags, SafeHandle hTransaction);

      /// <summary>CopyFileEx、MoveFileTransacted、および MoveFileWithProgress 関数で使用されるアプリケーション定義のコールバック関数。
      /// <para>コピーまたは移動操作の一部が完了したときに呼び出されます。</para>
      /// <para>LPPROGRESS_ROUTINE 型はこのコールバック関数へのポインタを定義します。</para>
      /// <para>NativeCopyMoveProgressRoutine はアプリケーション定義の関数名のプレースホルダーです。</para>
      /// </summary>
      [SuppressUnmanagedCodeSecurity]
      internal delegate CopyMoveProgressResult NativeCopyMoveProgressRoutine([MarshalAs(UnmanagedType.I8)] long totalFileSize, [MarshalAs(UnmanagedType.I8)] long totalBytesTransferred, [MarshalAs(UnmanagedType.I8)] long streamSize, [MarshalAs(UnmanagedType.I8)] long streamBytesTransferred, [MarshalAs(UnmanagedType.U4)] uint dwStreamNumber, [MarshalAs(UnmanagedType.U4)] CopyMoveProgressCallbackReason dwCallbackReason, IntPtr hSourceFile, IntPtr hDestinationFile, IntPtr lpData);

      /// <summary>あるファイルを別のファイルで置換します。元のファイルのバックアップコピーを作成するオプションがあります。置換ファイルは置換されたファイルの名前とIDを引き継ぎます。</summary>
      /// <returns>
      /// If the function succeeds, the return value is nonzero.
      /// If the function fails, the return value is zero. To get extended error information, call GetLastError.
      /// </returns>
      /// <remarks>Minimum supported client: Windows XP [desktop apps only]</remarks>
      /// <remarks>Minimum supported server: Windows Server 2003 [desktop apps only]</remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "ReplaceFileW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool ReplaceFile([MarshalAs(UnmanagedType.LPWStr)] string lpReplacedFileName, [MarshalAs(UnmanagedType.LPWStr)] string lpReplacementFileName, [MarshalAs(UnmanagedType.LPWStr)] string lpBackupFileName, FileSystemRights dwReplaceFlags, IntPtr lpExclude, IntPtr lpReserved);

      /// <summary>ファイルまたはディレクトリの属性を設定します。</summary>
      /// <returns>
      /// If the function succeeds, the return value is nonzero.
      /// If the function fails, the return value is zero. To get extended error information, call GetLastError.
      /// </returns>
      /// <remarks>Minimum supported client: Windows XP</remarks>
      /// <remarks>Minimum supported server: Windows Server 2003</remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [SuppressMessage("Microsoft.Usage", "CA2205:UseManagedEquivalentsOfWin32Api")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "SetFileAttributesW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool SetFileAttributes([MarshalAs(UnmanagedType.LPWStr)] string lpFileName, [MarshalAs(UnmanagedType.U4)] FileAttributes dwFileAttributes);

      /// <summary>トランザクション操作として、ファイルまたはディレクトリの属性を設定します。</summary>
      /// <returns>
      /// If the function succeeds, the return value is nonzero.
      /// If the function fails, the return value is zero. To get extended error information, call GetLastError.
      /// </returns>
      /// <remarks>Minimum supported client: Windows Vista [desktop apps only]</remarks>
      /// <remarks>Minimum supported server: Windows Server 2008 [desktop apps only]</remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "SetFileAttributesTransactedW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool SetFileAttributesTransacted([MarshalAs(UnmanagedType.LPWStr)] string lpFileName, [MarshalAs(UnmanagedType.U4)] FileAttributes dwFileAttributes, SafeHandle hTransaction);

      /// <summary>指定されたファイルのファイルポインタを移動します。</summary>
      /// <returns>
      /// If the function succeeds, the return value is nonzero.
      /// If the function fails, the return value is zero. To get extended error information, call GetLastError.
      /// </returns>
      /// <remarks>Minimum supported client: Windows XP [desktop apps | UWP apps]</remarks>
      /// <remarks>Minimum supported server: Windows Server 2003 [desktop apps | UWP apps]</remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool SetFilePointerEx(SafeFileHandle hFile, [MarshalAs(UnmanagedType.U8)] ulong liDistanceToMove, IntPtr lpNewFilePointer, [MarshalAs(UnmanagedType.U4)] SeekOrigin dwMoveMethod);

      /// <summary>指定されたファイルまたはディレクトリの作成日時、最終アクセス日時、または最終変更日時を設定します。</summary>
      /// <returns>
      /// If the function succeeds, the return value is nonzero.
      /// If the function fails, the return value is zero. To get extended error information, call GetLastError.
      /// </returns>
      /// <remarks>Minimum supported client: Windows XP [desktop apps only]</remarks>
      /// <remarks>Minimum supported server: Windows Server 2003 [desktop apps only]</remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool SetFileTime(SafeFileHandle hFile, SafeGlobalMemoryBufferHandle lpCreationTime, SafeGlobalMemoryBufferHandle lpLastAccessTime, SafeGlobalMemoryBufferHandle lpLastWriteTime);

      /// <summary>開いているファイル内の領域のロックを解除します。領域のロックを解除すると、他のプロセスがその領域にアクセスできるようになります。</summary>
      /// <returns>
      /// If the function succeeds, the return value is nonzero.
      /// If the function fails, the return value is zero. To get extended error information, call GetLastError.
      /// </returns>
      /// <remarks>Minimum supported client: Windows XP</remarks>
      /// <remarks>Minimum supported server: Windows Server 2003</remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool UnlockFile(SafeFileHandle hFile, [MarshalAs(UnmanagedType.U4)] uint dwFileOffsetLow, [MarshalAs(UnmanagedType.U4)] uint dwFileOffsetHigh, [MarshalAs(UnmanagedType.U4)] uint nNumberOfBytesToUnlockLow, [MarshalAs(UnmanagedType.U4)] uint nNumberOfBytesToUnlockHigh);

      /// <summary>呼び出しプロセスのアドレス空間からファイルのマップされたビューをアンマップします。</summary>
      /// <remarks>サポートされる最小クライアント: Windows XP。</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2003。</remarks>
      /// <param name="lpBaseAddress">ベースアドレス。</param>
      /// <returns>
      ///   If the function succeeds, the return value is nonzero. If the function fails, the return value is zero. To get extended error
      ///   information, call GetLastError.
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = false, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool UnmapViewOfFile(SafeLocalMemoryBufferHandle lpBaseAddress);

      
      /// <summary>指定されたファイルまたはディレクトリ内の ::$DATA ストリーム型の最初のストリームを列挙します。</summary>
      /// <returns>
      /// If the function succeeds, the return value is a search handle that can be used in subsequent calls to the <see cref="FindNextStreamW"/> function.
      /// If the function fails, the return value is INVALID_HANDLE_VALUE. To get extended error information, call GetLastError.
      /// If no streams can be found, the function fails and GetLastError returns <see cref="Win32Errors.ERROR_HANDLE_EOF"/> (38).
      /// </returns>
      /// <remarks>Minimum supported client: Windows Vista [desktop apps only]</remarks>
      /// <remarks>Minimum supported server: Windows Server 2003 [desktop apps only]</remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      internal static extern SafeFindFileHandle FindFirstStreamW(string lpFileName, STREAM_INFO_LEVELS infoLevel, SafeGlobalMemoryBufferHandle lpFindStreamData, int dwFlags);


      /// <summary>トランザクション操作として、指定されたファイルまたはディレクトリ内の最初のストリームを列挙します。</summary>
      /// <returns>
      /// If the function succeeds, the return value is a search handle that can be used in subsequent calls to the <see cref="FindNextStreamW"/> function.
      /// If the function fails, the return value is INVALID_HANDLE_VALUE. To get extended error information, call GetLastError.
      /// </returns>
      /// <remarks>Minimum supported client: Windows Vista [desktop apps only]</remarks>
      /// <remarks>Minimum supported server: Windows Server 2003 [desktop apps only]</remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      internal static extern SafeFindFileHandle FindFirstStreamTransactedW(string lpFileName, STREAM_INFO_LEVELS infoLevel, SafeGlobalMemoryBufferHandle lpFindStreamData, int dwFlags, SafeHandle hTransaction);


      /// <summary><see cref="FindFirstStreamW"/> 関数の前回の呼び出しで開始されたストリーム検索を続行します。</summary>
      /// <returns>
      /// If the function succeeds, the return value is nonzero.
      /// If the function fails, the return value is zero. To get extended error information, call GetLastError. If no more streams can be found, GetLastError returns <see cref="Win32Errors.ERROR_HANDLE_EOF"/> (38).
      /// </returns>
      /// <remarks>Minimum supported client: Windows Vista [desktop apps only]</remarks>
      /// <remarks>Minimum supported server: Windows Server 2003 [desktop apps only]</remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool FindNextStreamW(SafeFindFileHandle handle, SafeGlobalMemoryBufferHandle lpFindStreamData);



      #region Restart Manager

      private const int CCH_RM_MAX_APP_NAME = 255;
      private const int CCH_RM_MAX_SVC_NAME = 63;


      internal enum RM_APP_TYPE
      {
         RmUnknownApp = 0,
         RmMainWindow = 1,
         RmOtherWindow = 2,
         RmService = 3,
         RmExplorer = 4,
         RmConsole = 5,
         RmCritical = 1000
      }


      [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
      internal struct RM_UNIQUE_PROCESS
      {
         [MarshalAs(UnmanagedType.I4)] public readonly int dwProcessId;
         [MarshalAs(UnmanagedType.Struct)] public readonly FILETIME ProcessStartTime;
      }


      [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
      internal struct RM_PROCESS_INFO
      {
         [MarshalAs(UnmanagedType.Struct)] public RM_UNIQUE_PROCESS Process;
         [MarshalAs(UnmanagedType.ByValTStr, SizeConst = CCH_RM_MAX_APP_NAME + 1)] public readonly string strAppName;
         [MarshalAs(UnmanagedType.ByValTStr, SizeConst = CCH_RM_MAX_SVC_NAME + 1)] public readonly string strServiceShortName;
         [MarshalAs(UnmanagedType.I4)] public readonly RM_APP_TYPE ApplicationType;
         [MarshalAs(UnmanagedType.U4)] public readonly uint AppStatus;
         [MarshalAs(UnmanagedType.U4)] public readonly uint TSSessionId;
         [MarshalAs(UnmanagedType.Bool)] public readonly bool bRestartable;
      }


      /// <summary>再起動マネージャーセッションを終了します。この関数は、RmStartSession 関数を呼び出して以前にセッションを開始したプライマリインストーラーが呼び出す必要があります。</summary>
      /// <para>The RmEndSession function can be called by a secondary installer that is joined to the session once no more resources need to be registered by the secondary installer.</para>
      /// <para>&#160;</para>
      /// <returns>This is the most recent error received. The function can return one of the system error codes that are defined in Winerror.h.</returns>
      /// <para>&#160;</para>
      /// <remarks>
      /// <para>Minimum supported client: Windows Vista [desktop apps only]</para>
      /// <para>Minimum supported server: Windows Server 2008 [desktop apps only]</para>
      /// </remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("rstrtmgr.dll", SetLastError = true, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.I4)]
      internal static extern int RmEndSession([MarshalAs(UnmanagedType.U4)] uint pSessionHandle);


      /// <summary>再起動マネージャーセッションに登録されたリソースを現在使用しているすべてのアプリケーションとサービスの一覧を取得します。</summary>
      /// <para>&#160;</para>
      /// <returns>This is the most recent error received. The function can return one of the system error codes that are defined in Winerror.h.</returns>
      /// <para>&#160;</para>
      /// <remarks>
      /// <para>Minimum supported client: Windows Vista [desktop apps only]</para>
      /// <para>Minimum supported server: Windows Server 2008 [desktop apps only]</para>
      /// </remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("rstrtmgr.dll", SetLastError = true, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.I4)]
      internal static extern int RmGetList([MarshalAs(UnmanagedType.U4)] uint dwSessionHandle, [MarshalAs(UnmanagedType.U4)] out uint pnProcInfoNeeded, [MarshalAs(UnmanagedType.U4)] ref uint pnProcInfo, [MarshalAs(UnmanagedType.LPArray)] [In, Out] RM_PROCESS_INFO[] rgAffectedApps, [MarshalAs(UnmanagedType.U4)] ref uint lpdwRebootReasons);


      /// <summary>再起動マネージャーセッションにリソースを登録します。再起動マネージャーは、セッションに登録されたリソースの一覧を使用して、シャットダウンおよび再起動が必要なアプリケーションとサービスを判断します。</summary>
      /// <para>Resources can be identified by filenames, service short names, or RM_UNIQUE_PROCESS structures that describe running applications.</para>
      /// <para>The RmRegisterResources function can be used by a primary or secondary installer.</para>
      /// <para>&#160;</para>
      /// <returns>This is the most recent error received. The function can return one of the system error codes that are defined in Winerror.h.</returns>
      /// <para>&#160;</para>
      /// <remarks>
      /// <para>Minimum supported client: Windows Vista [desktop apps only]</para>
      /// <para>Minimum supported server: Windows Server 2008 [desktop apps only]</para>
      /// </remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("rstrtmgr.dll", SetLastError = true, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.I4)]
      internal static extern int RmRegisterResources([MarshalAs(UnmanagedType.U4)] uint pSessionHandle, [MarshalAs(UnmanagedType.U4)] uint nFiles, [MarshalAs(UnmanagedType.LPArray)] string[] rgsFilenames, [MarshalAs(UnmanagedType.U4)] uint nApplications, [In] RM_UNIQUE_PROCESS[] rgApplications, [MarshalAs(UnmanagedType.U4)] uint nServices, [MarshalAs(UnmanagedType.LPArray)] string[] rgsServiceNames);


      /// <summary>新しい再起動マネージャーセッションを開始します。ユーザーセッションあたり最大 64 の再起動マネージャーセッションをシステム上で同時に開くことができます。</summary>
      /// <para>When this function starts a session, it returns a session handle and session key that can be used in subsequent calls to the Restart Manager API.</para>
      /// <para>&#160;</para>
      /// <returns>This is the most recent error received. The function can return one of the system error codes that are defined in Winerror.h.</returns>
      /// <para>&#160;</para>
      /// <remarks>
      /// <para>Minimum supported client: Windows Vista [desktop apps only]</para>
      /// <para>Minimum supported server: Windows Server 2008 [desktop apps only]</para>
      /// </remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("rstrtmgr.dll", SetLastError = true, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.I4)]
      internal static extern int RmStartSession([MarshalAs(UnmanagedType.U4)] out uint pSessionHandle, [MarshalAs(UnmanagedType.I4)] int dwSessionFlags, [MarshalAs(UnmanagedType.LPWStr)] string strSessionKey);
      
      #endregion // Restart Manager
   }
}
