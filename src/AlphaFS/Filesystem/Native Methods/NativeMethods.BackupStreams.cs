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
      /// <summary>BackupRead 関数は、セキュリティ情報を含むファイルまたはディレクトリのバックアップに使用できます。
      ///   <para>この関数は、指定されたファイルまたはディレクトリに関連付けられたデータをバッファに読み込みます。</para>
      ///   <para>その後、WriteFile 関数を使用してバックアップメディアに書き込むことができます。</para>
      /// </summary>
      /// <remarks>
      ///   <para>この関数は、暗号化ファイルシステム (EFS) で暗号化されたファイルのバックアップには使用しないでください。</para>
      ///   <para>その目的には ReadEncryptedFileRaw を使用してください。</para>
      ///   <para>サポートされる最小クライアント: Windows XP [デスクトップアプリのみ]</para>
      ///   <para>サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリのみ]</para>
      /// </remarks>
      /// <param name="hFile">ファイルハンドル。</param>
      /// <param name="lpBuffer">バッファ。</param>
      /// <param name="nNumberOfBytesToRead">読み取るバイト数。</param>
      /// <param name="lpNumberOfBytesRead">[out] 読み取られたバイト数。</param>
      /// <param name="bAbort">中断する場合は true。</param>
      /// <param name="bProcessSecurity">セキュリティを処理する場合は true。</param>
      /// <param name="lpContext">[out] コンテキスト。</param>
      /// <returns>
      ///   <para>関数が成功した場合、戻り値はゼロ以外です。</para>
      ///   <para>関数が失敗した場合、戻り値はゼロで、I/O エラーが発生したことを示します。拡張エラー情報を取得するには
      ///   GetLastError を呼び出してください。</para>
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool BackupRead(SafeFileHandle hFile, SafeGlobalMemoryBufferHandle lpBuffer, [MarshalAs(UnmanagedType.U4)] uint nNumberOfBytesToRead, [MarshalAs(UnmanagedType.U4)] out uint lpNumberOfBytesRead, [MarshalAs(UnmanagedType.Bool)] bool bAbort, [MarshalAs(UnmanagedType.Bool)] bool bProcessSecurity, ref IntPtr lpContext);


      /// <summary>BackupSeek 関数は、<see cref="BackupRead"/> または <see cref="BackupWrite"/> 関数で最初にアクセスされた
      ///   データストリーム内を前方にシークします。
      ///   <para>この関数は、指定されたファイルまたはディレクトリに関連付けられたデータをバッファに読み込み、WriteFile 関数を使用して
      ///   バックアップメディアに書き込むことができます。</para>
      /// </summary>
      /// <remarks>
      ///   <para>アプリケーションは BackupSeek 関数を使用して、エラーを引き起こすデータストリームの部分をスキップします。</para>
      ///   <para>この関数はストリームヘッダーをまたいでシークしません。例えば、この関数はストリーム名のスキップには使用できません。</para>
      ///   <para>アプリケーションがサブストリームの終端を超えてシークしようとすると、関数は失敗し、lpdwLowByteSeeked および
      ///   lpdwHighByteSeeked パラメータは</para>
      ///   <para>関数が実際にシークしたバイト数を示し、ファイル位置は次のストリームヘッダーの先頭に配置されます。</para>
      ///   <para>&#160;</para>
      ///   <para>サポートされる最小クライアント: Windows XP [デスクトップアプリのみ]</para>
      ///   <para>サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリのみ]</para>
      /// </remarks>
      /// <param name="hFile">ファイルハンドル。</param>
      /// <param name="dwLowBytesToSeek">シークする下位バイト数。</param>
      /// <param name="dwHighBytesToSeek">シークする上位バイト数。</param>
      /// <param name="lpdwLowBytesSeeked">[out] 実際にシークされた下位バイト数。</param>
      /// <param name="lpdwHighBytesSeeked">[out] 実際にシークされた上位バイト数。</param>
      /// <param name="lpContext">[out] コンテキスト。</param>
      /// <returns>
      ///   <para>関数が要求された量をシークできた場合、ゼロ以外の値を返します。</para>
      ///   <para>関数が要求された量をシークできなかった場合、ゼロを返します。拡張エラー情報を取得するには
      ///   GetLastError を呼び出してください。</para>
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool BackupSeek(SafeFileHandle hFile, [MarshalAs(UnmanagedType.U4)] uint dwLowBytesToSeek, [MarshalAs(UnmanagedType.U4)] uint dwHighBytesToSeek, [MarshalAs(UnmanagedType.U4)] out uint lpdwLowBytesSeeked, [MarshalAs(UnmanagedType.U4)] out uint lpdwHighBytesSeeked, ref IntPtr lpContext);


      /// <summary>BackupWrite 関数は、<see cref="BackupRead"/> を使用してバックアップされたファイルまたはディレクトリの復元に使用できます。
      ///   <para>ReadFile 関数を使用してバックアップメディアからデータストリームを取得し、BackupWrite を使用して
      ///   指定されたファイルまたはディレクトリにデータを書き込みます。</para>
      ///   <para>&#160;</para>
      /// </summary>
      /// <remarks>
      ///   <para>この関数は、暗号化ファイルシステム (EFS) で暗号化されたファイルの復元には使用しないでください。
      ///   その目的には WriteEncryptedFileRaw を使用してください。</para>
      ///   <para>サポートされる最小クライアント: Windows XP [デスクトップアプリのみ]</para>
      ///   <para>サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリのみ]</para>
      /// </remarks>
      /// <param name="hFile">ファイルハンドル。</param>
      /// <param name="lpBuffer">バッファ。</param>
      /// <param name="nNumberOfBytesToWrite">書き込むバイト数。</param>
      /// <param name="lpNumberOfBytesWritten">[out] 書き込まれたバイト数。</param>
      /// <param name="bAbort">中断する場合は true。</param>
      /// <param name="bProcessSecurity">セキュリティを処理する場合は true。</param>
      /// <param name="lpContext">[out] コンテキスト。</param>
      /// <returns>
      ///   <para>関数が成功した場合、戻り値はゼロ以外です。</para>
      ///   <para>関数が失敗した場合、戻り値はゼロで、I/O エラーが発生したことを示します。拡張エラー情報を取得するには
      ///   GetLastError を呼び出してください。</para>
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool BackupWrite(SafeFileHandle hFile, SafeGlobalMemoryBufferHandle lpBuffer, [MarshalAs(UnmanagedType.U4)] uint nNumberOfBytesToWrite, [MarshalAs(UnmanagedType.U4)] out uint lpNumberOfBytesWritten, [MarshalAs(UnmanagedType.Bool)] bool bAbort, [MarshalAs(UnmanagedType.Bool)] bool bProcessSecurity, ref IntPtr lpContext);
   }
}
