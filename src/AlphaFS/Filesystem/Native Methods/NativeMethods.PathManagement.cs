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
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;

namespace Alphaleonis.Win32.Filesystem
{
   internal static partial class NativeMethods
   {
      /// <summary>指定されたファイルまたはディレクトリの完全パスとファイル名を取得します。</summary>
      /// <returns>関数がその他の理由で失敗した場合、戻り値はゼロです。拡張エラー情報を取得するには GetLastError を呼び出してください。</returns>
      /// <remarks>GetFullPathName 関数はマルチスレッドアプリケーションや共有ライブラリコードには推奨されません。</remarks>
      /// <remarks>サポートされる最小クライアント: Windows XP [デスクトップアプリのみ]</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリのみ]</remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "GetFullPathNameW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.U4)]
      internal static extern uint GetFullPathName([MarshalAs(UnmanagedType.LPWStr)] string lpFileName, [MarshalAs(UnmanagedType.U4)] uint nBufferLength, StringBuilder lpBuffer, IntPtr lpFilePart);

      /// <summary>トランザクション操作として、指定されたファイルまたはディレクトリの完全パスとファイル名を取得します。</summary>
      /// <returns>関数がその他の理由で失敗した場合、戻り値はゼロです。拡張エラー情報を取得するには GetLastError を呼び出してください。</returns>
      /// <remarks>サポートされる最小クライアント: Windows Vista [デスクトップアプリのみ]</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2008 [デスクトップアプリのみ]</remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "GetFullPathNameTransactedW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.U4)]
      internal static extern uint GetFullPathNameTransacted([MarshalAs(UnmanagedType.LPWStr)] string lpFileName, [MarshalAs(UnmanagedType.U4)] uint nBufferLength, StringBuilder lpBuffer, IntPtr lpFilePart, SafeHandle hTransaction);

      /// <summary>指定されたパスをロング形式に変換します。</summary>
      /// <returns>
      /// 関数が成功した場合、戻り値はゼロ以外です。
      /// 関数が失敗した場合、戻り値はゼロです。拡張エラー情報を取得するには GetLastError を呼び出してください。
      /// </returns>
      /// <remarks>サポートされる最小クライアント: Windows XP [デスクトップアプリのみ]</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリのみ]</remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "GetLongPathNameW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.U4)]
      internal static extern uint GetLongPathName([MarshalAs(UnmanagedType.LPWStr)] string lpszShortPath, StringBuilder lpszLongPath, [MarshalAs(UnmanagedType.U4)] uint cchBuffer);

      /// <summary>トランザクション操作として、指定されたパスをロング形式に変換します。</summary>
      /// <returns>
      /// 関数が成功した場合、戻り値はゼロ以外です。
      /// 関数が失敗した場合、戻り値はゼロです。拡張エラー情報を取得するには GetLastError を呼び出してください。
      /// </returns>
      /// <remarks>サポートされる最小クライアント: Windows Vista [デスクトップアプリのみ]</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2008 [デスクトップアプリのみ]</remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "GetLongPathNameTransactedW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.U4)]
      internal static extern uint GetLongPathNameTransacted([MarshalAs(UnmanagedType.LPWStr)] string lpszShortPath, StringBuilder lpszLongPath, [MarshalAs(UnmanagedType.U4)] uint cchBuffer, SafeHandle hTransaction);

      /// <summary>指定されたパスのショートパス形式を取得します。</summary>
      /// <returns>
      /// 関数が成功した場合、戻り値はゼロ以外です。
      /// 関数が失敗した場合、戻り値はゼロです。拡張エラー情報を取得するには GetLastError を呼び出してください。
      /// </returns>
      /// <remarks>サポートされる最小クライアント: Windows XP</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2003</remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "GetShortPathNameW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.U4)]
      internal static extern uint GetShortPathName([MarshalAs(UnmanagedType.LPWStr)] string lpszLongPath, StringBuilder lpszShortPath, [MarshalAs(UnmanagedType.U4)] uint cchBuffer);
   }
}
