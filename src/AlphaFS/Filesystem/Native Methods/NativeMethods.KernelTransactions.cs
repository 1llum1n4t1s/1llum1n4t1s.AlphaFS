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

namespace Alphaleonis.Win32.Filesystem
{
   internal static partial class NativeMethods
   {
      /// <summary>
      ///   新しいトランザクションオブジェクトを作成します。
      /// </summary>
      /// <remarks>
      ///   <para>トランザクションハンドルを閉じるには <see cref="CloseHandle"/> 関数を使用します。クライアントがトランザクションハンドルで
      ///   CommitTransaction 関数を呼び出す前に最後のトランザクションハンドルが閉じられると、KTM はトランザクションをロールバックします。</para>
      ///   <para>サポートされる最小クライアント: Windows Vista</para>
      ///   <para>サポートされる最小サーバー: Windows Server 2008</para>
      /// </remarks>
      /// <returns>
      ///   <para>関数が成功した場合、戻り値はトランザクションへのハンドルです。</para>
      ///   <para>関数が失敗した場合、戻り値は INVALID_HANDLE_VALUE です。拡張エラー情報を取得するには GetLastError 関数を呼び出してください。</para>
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("ktmw32.dll", SetLastError = true, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      internal static extern SafeKernelTransactionHandle CreateTransaction([MarshalAs(UnmanagedType.LPStruct)] Security.NativeMethods.SecurityAttributes lpTransactionAttributes, IntPtr uow, [MarshalAs(UnmanagedType.U4)] uint createOptions, [MarshalAs(UnmanagedType.U4)] uint isolationLevel, [MarshalAs(UnmanagedType.U4)] uint isolationFlags, [MarshalAs(UnmanagedType.U4)] int timeout, [MarshalAs(UnmanagedType.LPWStr)] string description);

      /// <summary>指定されたトランザクションのコミットを要求します。</summary>
      /// <remarks>
      ///   <para>TRANSACTION_COMMIT 権限で開かれたまたは作成された任意のトランザクションハンドルをコミットできます。
      ///   作成者だけでなく、任意のアプリケーションがトランザクションをコミットできます。</para>
      ///   <para>この関数は、トランザクションがまだアクティブで、準備済み、事前準備済み、またはロールバック済みでない場合にのみ呼び出すことができます。</para>
      ///   <para>サポートされる最小クライアント: Windows Vista</para>
      ///   <para>サポートされる最小サーバー: Windows Server 2008</para>
      /// </remarks>
      /// <returns>
      ///   <para>関数が成功した場合、戻り値はゼロ以外です。</para>
      ///   <para>関数が失敗した場合、戻り値は 0 (ゼロ) です。拡張エラー情報を取得するには GetLastError 関数を呼び出してください。</para>
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("ktmw32.dll", SetLastError = true, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool CommitTransaction(SafeHandle hTrans);

      /// <summary>
      ///   指定されたトランザクションのロールバックを要求します。この関数は同期的です。
      /// </summary>
      /// <returns>
      ///   <para>関数が成功した場合、戻り値はゼロ以外です。</para>
      ///   <para>関数が失敗した場合、戻り値はゼロです。拡張エラー情報を取得するには GetLastError 関数を呼び出してください。</para>
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("ktmw32.dll", SetLastError = true, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool RollbackTransaction(SafeHandle hTrans);
   }
}
