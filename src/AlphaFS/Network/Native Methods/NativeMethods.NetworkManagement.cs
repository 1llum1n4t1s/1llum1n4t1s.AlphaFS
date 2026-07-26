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

namespace Alphaleonis.Win32.Network
{
   internal static partial class NativeMethods
   {
      /// <summary>NetApiBufferFree 関数は、ネットワーク管理関数が確保したメモリを解放します。</summary>
      /// <returns>成功した場合は NERR_Success。無効なポインタを渡した場合は ERROR_INVALID_PARAMETER。</returns>
      /// <remarks>
      /// <para>netapi32 の各列挙 / 情報取得関数が返すバッファは、この関数で解放しなければなりません。</para>
      /// <para>サポートされる最小クライアント: Windows 2000 Professional [desktop apps only]</para>
      /// <para>サポートされる最小サーバー: Windows 2000 Server [desktop apps only]</para>
      /// </remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("netapi32.dll", SetLastError = true, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.U4)]
      internal static extern uint NetApiBufferFree(IntPtr buffer);


      /// <summary>NetServerDiskEnum 関数はサーバー上のディスクドライブのリストを取得します.</summary>
      /// <returns>
      /// If the function succeeds, the return value is NERR_Success.
      /// If the function fails, the return value is a system error code.
      /// </returns>
      /// <remarks>
      /// <para>The function returns an array of three-character strings (a drive letter, a colon, and a terminating null character).</para>
      /// <para>Only members of the Administrators or Server Operators local group can successfully execute the NetServerDiskEnum function on a remote computer.</para>
      /// <para>サポートされる最小クライアント: Windows 2000 Professional [desktop apps only]</para>
      /// <para>サポートされる最小サーバー: Windows 2000 Server [desktop apps only]</para>
      /// </remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("netapi32.dll", SetLastError = true, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.U4)]
      internal static extern uint NetServerDiskEnum([MarshalAs(UnmanagedType.LPWStr)] string serverName, [MarshalAs(UnmanagedType.U4)] uint level, out SafeNetApiBufferHandle bufPtr, [MarshalAs(UnmanagedType.I4)] int prefMaxLen, [MarshalAs(UnmanagedType.U4)] out uint entriesRead, [MarshalAs(UnmanagedType.U4)] out uint totalEntries, [MarshalAs(UnmanagedType.U4)] ref uint resumeHandle);
   }
}
