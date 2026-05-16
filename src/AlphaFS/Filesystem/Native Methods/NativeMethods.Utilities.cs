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
using System;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Security;

namespace Alphaleonis.Win32.Filesystem
{
   internal static partial class NativeMethods
   {
      internal static uint GetHighOrderDword(long highPart)
      {
         return (uint) ((highPart >> 32) & 0xFFFFFFFF);
      }


      internal static uint GetLowOrderDword(long lowPart)
      {
         return (uint) (lowPart & 0xFFFFFFFF);
      }


      internal static long LuidToLong(LUID luid)
      {
         var high = (ulong) luid.HighPart << 32;
         var low = (ulong) luid.LowPart & 0x00000000FFFFFFFF;

         return unchecked((long) (high | low));
      }


      internal static LUID LongToLuid(long lluid)
      {
         return new LUID {HighPart = (uint) (lluid >> 32), LowPart = (uint) (lluid & 0xFFFFFFFF)};
      }


      internal static long ToLong(uint highPart, uint lowPart)
      {
         return ((long) highPart << 32) | ((long) lowPart & 0xFFFFFFFF);
      }


      /// <summary>現在のハンドルが null でなく、閉じられておらず、無効でないことを確認します。</summary>
      /// <param name="handle">確認する現在のハンドル。</param>
      /// <param name="throwException"><c>true</c> の場合 <exception cref="Resources.Handle_Is_Invalid"/> をスローし、<c>false</c> の場合はこの例外を発生させません。</param>
      /// <returns>成功した場合は <c>true</c>、それ以外は <c>false</c>。</returns>
      /// <exception cref="ArgumentException"/>
      internal static bool IsValidHandle(SafeHandle handle, bool throwException = true)
      {
         if (null == handle || handle.IsClosed || handle.IsInvalid)
         {
            CloseSafeHandle(handle);

            if (throwException)
            {
               throw new ArgumentException(Resources.Handle_Is_Invalid, "handle");
            }

            return false;
         }

         return true;
      }


      /// <summary>現在のハンドルが null でなく、閉じられておらず、無効でないことを確認します。</summary>
      /// <param name="handle">確認する現在のハンドル。</param>
      /// <param name="lastError">Marshal.GetLastWin32Error() の結果。</param>
      /// <param name="throwException"><c>true</c> の場合 <exception cref="Resources.Handle_Is_Invalid_Win32Error"/> をスローし、<c>false</c> の場合はこの例外を発生させません。</param>
      /// <returns>成功した場合は <c>true</c>、それ以外は <c>false</c>。</returns>
      /// <exception cref="ArgumentException"/>
      internal static bool IsValidHandle(SafeHandle handle, int lastError, bool throwException = true)
      {
         if (null == handle || handle.IsClosed || handle.IsInvalid)
         {
            CloseSafeHandle(handle);

            if (throwException)
            {
               throw new ArgumentException(string.Format(CultureInfo.InvariantCulture, Resources.Handle_Is_Invalid_Win32Error, lastError), "handle");
            }

            return false;
         }

         return true;
      }


      /// <summary>現在のハンドルが null でなく、閉じられておらず、無効でないことを確認します。</summary>
      /// <param name="handle">確認する現在のハンドル。</param>
      /// <param name="lastError">Marshal.GetLastWin32Error() の結果。</param>
      /// <param name="path">例外が発生したパス。</param>
      /// <param name="throwException"><c>true</c> の場合 <exception cref="Resources.Handle_Is_Invalid_Win32Error"/> をスローし、<c>false</c> の場合はこの例外を発生させません。</param>
      /// <returns>成功した場合は <c>true</c>、それ以外は <c>false</c>。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="Exception"/>
      internal static bool IsValidHandle(SafeHandle handle, int lastError, string path, bool throwException = true)
      {
         if (null == handle || handle.IsClosed || handle.IsInvalid)
         {
            CloseSafeHandle(handle);

            if (throwException)
            {
               NativeError.ThrowException(lastError, path);
            }

            return false;
         }

         return true;
      }


      /// <summary>現在のハンドルが null でなく、閉じられておらず、無効でないことを確認します。</summary>
      /// <param name="handle">確認する現在のハンドル。</param>
      /// <param name="lastError">Marshal.GetLastWin32Error() の結果。</param>
      /// <param name="isFolder"><c>true</c> の場合ソースがディレクトリ、<c>false</c> の場合ファイル、<c>null</c> の場合物理デバイスを示します。</param>
      /// <param name="path">例外が発生したパス。</param>
      /// <param name="throwException"><c>true</c> の場合 <exception cref="Resources.Handle_Is_Invalid_Win32Error"/> をスローし、<c>false</c> の場合はこの例外を発生させません。</param>
      /// <returns>成功した場合は <c>true</c>、それ以外は <c>false</c>。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="Exception"/>
      internal static bool CloseHandleAndPossiblyThrowException(SafeHandle handle, int lastError, bool? isFolder, string path, bool throwException = true)
      {
         if (null == handle || handle.IsClosed || handle.IsInvalid)
         {
            CloseSafeHandle(handle);

            if (throwException)
            {
               NativeError.ThrowException(lastError, isFolder, path);
            }

            return false;
         }

         return true;
      }


      internal static void CloseSafeHandle(SafeHandle handle)
      {
         if (null != handle && !handle.IsClosed)
         {
            handle.Close();
         }
         // 値渡しのため `handle = null` は呼び出し元に伝播しない死コードだった。
         // SafeHandle.Close() は冪等なので、呼び出し元での null 化は不要。
      }


      /// <summary>指定された種類の重大なエラーをシステムが処理するか、プロセスが処理するかを制御します。</summary>
      /// <remarks>
      ///   エラーモードはプロセス全体に設定されるため、マルチスレッドアプリケーションが異なるエラーモード属性を設定しないようにする必要があります。
      ///   そうしないと、一貫性のないエラー処理が発生する可能性があります。
      /// </remarks>
      /// <remarks>サポートされる最小クライアント: Windows XP [デスクトップアプリのみ]。</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリのみ]。</remarks>
      /// <param name="uMode">モード。</param>
      /// <returns>戻り値はエラーモードビット属性の以前の状態です。</returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = false, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.U4)]
      private static extern ErrorMode SetErrorMode(ErrorMode uMode);


      /// <summary>指定された種類の重大なエラーをシステムが処理するか、呼び出しスレッドが処理するかを制御します。</summary>
      /// <remarks>
      ///   エラーモードはプロセス全体に設定されるため、マルチスレッドアプリケーションが異なるエラーモード属性を設定しないようにする必要があります。
      ///   そうしないと、一貫性のないエラー処理が発生する可能性があります。
      /// </remarks>
      /// <remarks>サポートされる最小クライアント: Windows 7 [デスクトップアプリのみ]。</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2008 R2 [デスクトップアプリのみ]。</remarks>
      /// <param name="dwNewMode">新しいモード。</param>
      /// <param name="lpOldMode">[out] 以前のモード。</param>
      /// <returns>戻り値はエラーモードビット属性の以前の状態です。</returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = false, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      private static extern bool SetThreadErrorMode(ErrorMode dwNewMode, [MarshalAs(UnmanagedType.U4)] out ErrorMode lpOldMode);
   }
}
