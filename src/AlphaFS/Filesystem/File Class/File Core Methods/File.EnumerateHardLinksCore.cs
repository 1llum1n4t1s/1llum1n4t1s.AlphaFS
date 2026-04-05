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
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class File
   {
      /// <summary>[AlphaFS] 指定された<paramref name="path"/>へのすべてのハードリンクの列挙を作成します。</summary>
      /// <returns>指定された<paramref name="path"/>へのすべてのハードリンクの<see cref="string"/>の列挙可能なコレクション</returns>
      /// <exception cref="PlatformNotSupportedException">The operating system is older than Windows Vista.</exception>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">ファイルの名前。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      internal static IEnumerable<string> EnumerateHardLinksCore(KernelTransaction transaction, string path, PathFormat pathFormat)
      {
         if (!NativeMethods.IsAtLeastWindowsVista)
         {
            throw new PlatformNotSupportedException(new Win32Exception((int) Win32Errors.ERROR_OLD_WIN_VERSION).Message);
         }

         var pathLp = Path.GetExtendedLengthPathCore(transaction, path, pathFormat, GetFullPathOptions.RemoveTrailingDirectorySeparator | GetFullPathOptions.FullCheck);

         // デフォルトのバッファ長。必要に応じて拡張されますが、通常は発生しません。
         uint length = NativeMethods.MaxPathUnicode;
         var builder = new StringBuilder((int) length);


      getFindFirstFileName:

      using (var safeHandle = null == transaction

                // FindFirstFileNameW() / FindFirstFileNameTransactedW() / FindNextFileNameW()
                // 2013-01-13: MSDNはLongPathの使用を確認していませんが、この関数のUnicodeバージョンが存在します。
                // 2017-05-30: FindFirstFileNameW() MSDNはLongPathの使用を確認: Windows 10 バージョン1607以降
                ? NativeMethods.FindFirstFileNameW(pathLp, 0, out length, builder) : NativeMethods.FindFirstFileNameTransactedW(pathLp, 0, out length, builder, transaction.SafeHandle))
      {
         var lastError = Marshal.GetLastWin32Error();

         if (!NativeMethods.IsValidHandle(safeHandle, false))
         {
            switch ((uint)lastError)
            {
               case Win32Errors.ERROR_MORE_DATA:
                  builder = new StringBuilder((int)length);
                  goto getFindFirstFileName;

               default:
                  // 関数が失敗した場合、戻り値はINVALID_HANDLE_VALUEです。
                  NativeError.ThrowException(lastError, pathLp);
                  break;
            }
         }

         yield return builder.ToString();


         do
         {
            while (!NativeMethods.FindNextFileNameW(safeHandle, out length, builder))
            {
               lastError = Marshal.GetLastWin32Error();

               switch ((uint)lastError)
               {
                  // 列挙の終端に到達しました。
                  case Win32Errors.ERROR_HANDLE_EOF:
                     yield break;

                  case Win32Errors.ERROR_MORE_DATA:
                     builder = new StringBuilder((int)length);
                     continue;

                  default:
                     // 関数が失敗した場合、戻り値はゼロ(0)です。
                     NativeError.ThrowException(lastError);
                     break;
               }
            }


            yield return builder.ToString();

         } while (true);
      }
      }
   }
}
