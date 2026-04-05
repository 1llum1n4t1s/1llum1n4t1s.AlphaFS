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
using System.Runtime.InteropServices;
using System.Security;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class Volume
   {
      /// <summary>[AlphaFS] ドライブ文字またはマウントフォルダーを削除します。</summary>
      /// <remarks>マウントフォルダーの削除は、基になるディレクトリの削除を引き起こしません。</remarks>
      /// <remarks>
      ///   <paramref name="volumeMountPoint"/> パラメーターがマウントフォルダーではないディレクトリの場合、関数は何もしません。
      ///   ディレクトリは削除されません。
      /// </remarks>
      /// <remarks>
      ///   ボリュームマウントポイントに実際にボリュームがマウントされていない場合に、ボリュームマウントポイントからボリュームの
      ///   マウント解除を試みてもエラーにはなりません。
      /// </remarks>
      /// <param name="volumeMountPoint">削除するドライブ文字またはマウントフォルダー。例: X:\ または Y:\MountX\。</param>      
      [SecurityCritical]
      public static void DeleteVolumeMountPoint(string volumeMountPoint)
      {
         DeleteVolumeMountPointCore(null, volumeMountPoint, false, false, PathFormat.RelativePath);
      }




      /// <summary>ドライブ文字またはマウントフォルダーを削除します。
      /// <remarks>
      ///   <para>ボリュームマウントポイントに実際にボリュームがマウントされていない場合に、マウント解除を試みてもエラーにはなりません。</para>
      ///   <para>マウントフォルダーの削除は、基になるディレクトリの削除を引き起こしません。</para>
      /// </remarks>
      /// </summary>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="NotSupportedException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="volumeMountPoint">削除するドライブ文字またはマウントフォルダー。例: X:\ または Y:\MountX\。</param>
      /// <param name="continueOnException"><c>true</c> はリソース不足などの失敗から発生する可能性のある例外を抑制します。</param>
      /// <param name="continueIfJunction"><c>true</c> はこのマウントポイントがジャンクションである場合の例外を抑制します。</param>
      /// <param name="pathFormat">パスパラメーターの形式を示します。</param>
      [SecurityCritical]
      internal static void DeleteVolumeMountPointCore(KernelTransaction transaction, string volumeMountPoint, bool continueOnException, bool continueIfJunction, PathFormat pathFormat)
      {
         if (pathFormat != PathFormat.LongFullPath)
         {
            Path.CheckSupportedPathFormat(volumeMountPoint, true, true);
         }

         volumeMountPoint = Path.GetExtendedLengthPathCore(transaction, volumeMountPoint, pathFormat, GetFullPathOptions.RemoveTrailingDirectorySeparator);


         using (new NativeMethods.ChangeErrorMode(NativeMethods.ErrorMode.FailCriticalErrors))
         {
            // DeleteVolumeMountPoint()
            // 2013-01-13: MSDN は LongPath の使用を確認していませんが、この関数の Unicode バージョンが存在します。

            // 末尾のバックスラッシュが必要です。
            var success = NativeMethods.DeleteVolumeMountPoint(Path.AddTrailingDirectorySeparator(volumeMountPoint, false));

            var lastError = Marshal.GetLastWin32Error();
            if (!success && !continueOnException)
            {
               if (lastError == Win32Errors.ERROR_INVALID_PARAMETER && continueIfJunction)
               {
                  return;
               }

               if (lastError == Win32Errors.ERROR_FILE_NOT_FOUND)
               {
                  lastError = (int)Win32Errors.ERROR_PATH_NOT_FOUND;
               }

               NativeError.ThrowException(lastError, volumeMountPoint);
            }
         }
      }
   }
}
