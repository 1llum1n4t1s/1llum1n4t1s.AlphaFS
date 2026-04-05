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
using System.Runtime.InteropServices;
using System.Security;
using System.Text;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class Volume
   {
      /// <summary>[AlphaFS] 指定されたボリューム上のすべてのマウントフォルダー（ボリュームマウントポイント）の <see cref="String"/> の列挙可能なコレクションを返します。</summary>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="ArgumentException"/>
      /// <param name="volumeGuid">ボリューム <see cref="Guid"/> を含む <see cref="string"/>。</param>
      /// <returns>指定されたボリューム上のすべてのボリュームマウントポイントの <see cref="String"/> の列挙可能なコレクション。</returns>
      [SecurityCritical]
      public static IEnumerable<string> EnumerateVolumeMountPoints(string volumeGuid)
      {
         if (Utils.IsNullOrWhiteSpace(volumeGuid))
         {
            throw new ArgumentNullException("volumeGuid");
         }

         if (!volumeGuid.StartsWith(Path.VolumePrefix + "{", StringComparison.OrdinalIgnoreCase))
         {
            throw new ArgumentException(Resources.Not_A_Valid_Guid, "volumeGuid");
         }


         // 末尾のバックスラッシュが必要です。
         volumeGuid = Path.AddTrailingDirectorySeparator(volumeGuid, false);


         var buffer = new StringBuilder(NativeMethods.MaxPathUnicode);


         using (new NativeMethods.ChangeErrorMode(NativeMethods.ErrorMode.FailCriticalErrors))
         using (var handle = NativeMethods.FindFirstVolumeMountPoint(volumeGuid, buffer, (uint)buffer.Capacity))
         {
            var lastError = Marshal.GetLastWin32Error();

            if (!NativeMethods.IsValidHandle(handle, false))
            {
               switch ((uint)lastError)
               {
                  case Win32Errors.ERROR_NO_MORE_FILES:
                  case Win32Errors.ERROR_PATH_NOT_FOUND: // Observed with USB stick, FAT32 formatted.
                     yield break;

                  default:
                     NativeError.ThrowException(lastError, volumeGuid);
                     break;
               }
            }

            yield return buffer.ToString();


            while (NativeMethods.FindNextVolumeMountPoint(handle, buffer, (uint)buffer.Capacity))
            {
               lastError = Marshal.GetLastWin32Error();

               var throwException = lastError != Win32Errors.ERROR_NO_MORE_FILES && lastError != Win32Errors.ERROR_PATH_NOT_FOUND && lastError != Win32Errors.ERROR_MORE_DATA;

               if (!NativeMethods.IsValidHandle(handle, lastError, volumeGuid, throwException))
               {
                  yield break;
               }

               yield return buffer.ToString();
            }
         }
      }
   }
}
