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
   public static partial class Volume
   {
      /// <summary>[AlphaFS]
      ///   指定されたボリュームマウントポイント（ドライブ文字、ボリューム GUID パス、またはマウントフォルダー）に関連付けられた
      ///   ボリュームの <see cref="Guid"/> パスを取得します。
      /// </summary>
      /// <exception cref="ArgumentNullException"/>
      /// <param name="volumeMountPoint">
      ///   マウントフォルダーのパス（例: "Y:\MountX\"）またはドライブ文字（例: "X:\"）。
      /// </param>
      /// <returns>"\\?\Volume{GUID}\" 形式の一意のボリューム名。</returns>
      [SuppressMessage("Microsoft.Interoperability", "CA1404:CallGetLastErrorImmediatelyAfterPInvoke", Justification = "Marshal.GetLastWin32Error() is manipulated.")]
      [SecurityCritical]
      public static string GetVolumeGuid(string volumeMountPoint)
      {
         if (Utils.IsNullOrWhiteSpace(volumeMountPoint))
         {
            throw new ArgumentNullException("volumeMountPoint");
         }

         // 文字列は末尾のバックスラッシュ ('\') で終わる必要があります。
         volumeMountPoint = Path.GetFullPathCore(null, false, volumeMountPoint, GetFullPathOptions.AsLongPath | GetFullPathOptions.AddTrailingDirectorySeparator | GetFullPathOptions.FullCheck);

         var volumeGuid = new StringBuilder(100);
         var uniqueName = new StringBuilder(100);

         using (new NativeMethods.ChangeErrorMode(NativeMethods.ErrorMode.FailCriticalErrors))
         {
            // GetVolumeNameForVolumeMountPoint()
            // 2013-07-18: MSDN は LongPath の使用を確認していませんが、この関数の Unicode バージョンが存在します。

            if (!NativeMethods.GetVolumeNameForVolumeMountPoint(volumeMountPoint, volumeGuid, (uint)volumeGuid.Capacity))
            {
               var lastError = (uint) Marshal.GetLastWin32Error();

               if (lastError != Win32Errors.ERROR_MORE_DATA)
                  NativeError.ThrowException(lastError, volumeMountPoint);

               return null;
            }

            // 文字列は末尾のバックスラッシュで終わる必要があります。
            if (!NativeMethods.GetVolumeNameForVolumeMountPoint(Path.AddTrailingDirectorySeparator(volumeGuid.ToString(), false), uniqueName, (uint)uniqueName.Capacity))
            {
               var lastError = (uint) Marshal.GetLastWin32Error();

               if (lastError != Win32Errors.ERROR_MORE_DATA)
                  NativeError.ThrowException(lastError, volumeMountPoint);

               return null;
            }

            return uniqueName.ToString();
         }
      }
   }
}
