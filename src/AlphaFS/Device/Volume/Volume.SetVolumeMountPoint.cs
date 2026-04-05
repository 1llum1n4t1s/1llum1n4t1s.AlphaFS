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
   public static partial class Volume
   {
      /// <summary>[AlphaFS] ボリュームをドライブ文字または別のボリュームのディレクトリに関連付けます。</summary>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <param name="volumeMountPoint">
      ///   ボリュームに関連付けるユーザーモードパス。ドライブ文字（例: "X:\"）
      ///   または別のボリュームのディレクトリ（例: "Y:\MountX\"）を指定できます。
      /// </param>
      /// <param name="volumeGuid">ボリューム <see cref="Guid"/> を含む <see cref="string"/>。</param>      
      [SuppressMessage("Microsoft.Design", "CA1062:Validate arguments of public methods", MessageId = "1", Justification = "Utils.IsNullOrWhiteSpace validates arguments.")]
      [SecurityCritical]
      public static void SetVolumeMountPoint(string volumeMountPoint, string volumeGuid)
      {
         if (Utils.IsNullOrWhiteSpace(volumeMountPoint))
         {
            throw new ArgumentNullException("volumeMountPoint");
         }

         if (Utils.IsNullOrWhiteSpace(volumeGuid))
         {
            throw new ArgumentNullException("volumeGuid");
         }

         if (!volumeGuid.StartsWith(Path.VolumePrefix + "{", StringComparison.OrdinalIgnoreCase))
         {
            throw new ArgumentException(Resources.Not_A_Valid_Guid, "volumeGuid");
         }


         volumeMountPoint = Path.GetFullPathCore(null, false, volumeMountPoint, GetFullPathOptions.AsLongPath | GetFullPathOptions.AddTrailingDirectorySeparator | GetFullPathOptions.FullCheck);


         // この文字列は "\\?\Volume{GUID}\" の形式である必要があります
         volumeGuid = Path.AddTrailingDirectorySeparator(volumeGuid, false);


         // ChangeErrorMode は Win32 SetThreadErrorMode() メソッド用で、ポップアップの抑制に使用されます。
         using (new NativeMethods.ChangeErrorMode(NativeMethods.ErrorMode.FailCriticalErrors))
         {
            // SetVolumeMountPoint()
            // 2014-01-29: MSDN は LongPath の使用を確認していませんが、この関数の Unicode バージョンが存在します。

            // 文字列は末尾のバックスラッシュで終わる必要があります。
            var success = NativeMethods.SetVolumeMountPoint(volumeMountPoint, volumeGuid);

            var lastError = Marshal.GetLastWin32Error();
            if (!success)
            {
               // lpszVolumeMountPoint パラメーターにマウントフォルダーへのパスが含まれている場合、
               // ディレクトリが空であっても GetLastError は ERROR_DIR_NOT_EMPTY を返します。

               if (lastError != Win32Errors.ERROR_DIR_NOT_EMPTY)
               {
                  NativeError.ThrowException(lastError, volumeGuid);
               }
            }
         }
      }
   }
}
