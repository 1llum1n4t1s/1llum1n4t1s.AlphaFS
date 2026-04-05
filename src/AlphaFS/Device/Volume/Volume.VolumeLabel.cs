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
      /// <summary>[AlphaFS] 現在のディレクトリのルートであるファイルシステムボリュームのラベルを削除します。</summary>
      [SecurityCritical]
      public static void DeleteCurrentVolumeLabel()
      {
         SetVolumeLabel(null, null);
      }


      /// <summary>[AlphaFS] ファイルシステムボリュームのラベルを削除します。</summary>
      /// <exception cref="ArgumentNullException"/>
      /// <param name="rootPathName">ファイルシステムボリュームのルートディレクトリ。この関数がラベルを削除するボリュームです。</param>
      [SecurityCritical]
      public static void DeleteVolumeLabel(string rootPathName)
      {
         if (Utils.IsNullOrWhiteSpace(rootPathName))
         {
            throw new ArgumentNullException("rootPathName");
         }


         SetVolumeLabel(rootPathName, null);
      }


      /// <summary>[AlphaFS] ファイルシステムボリュームのラベルを取得します。</summary>
      /// <param name="volumePath">
      ///   ボリュームへのパス。例: "C:\"、"\\server\share"、または "\\?\Volume{c0580d5e-2ad6-11dc-9924-806e6f6e6963}\"。
      /// </param>
      /// <returns>ファイルシステムボリュームのラベル。ボリュームラベルは一般的に必須ではないため、この関数は <c>string.Empty</c> を返すことがあります。</returns>
      [SecurityCritical]
      public static string GetVolumeLabel(string volumePath)
      {
         return new VolumeInfo(volumePath, true, true).Name;
      }


      /// <summary>[AlphaFS] 現在のディレクトリのルートであるファイルシステムボリュームのラベルを設定します。</summary>
      /// <exception cref="ArgumentNullException"/>
      /// <param name="volumeName">ボリュームの名前。</param>
      [SecurityCritical]
      public static void SetCurrentVolumeLabel(string volumeName)
      {
         if (Utils.IsNullOrWhiteSpace(volumeName))
         {
            throw new ArgumentNullException("volumeName");
         }


         var success = NativeMethods.SetVolumeLabel(null, volumeName);

         var lastError = Marshal.GetLastWin32Error();
         if (!success)
         {
            NativeError.ThrowException(lastError, volumeName);
         }
      }


      /// <summary>[AlphaFS] ファイルシステムボリュームのラベルを設定します。</summary>
      /// <param name="volumePath">
      ///   <para>ボリュームへのパス。例: "C:\"、"\\server\share"、または "\\?\Volume{c0580d5e-2ad6-11dc-9924-806e6f6e6963}\"</para>
      ///   <para>このパラメーターが <c>null</c> の場合、関数は現在のドライブを使用します。</para>
      /// </param>
      /// <param name="volumeName">
      ///   <para>ボリュームの名前。</para>
      ///   <para>このパラメーターが <c>null</c> の場合、関数は指定されたボリュームから既存のラベルを削除し、</para>
      ///   <para>新しいラベルを割り当てません。</para>
      /// </param>
      [SecurityCritical]
      public static void SetVolumeLabel(string volumePath, string volumeName)
      {
         // rootPathName == null は許可されています。現在のドライブを意味します。

         // ボリュームラベルの設定は、ローカルリソースを指す論理ドライブにのみ適用されます。
         //if (!Path.IsLocalPath(rootPathName))
         //return false;

         volumePath = Path.AddTrailingDirectorySeparator(volumePath, false);

         // NTFS は Windows Server 2003 以降、ボリュームラベルに 32 文字の制限を使用しています。
         using (new NativeMethods.ChangeErrorMode(NativeMethods.ErrorMode.FailCriticalErrors))
         {
            var success = NativeMethods.SetVolumeLabel(volumePath, volumeName);

            var lastError = Marshal.GetLastWin32Error();
            if (!success)
            {
               NativeError.ThrowException(lastError, volumePath, volumeName);
            }
         }
      }
   }
}
