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
      /// <summary>[AlphaFS] MS-DOS デバイス名を定義、再定義、または削除します。</summary>
      /// <param name="deviceName">定義、再定義、または削除するデバイスを指定する MS-DOS デバイス名文字列。</param>
      /// <param name="targetPath">このデバイスを実装する MS-DOS パス。</param>
      [SecurityCritical]
      public static void DefineDosDevice(string deviceName, string targetPath)
      {
         DefineDosDeviceCore(true, deviceName, targetPath, DosDeviceAttributes.None, false);
      }

      /// <summary>[AlphaFS] MS-DOS デバイス名を定義、再定義、または削除します。</summary>
      /// <param name="deviceName">
      ///   定義、再定義、または削除するデバイスを指定する MS-DOS デバイス名文字列。
      /// </param>
      /// <param name="targetPath">
      ///   このデバイスを実装する MS-DOS パス。<paramref name="deviceAttributes"/> パラメーターに
      ///   <see cref="DosDeviceAttributes.RawTargetPath"/> フラグが指定されている場合、<paramref name="targetPath"/> はそのまま使用されます。
      /// </param>
      /// <param name="deviceAttributes">
      ///   DefineDosDevice 関数の制御可能な側面。デフォルトと組み合わされる <see cref="DosDeviceAttributes"/> フラグ。
      /// </param>      
      [SecurityCritical]
      public static void DefineDosDevice(string deviceName, string targetPath, DosDeviceAttributes deviceAttributes)
      {
         DefineDosDeviceCore(true, deviceName, targetPath, deviceAttributes, false);
      }




      /// <summary>MS-DOS デバイス名を定義、再定義、または削除します。</summary>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <param name="isDefine">
      ///   <c>true</c> は新しい MS-DOS デバイスを定義します。<c>false</c> は以前に定義された MS-DOS デバイスを削除します。
      /// </param>
      /// <param name="deviceName">
      ///   定義、再定義、または削除するデバイスを指定する MS-DOS デバイス名文字列。
      /// </param>
      /// <param name="targetPath">
      ///   このデバイスを実装するパス文字列へのポインター。<see cref="DosDeviceAttributes.RawTargetPath"/> フラグが指定されていない限り、
      ///   文字列は MS-DOS パス文字列です。指定されている場合、この文字列はパス文字列です。
      /// </param>
      /// <param name="deviceAttributes">
      ///   DefineDosDevice 関数の制御可能な側面。デフォルトと組み合わされる <see cref="DosDeviceAttributes"/> フラグ。
      /// </param>
      /// <param name="exactMatch">
      ///   正確な名前一致の場合のみ MS-DOS デバイスを削除します。<paramref name="exactMatch"/> が <c>true</c> の場合、
      ///   <paramref name="targetPath"/> はマッピングの作成に使用されたパスと同じでなければなりません。
      /// </param>
      [SecurityCritical]
      internal static void DefineDosDeviceCore(bool isDefine, string deviceName, string targetPath, DosDeviceAttributes deviceAttributes, bool exactMatch)
      {
         if (Utils.IsNullOrWhiteSpace(deviceName))
         {
            throw new ArgumentNullException("deviceName");
         }

         if (isDefine)
         {
            // targetPath は null が許可されています。

            // いかなる場合も末尾のバックスラッシュ ("\") は許可されません。
            deviceName = Path.GetRegularPathCore(deviceName, GetFullPathOptions.RemoveTrailingDirectorySeparator | GetFullPathOptions.CheckInvalidPathChars, false);

            using (new NativeMethods.ChangeErrorMode(NativeMethods.ErrorMode.FailCriticalErrors))
            {
               var success = NativeMethods.DefineDosDevice(deviceAttributes, deviceName, targetPath);

               var lastError = Marshal.GetLastWin32Error();
               if (!success)
               {
                  NativeError.ThrowException(lastError, deviceName, targetPath);
               }
            }
         }

         else
         {
            // このデバイスを実装するパス文字列へのポインター。
            // DDD_RAW_TARGET_PATH フラグが指定されていない限り、文字列は MS-DOS パス文字列です。

            if (exactMatch && !Utils.IsNullOrWhiteSpace(targetPath))
            {
               deviceAttributes = deviceAttributes | DosDeviceAttributes.ExactMatchOnRemove | DosDeviceAttributes.RawTargetPath;
            }

            // MS-DOS デバイス名を削除します。まず、シンボリックリンクから Windows NT デバイスの名前を取得し、
            // 次に名前空間からシンボリックリンクを削除します。

            DefineDosDevice(deviceName, targetPath, deviceAttributes);
         }
      }
   }
}
