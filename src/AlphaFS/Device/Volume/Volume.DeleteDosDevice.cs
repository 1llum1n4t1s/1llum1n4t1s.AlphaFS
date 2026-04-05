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

using System.Diagnostics.CodeAnalysis;
using System.Security;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class Volume
   {
      /// <summary>[AlphaFS] MS-DOS デバイス名を削除します。</summary>
      /// <param name="deviceName">削除するデバイスを指定する MS-DOS デバイス名。</param>
      [SecurityCritical]
      public static void DeleteDosDevice(string deviceName)
      {
         DefineDosDeviceCore(false, deviceName, null, DosDeviceAttributes.RemoveDefinition, false);
      }

      /// <summary>[AlphaFS] MS-DOS デバイス名を削除します。</summary>
      /// <param name="deviceName">削除するデバイスを指定する MS-DOS デバイス名文字列。</param>
      /// <param name="targetPath">
      ///   このデバイスを実装するパス文字列へのポインター。<see cref="DosDeviceAttributes.RawTargetPath"/> フラグが指定されていない限り、
      ///   文字列は MS-DOS パス文字列です。指定されている場合、この文字列はパス文字列です。
      /// </param>
      [SecurityCritical]
      public static void DeleteDosDevice(string deviceName, string targetPath)
      {
         DefineDosDeviceCore(false, deviceName, targetPath, DosDeviceAttributes.RemoveDefinition, false);
      }

      /// <summary>[AlphaFS] MS-DOS デバイス名を削除します。</summary>
      /// <param name="deviceName">削除するデバイスを指定する MS-DOS デバイス名文字列。</param>
      /// <param name="targetPath">
      ///   このデバイスを実装するパス文字列へのポインター。<see cref="DosDeviceAttributes.RawTargetPath"/> フラグが指定されていない限り、
      ///   文字列は MS-DOS パス文字列です。指定されている場合、この文字列はパス文字列です。
      /// </param>
      /// <param name="exactMatch">
      ///   正確な名前一致の場合のみ MS-DOS デバイスを削除します。<paramref name="exactMatch"/> が <c>true</c> の場合、
      ///   <paramref name="targetPath"/> はマッピングの作成に使用されたパスと同じでなければなりません。
      /// </param>
      [SecurityCritical]
      public static void DeleteDosDevice(string deviceName, string targetPath, bool exactMatch)
      {
         DefineDosDeviceCore(false, deviceName, targetPath, DosDeviceAttributes.RemoveDefinition, exactMatch);
      }

      /// <summary>[AlphaFS] MS-DOS デバイス名を削除します。</summary>
      /// <param name="deviceName">削除するデバイスを指定する MS-DOS デバイス名文字列。</param>
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
      [SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes")]
      [SecurityCritical]
      public static void DeleteDosDevice(string deviceName, string targetPath, DosDeviceAttributes deviceAttributes, bool exactMatch)
      {
         DefineDosDeviceCore(false, deviceName, targetPath, deviceAttributes, exactMatch);
      }
   }
}
