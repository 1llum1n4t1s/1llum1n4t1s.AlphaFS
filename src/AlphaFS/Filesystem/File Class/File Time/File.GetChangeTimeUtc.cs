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

using Microsoft.Win32.SafeHandles;
using System;
using System.Security;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class File
   {
      /// <summary>[AlphaFS] 指定されたファイルの変更日時を協定世界時(UTC)形式で取得します。</summary>
      /// <returns>指定されたファイルの変更日時に設定された<see cref="DateTime"/>構造体。この値はUTC時刻で表されます。</returns>
      /// <exception cref="PlatformNotSupportedException">The operating system is older than Windows Vista.</exception>
      /// <param name="path">変更日時情報を協定世界時(UTC)形式で取得するファイル。</param>
      [SecurityCritical]
      public static DateTime GetChangeTimeUtc(string path)
      {
         return GetChangeTimeCore(null, null, false, path, true, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] 指定されたファイルの変更日時を協定世界時(UTC)形式で取得します。</summary>
      /// <returns>指定されたファイルの変更日時に設定された<see cref="DateTime"/>構造体。この値はUTC時刻で表されます。</returns>
      /// <exception cref="PlatformNotSupportedException">The operating system is older than Windows Vista.</exception>
      /// <param name="path">変更日時情報を協定世界時(UTC)形式で取得するファイル。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static DateTime GetChangeTimeUtc(string path, PathFormat pathFormat)
      {
         return GetChangeTimeCore(null, null, false, path, true, pathFormat);
      }


      /// <summary>[AlphaFS] 指定されたファイルの変更日時を協定世界時(UTC)形式で取得します。</summary>
      /// <returns>指定されたファイルの変更日時に設定された<see cref="DateTime"/>構造体。この値はUTC時刻で表されます。</returns>
      /// <exception cref="PlatformNotSupportedException">The operating system is older than Windows Vista.</exception>
      /// <param name="safeFileHandle">情報を取得するファイルまたはディレクトリへのオープンハンドル。</param>
      [SecurityCritical]
      public static DateTime GetChangeTimeUtc(SafeFileHandle safeFileHandle)
      {
         return GetChangeTimeCore(null, safeFileHandle, false, null, true, PathFormat.LongFullPath);
      }
   }
}
