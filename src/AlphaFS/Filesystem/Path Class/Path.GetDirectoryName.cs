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
using System.Security;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class Path
   {
      #region .NET

      /// <summary>指定されたパス文字列のディレクトリ情報を返します。</summary>
      /// <returns>
      ///   <para><paramref name="path"/> のディレクトリ情報。<paramref name="path"/> がルートディレクトリを示す場合、または
      ///   <c>null</c> の場合は <c>null</c>。</para>
      ///   <para><paramref name="path"/> にディレクトリ情報が含まれていない場合は <see cref="string.Empty"/> を返します。</para>
      /// </returns>
      /// <exception cref="ArgumentException"/>
      /// <param name="path">ファイルまたはディレクトリのパス。</param>
      [SecurityCritical]
      public static string GetDirectoryName(string path)
      {
         return GetDirectoryNameCore(path, true);
      }

      #endregion // .NET


      /// <summary>[AlphaFS] 指定されたパス文字列のディレクトリ情報を返します。</summary>
      /// <returns>
      ///   <paramref name="path"/> のディレクトリ情報。<paramref name="path"/> がルートディレクトリを示す場合、または
      ///   <c>null</c> の場合は <c>null</c>。<paramref name="path"/> にディレクトリ情報が含まれていない場合は <see cref="string.Empty"/> を返します。
      /// </returns>
      /// <exception cref="ArgumentException"/>
      /// <param name="path">ファイルまたはディレクトリのパス。</param>
      /// <param name="checkInvalidPathChars"><c>true</c> の場合、<paramref name="path"/> の無効なパス文字をチェックします。</param>
      [SecurityCritical]
      public static string GetDirectoryName(string path, bool checkInvalidPathChars)
      {
         return GetDirectoryNameCore(path, checkInvalidPathChars);
      }
   }
}
