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
using System.IO;
using System.Security;
using System.Text;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class File
   {
      #region .NET

      /// <summary>新しいファイルを作成し、指定された文字列をファイルに書き込み、ファイルを閉じます。対象ファイルが既に存在する場合は上書きされます。</summary>
      /// <remarks>このメソッドはBOM(バイトオーダーマーク)なしのUTF-8エンコーディングを使用します。</remarks>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="ArgumentOutOfRangeException"/>
      /// <exception cref="FileNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="SecurityException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <exception cref="PlatformNotSupportedException">The operating system is older than Windows Vista.</exception>
      /// <param name="path">書き込むファイル。</param>
      /// <param name="contents">ファイルに書き込む文字列。</param>
      [SecurityCritical]
      public static void WriteAllText(string path, string contents)
      {
         WriteAppendAllLinesCore(null, path, new[] {contents}, new UTF8Encoding(false, true), false, false, PathFormat.RelativePath);
      }


      /// <summary>新しいファイルを作成し、指定されたエンコーディングを使用して指定された文字列をファイルに書き込み、ファイルを閉じます。対象ファイルが既に存在する場合は上書きされます。</summary>
      /// <param name="path">書き込むファイル。</param>
      /// <param name="contents">ファイルに書き込む文字列。</param>
      /// <param name="encoding">ファイルの内容に適用される<see cref="Encoding"/>。</param>
      [SecurityCritical]
      public static void WriteAllText(string path, string contents, Encoding encoding)
      {
         WriteAppendAllLinesCore(null, path, new[] {contents}, encoding, false, false, PathFormat.RelativePath);
      }

      #endregion // .NET


      /// <summary>[AlphaFS] 新しいファイルを作成し、指定された文字列をファイルに書き込み、ファイルを閉じます。対象ファイルが既に存在する場合は上書きされます。</summary>
      /// <remarks>このメソッドはBOM(バイトオーダーマーク)なしのUTF-8エンコーディングを使用します。</remarks>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="ArgumentOutOfRangeException"/>
      /// <exception cref="FileNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="SecurityException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <exception cref="PlatformNotSupportedException">The operating system is older than Windows Vista.</exception>
      /// <param name="path">書き込むファイル。</param>
      /// <param name="contents">ファイルに書き込む文字列。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static void WriteAllText(string path, string contents, PathFormat pathFormat)
      {
         WriteAppendAllLinesCore(null, path, new[] {contents}, new UTF8Encoding(false, true), false, false, pathFormat);
      }


      /// <summary>[AlphaFS] 新しいファイルを作成し、指定されたエンコーディングを使用して指定された文字列をファイルに書き込み、ファイルを閉じます。対象ファイルが既に存在する場合は上書きされます。</summary>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="ArgumentOutOfRangeException"/>
      /// <exception cref="FileNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="SecurityException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <exception cref="PlatformNotSupportedException">The operating system is older than Windows Vista.</exception>
      /// <param name="path">書き込むファイル。</param>
      /// <param name="contents">ファイルに書き込む文字列。</param>
      /// <param name="encoding">ファイルの内容に適用される<see cref="Encoding"/>。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static void WriteAllText(string path, string contents, Encoding encoding, PathFormat pathFormat)
      {
         WriteAppendAllLinesCore(null, path, new[] {contents}, encoding, false, false, pathFormat);
      }
   }
}
