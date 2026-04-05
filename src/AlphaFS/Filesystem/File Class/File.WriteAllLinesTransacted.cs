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
using System.IO;
using System.Security;
using System.Text;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class File
   {
      #region .NET

      /// <summary>[AlphaFS] 新しいファイルを作成し、文字列のコレクションをファイルに書き込み、ファイルを閉じます。</summary>
      /// <remarks>このメソッドのデフォルト動作は、BOM(バイトオーダーマーク)なしのUTF-8エンコーディングを使用してデータを書き出すことです。</remarks>
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
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">書き込むファイル。</param>
      /// <param name="contents">ファイルに書き込む行。</param>
      [SecurityCritical]
      public static void WriteAllLinesTransacted(KernelTransaction transaction, string path, IEnumerable<string> contents)
      {
         WriteAppendAllLinesCore(transaction, path, contents, new UTF8Encoding(false, true), false, true, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] 指定されたエンコーディングを使用して新しいファイルを作成し、文字列のコレクションをファイルに書き込み、ファイルを閉じます。</summary>
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
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">書き込むファイル。</param>
      /// <param name="contents">ファイルに書き込む行。</param>
      /// <param name="encoding">使用する文字<see cref="Encoding"/>。</param>
      [SecurityCritical]
      public static void WriteAllLinesTransacted(KernelTransaction transaction, string path, IEnumerable<string> contents, Encoding encoding)
      {
         WriteAppendAllLinesCore(transaction, path, contents, encoding, false, true, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] 新しいファイルを作成し、文字列のコレクションをファイルに書き込み、ファイルを閉じます。</summary>
      /// <remarks>このメソッドのデフォルト動作は、BOM(バイトオーダーマーク)なしのUTF-8エンコーディングを使用してデータを書き出すことです。</remarks>
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
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">書き込むファイル。</param>
      /// <param name="contents">ファイルに書き込む文字列配列。</param>
      [SecurityCritical]
      public static void WriteAllLinesTransacted(KernelTransaction transaction, string path, string[] contents)
      {
         WriteAppendAllLinesCore(transaction, path, contents, new UTF8Encoding(false, true), false, true, PathFormat.RelativePath);
      }
      

      /// <summary>[AlphaFS] 指定されたエンコーディングを使用して新しいファイルを作成し、文字列のコレクションをファイルに書き込み、ファイルを閉じます。</summary>
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
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">書き込むファイル。</param>
      /// <param name="contents">ファイルに書き込む文字列配列。</param>
      /// <param name="encoding">使用する文字<see cref="Encoding"/>。</param>
      [SecurityCritical]
      public static void WriteAllLinesTransacted(KernelTransaction transaction, string path, string[] contents, Encoding encoding)
      {
         WriteAppendAllLinesCore(transaction, path, contents, encoding, false, true, PathFormat.RelativePath);
      }

      #endregion // .NET


      /// <summary>[AlphaFS] 新しいファイルを作成し、文字列のコレクションをファイルに書き込み、ファイルを閉じます。</summary>
      /// <remarks>このメソッドのデフォルト動作は、BOM(バイトオーダーマーク)なしのUTF-8エンコーディングを使用してデータを書き出すことです。</remarks>
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
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">書き込むファイル。</param>
      /// <param name="contents">ファイルに書き込む行。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static void WriteAllLinesTransacted(KernelTransaction transaction, string path, IEnumerable<string> contents, PathFormat pathFormat)
      {
         WriteAppendAllLinesCore(transaction, path, contents, new UTF8Encoding(false, true), false, true, pathFormat);
      }


      /// <summary>[AlphaFS] 指定されたエンコーディングを使用して新しいファイルを作成し、文字列のコレクションをファイルに書き込み、ファイルを閉じます。</summary>
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
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">書き込むファイル。</param>
      /// <param name="contents">ファイルに書き込む行。</param>
      /// <param name="encoding">使用する文字<see cref="Encoding"/>。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static void WriteAllLinesTransacted(KernelTransaction transaction, string path, IEnumerable<string> contents, Encoding encoding, PathFormat pathFormat)
      {
         WriteAppendAllLinesCore(transaction, path, contents, encoding, false, true, pathFormat);
      }


      /// <summary>[AlphaFS] 指定されたエンコーディングを使用して新しいファイルを作成し、文字列のコレクションをファイルに書き込み、ファイルを閉じます。</summary>
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
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">書き込むファイル。</param>
      /// <param name="contents">ファイルに書き込む文字列配列。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static void WriteAllLinesTransacted(KernelTransaction transaction, string path, string[] contents, PathFormat pathFormat)
      {
         WriteAppendAllLinesCore(transaction, path, contents, new UTF8Encoding(false, true), false, true, pathFormat);
      }
      

      /// <summary>[AlphaFS] 指定されたエンコーディングを使用して新しいファイルを作成し、文字列のコレクションをファイルに書き込み、ファイルを閉じます。</summary>
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
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">書き込むファイル。</param>
      /// <param name="contents">ファイルに書き込む文字列配列。</param>
      /// <param name="encoding">使用する文字<see cref="Encoding"/>。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static void WriteAllLinesTransacted(KernelTransaction transaction, string path, string[] contents, Encoding encoding, PathFormat pathFormat)
      {
         WriteAppendAllLinesCore(transaction, path, contents, encoding, false, true, pathFormat);
      }
   }
}
