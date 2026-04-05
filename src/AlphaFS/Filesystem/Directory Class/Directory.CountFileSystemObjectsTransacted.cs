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
using System.Linq;
using System.Security;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class Directory
   {
      /// <summary>[AlphaFS] 指定されたディレクトリ内のファイルシステムオブジェクト（ファイル、フォルダ、またはその両方）をカウントします。</summary>
      /// <returns>カウントされたファイルシステムオブジェクトの数。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">The directory path.</param>
      /// <param name="options">ディレクトリの列挙方法を指定する <see cref="DirectoryEnumerationOptions"/> フラグ。</param>
      [SecurityCritical]
      public static long CountFileSystemObjectsTransacted(KernelTransaction transaction, string path, DirectoryEnumerationOptions options)
      {
         return EnumerateFileSystemEntryInfosCore<string>(null, transaction, path, Path.WildcardStarMatchAll, null, options, null, PathFormat.RelativePath).Count();
      }


      /// <summary>[AlphaFS] 指定されたディレクトリ内のファイルシステムオブジェクト（ファイル、フォルダ、またはその両方）をカウントします。</summary>
      /// <returns>カウントされたファイルシステムオブジェクトの数。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">The directory path.</param>
      /// <param name="options">ディレクトリの列挙方法を指定する <see cref="DirectoryEnumerationOptions"/> フラグ。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static long CountFileSystemObjectsTransacted(KernelTransaction transaction, string path, DirectoryEnumerationOptions options, PathFormat pathFormat)
      {
         return EnumerateFileSystemEntryInfosCore<string>(null, transaction, path, Path.WildcardStarMatchAll, null, options, null, pathFormat).Count();
      }
      

      /// <summary>[AlphaFS] 指定されたディレクトリ内のファイルシステムオブジェクト（ファイル、フォルダ、またはその両方）をカウントします。</summary>
      /// <returns>カウントされたファイルシステムオブジェクトの数。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">The directory path.</param>
      /// <param name="searchPattern">
      ///   ディレクトリ名と照合する検索文字列。対象: <paramref name="path"/>.
      ///   このパラメータには、有効なリテラルパスとワイルドカードの組み合わせを含めることができますが、
      ///   (<see cref="Path.WildcardStarMatchAll"/> and <see cref="Path.WildcardQuestion"/>) characters, but does not support regular expressions.
      /// </param>
      /// <param name="options">ディレクトリの列挙方法を指定する <see cref="DirectoryEnumerationOptions"/> フラグ。</param>
      [SecurityCritical]
      public static long CountFileSystemObjectsTransacted(KernelTransaction transaction, string path, string searchPattern, DirectoryEnumerationOptions options)
      {
         return EnumerateFileSystemEntryInfosCore<string>(null, transaction, path, searchPattern, null, options, null, PathFormat.RelativePath).Count();
      }


      /// <summary>[AlphaFS] 指定されたディレクトリ内のファイルシステムオブジェクト（ファイル、フォルダ、またはその両方）をカウントします。</summary>
      /// <returns>カウントされたファイルシステムオブジェクトの数。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">The directory path.</param>
      /// <param name="searchPattern">
      ///   ディレクトリ名と照合する検索文字列。対象: <paramref name="path"/>.
      ///   このパラメータには、有効なリテラルパスとワイルドカードの組み合わせを含めることができますが、
      ///   (<see cref="Path.WildcardStarMatchAll"/> and <see cref="Path.WildcardQuestion"/>) characters, but does not support regular expressions.
      /// </param>
      /// <param name="options">ディレクトリの列挙方法を指定する <see cref="DirectoryEnumerationOptions"/> フラグ。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static long CountFileSystemObjectsTransacted(KernelTransaction transaction, string path, string searchPattern, DirectoryEnumerationOptions options, PathFormat pathFormat)
      {
         return EnumerateFileSystemEntryInfosCore<string>(null, transaction, path, searchPattern, null, options, null, pathFormat).Count();
      }
   }
}
