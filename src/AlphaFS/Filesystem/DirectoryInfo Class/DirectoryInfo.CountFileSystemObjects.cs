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
   public sealed partial class DirectoryInfo
   {
      /// <summary>[AlphaFS] 指定されたディレクトリ内のファイルシステムオブジェクト（ファイル、フォルダー、またはその両方）をカウントします。</summary>
      /// <returns>カウントされたファイルシステムオブジェクトの数。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="options">ディレクトリの列挙方法を指定する <see cref="DirectoryEnumerationOptions"/> フラグ。</param>
      [SecurityCritical]
      public long CountFileSystemObjects(DirectoryEnumerationOptions options)
      {
         return Directory.EnumerateFileSystemEntryInfosCore<string>(null, Transaction, LongFullName, Path.WildcardStarMatchAll, null, options, null, PathFormat.LongFullPath).Count();
      }


      /// <summary>[AlphaFS] 指定されたディレクトリ内のファイルシステムオブジェクト（ファイル、フォルダー、またはその両方）をカウントします。</summary>
      /// <returns>カウントされたファイルシステムオブジェクトの数。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="searchPattern">
      ///   パス内のディレクトリ名と照合する検索文字列。
      ///   このパラメーターには、有効なリテラルパスとワイルドカード
      ///   （<see cref="Path.WildcardStarMatchAll"/> および <see cref="Path.WildcardQuestion"/>）文字の組み合わせを含めることができますが、正規表現はサポートされません。
      /// </param>
      /// <param name="options">ディレクトリの列挙方法を指定する <see cref="DirectoryEnumerationOptions"/> フラグ。</param>
      [SecurityCritical]
      public long CountFileSystemObjects(string searchPattern, DirectoryEnumerationOptions options)
      {
         return Directory.EnumerateFileSystemEntryInfosCore<string>(null, Transaction, LongFullName, searchPattern, null, options, null, PathFormat.LongFullPath).Count();
      }


      /// <summary>[AlphaFS] 指定されたディレクトリ内のファイルシステムオブジェクト（ファイル、フォルダー、またはその両方）をカウントします。</summary>
      /// <returns>カウントされたファイルシステムオブジェクトの数。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="searchPattern">
      ///   パス内のディレクトリ名と照合する検索文字列。
      ///   このパラメーターには、有効なリテラルパスとワイルドカード
      ///   （<see cref="Path.WildcardStarMatchAll"/> および <see cref="Path.WildcardQuestion"/>）文字の組み合わせを含めることができますが、正規表現はサポートされません。
      /// </param>
      /// <param name="options">ディレクトリの列挙方法を指定する <see cref="DirectoryEnumerationOptions"/> フラグ。</param>
      /// <param name="filters">処理で使用するカスタムフィルターの仕様。</param>
      [SecurityCritical]
      public long CountFileSystemObjects(string searchPattern, DirectoryEnumerationOptions options, DirectoryEnumerationFilters filters)
      {
         return Directory.EnumerateFileSystemEntryInfosCore<string>(null, Transaction, LongFullName, searchPattern, null, options, filters, PathFormat.LongFullPath).Count();
      }
   }
}
