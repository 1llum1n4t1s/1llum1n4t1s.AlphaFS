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

namespace Alphaleonis.Win32.Filesystem
{
   public sealed partial class DirectoryInfo
   {
      #region .NET

      /// <summary>現在のディレクトリ内のファイル情報の列挙可能なコレクションを返します。</summary>
      /// <returns>現在のディレクトリ内のファイルの列挙可能なコレクション。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      [SecurityCritical]
      public IEnumerable<FileInfo> EnumerateFiles()
      {
         return Directory.EnumerateFileSystemEntryInfosCore<FileInfo>(false, Transaction, LongFullName, Path.WildcardStarMatchAll, null, null, null, PathFormat.LongFullPath);
      }


      /// <summary>検索パターンに一致するファイル情報の列挙可能なコレクションを返します。</summary>
      /// <returns><paramref name="searchPattern"/> に一致するファイルの列挙可能なコレクション。</returns>
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
      [SecurityCritical]
      public IEnumerable<FileInfo> EnumerateFiles(string searchPattern)
      {
         return Directory.EnumerateFileSystemEntryInfosCore<FileInfo>(false, Transaction, LongFullName, searchPattern, null, null, null, PathFormat.LongFullPath);
      }


      /// <summary>指定された検索パターンおよびサブディレクトリ検索オプションに一致するファイル情報の列挙可能なコレクションを返します。</summary>
      /// <returns><paramref name="searchPattern"/> および <paramref name="searchOption"/> に一致するファイルの列挙可能なコレクション。</returns>
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
      /// <param name="searchOption">
      ///   <paramref name="searchOption"/> が現在のディレクトリのみを含むか、すべてのサブディレクトリを含むかを指定する
      ///   <see cref="SearchOption"/> 列挙値の 1 つ。
      /// </param>
      [SecurityCritical]
      public IEnumerable<FileInfo> EnumerateFiles(string searchPattern, SearchOption searchOption)
      {
         return Directory.EnumerateFileSystemEntryInfosCore<FileInfo>(false, Transaction, LongFullName, searchPattern, searchOption, null, null, PathFormat.LongFullPath);
      }

      #endregion // .NET


      /// <summary>[AlphaFS] 現在のディレクトリ内のファイル情報の列挙可能なコレクションを返します。</summary>
      /// <returns>現在のディレクトリ内のファイルの列挙可能なコレクション。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="options">ディレクトリの列挙方法を指定する <see cref="DirectoryEnumerationOptions"/> フラグ。</param>
      [SecurityCritical]
      public IEnumerable<FileInfo> EnumerateFiles(DirectoryEnumerationOptions options)
      {
         return Directory.EnumerateFileSystemEntryInfosCore<FileInfo>(false, Transaction, LongFullName, Path.WildcardStarMatchAll, null, options, null, PathFormat.LongFullPath);
      }


      /// <summary>[AlphaFS] 検索パターンに一致するファイル情報の列挙可能なコレクションを返します。</summary>
      /// <returns><paramref name="searchPattern"/> に一致するファイルの列挙可能なコレクション。</returns>
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
      public IEnumerable<FileInfo> EnumerateFiles(string searchPattern, DirectoryEnumerationOptions options)
      {
         return Directory.EnumerateFileSystemEntryInfosCore<FileInfo>(false, Transaction, LongFullName, searchPattern, null, options, null, PathFormat.LongFullPath);
      }
   }
}
