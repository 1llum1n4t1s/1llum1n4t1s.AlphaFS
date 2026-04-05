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
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Security;

namespace Alphaleonis.Win32.Filesystem
{
   public sealed partial class DirectoryInfo
   {
      #region .NET

      /// <summary>現在のディレクトリ内のファイルシステム情報の列挙可能なコレクションを返します。</summary>
      /// <returns>現在のディレクトリ内のファイルシステム情報の列挙可能なコレクション。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Infos")]
      [SecurityCritical]
      public IEnumerable<FileSystemInfo> EnumerateFileSystemInfos()
      {
         return Directory.EnumerateFileSystemEntryInfosCore<FileSystemInfo>(null, Transaction, LongFullName, Path.WildcardStarMatchAll, null, null, null, PathFormat.LongFullPath);
      }


      /// <summary>指定された検索パターンに一致するファイルシステム情報の列挙可能なコレクションを返します。</summary>
      /// <returns><paramref name="searchPattern"/> に一致するファイルシステム情報オブジェクトの列挙可能なコレクション。</returns>
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
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Infos")]
      [SecurityCritical]
      public IEnumerable<FileSystemInfo> EnumerateFileSystemInfos(string searchPattern)
      {
         return Directory.EnumerateFileSystemEntryInfosCore<FileSystemInfo>(null, Transaction, LongFullName, searchPattern, null, null, null, PathFormat.LongFullPath);
      }


      /// <summary>指定された検索パターンおよびサブディレクトリ検索オプションに一致するファイルシステム情報の列挙可能なコレクションを返します。</summary>
      /// <returns><paramref name="searchPattern"/> および <paramref name="searchOption"/> に一致するファイルシステム情報オブジェクトの列挙可能なコレクション。</returns>
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
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Infos")]
      [SecurityCritical]
      public IEnumerable<FileSystemInfo> EnumerateFileSystemInfos(string searchPattern, SearchOption searchOption)
      {
         return Directory.EnumerateFileSystemEntryInfosCore<FileSystemInfo>(null, Transaction, LongFullName, searchPattern, searchOption, null, null, PathFormat.LongFullPath);
      }

      #endregion // .NET

      
      /// <summary>[AlphaFS] 現在のディレクトリ内のファイルシステム情報の列挙可能なコレクションを返します。</summary>
      /// <returns>現在のディレクトリ内のファイルシステム情報の列挙可能なコレクション。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="options">ディレクトリの列挙方法を指定する <see cref="DirectoryEnumerationOptions"/> フラグ。</param>
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Infos")]
      [SecurityCritical]
      public IEnumerable<FileSystemInfo> EnumerateFileSystemInfos(DirectoryEnumerationOptions options)
      {
         return Directory.EnumerateFileSystemEntryInfosCore<FileSystemInfo>(null, Transaction, LongFullName, Path.WildcardStarMatchAll, null, options,  null, PathFormat.LongFullPath);
      }


      /// <summary>[AlphaFS] 指定された検索パターンに一致するファイルシステム情報の列挙可能なコレクションを返します。</summary>
      /// <returns><paramref name="searchPattern"/> に一致するファイルシステム情報オブジェクトの列挙可能なコレクション。</returns>
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
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Infos")]
      [SecurityCritical]
      public IEnumerable<FileSystemInfo> EnumerateFileSystemInfos(string searchPattern, DirectoryEnumerationOptions options)
      {
         return Directory.EnumerateFileSystemEntryInfosCore<FileSystemInfo>(null, Transaction, LongFullName, searchPattern, null, options,  null, PathFormat.LongFullPath);
      }
   }
}
