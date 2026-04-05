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
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Security;

namespace Alphaleonis.Win32.Filesystem
{
   public sealed partial class DirectoryInfo
   {
      #region .NET

      /// <summary>ディレクトリ内のすべてのファイルとサブディレクトリを表す、厳密に型指定された <see cref="FileSystemInfo"/> エントリの配列を返します。</summary>
      /// <returns>厳密に型指定された <see cref="FileSystemInfo"/> エントリの配列。</returns>
      /// <remarks>
      /// サブディレクトリの場合、このメソッドによって返される <see cref="FileSystemInfo"/> オブジェクトは、派生クラス <see cref="DirectoryInfo"/> にキャストできます。
      /// <see cref="FileSystemInfo.Attributes"/> プロパティによって返される <see cref="FileAttributes"/> 値を使用して、<see cref="FileSystemInfo"/> がファイルかディレクトリかを判断します。
      /// </remarks>
      /// <remarks>
      /// DirectoryInfo にファイルもディレクトリもない場合、このメソッドは空の配列を返します。このメソッドは再帰的ではありません。
      /// サブディレクトリの場合、このメソッドによって返される FileSystemInfo オブジェクトは、派生クラス DirectoryInfo にキャストできます。
      /// Attributes プロパテ���によって返される FileAttributes 値を使用して、FileSystemInfo がファイルかディレクトリかを判断します。
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Infos")]
      [SecurityCritical]
      public FileSystemInfo[] GetFileSystemInfos()
      {
         return Directory.EnumerateFileSystemEntryInfosCore<FileSystemInfo>(null, Transaction, LongFullName, Path.WildcardStarMatchAll, null, null, null, PathFormat.LongFullPath).ToArray();
      }


      /// <summary>指定された検索条件に一致するファイルおよびサブディレクトリを表す、厳密に型指定された <see cref="FileSystemInfo"/> オブジェクトの配列を取得します。</summary>
      /// <param name="searchPattern">
      ///   パス内のディレクトリ名と照合する検索文字列。
      ///   このパラメーターには、有効なリテラルパスとワイルドカード
      ///   （<see cref="Path.WildcardStarMatchAll"/> および <see cref="Path.WildcardQuestion"/>）文字の組み合わせを含めることができますが、正規表現はサポートされません。
      /// </param>
      /// <returns>厳密に型指定された <see cref="FileSystemInfo"/> エントリの配列。</returns>
      /// <remarks>
      /// サブディレクトリの場合、このメソッドによって返される <see cref="FileSystemInfo"/> オブジェクトは、派生クラス <see cref="DirectoryInfo"/> にキャストできます。
      /// <see cref="FileSystemInfo.Attributes"/> プロパティによって返される <see cref="FileAttributes"/> 値を使用して、<see cref="FileSystemInfo"/> がファイルかディレクトリかを判断します。
      /// </remarks>
      /// <remarks>
      /// DirectoryInfo にファイルもディレクトリもない場合、このメソッドは空の配列を返します。このメソッドは再帰的ではありません。
      /// サブディレクトリの場合、このメソッドによって返される FileSystemInfo オブジェクトは、派生クラス DirectoryInfo にキャストできます。
      /// Attributes プロパティによって返される FileAttributes 値を使用して、FileSystemInfo がファイルかディレクトリかを判断します。
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Infos")]
      [SecurityCritical]
      public FileSystemInfo[] GetFileSystemInfos(string searchPattern)
      {
         return Directory.EnumerateFileSystemEntryInfosCore<FileSystemInfo>(null, Transaction, LongFullName, searchPattern, null, null, null, PathFormat.LongFullPath).ToArray();
      }


      /// <summary>指定された検索条件に一致するファイルおよびサブディレクトリを表す、厳密に型指定された <see cref="FileSystemInfo"/> オブジェクトの配列を取得します。</summary>
      /// <param name="searchPattern">
      ///   パス内のディレクトリ名と照合する検索文字列。
      ///   このパラメーターには、有効なリテラルパスとワイルドカード
      ///   （<see cref="Path.WildcardStarMatchAll"/> および <see cref="Path.WildcardQuestion"/>）文字の組み合わせを含めることができますが、正規表現はサポートされません。
      /// </param>
      /// <param name="searchOption">
      ///   <paramref name="searchOption"/> が現在のディレクトリのみを含むか、すべてのサブディレクトリを含むかを指定する
      ///   <see cref="SearchOption"/> 列挙値の 1 つ。
      /// </param>
      /// <returns>厳密に型指定された <see cref="FileSystemInfo"/> エントリの配列。</returns>
      /// <remarks>
      /// サブディレクトリの場合、このメソッドによって返される <see cref="FileSystemInfo"/> オブジェクトは、派生クラス <see cref="DirectoryInfo"/> にキャストできます。
      /// <see cref="FileSystemInfo.Attributes"/> プロパティによって返される <see cref="FileAttributes"/> 値を使用して、<see cref="FileSystemInfo"/> がファイルかディレクトリかを判断します。
      /// </remarks>
      /// <remarks>
      /// DirectoryInfo にファイルもディレクトリもない場合、このメソッドは空の配列を返します。このメソッドは再帰的ではありません。
      /// サブディレクトリの場合、このメソッドによって返される FileSystemInfo オブジェクトは、派生クラス DirectoryInfo にキャストできます。
      /// Attributes プロパティによって返される FileAttributes 値を使用して、FileSystemInfo がファイルかディレクトリかを判断します。
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Infos")]
      [SecurityCritical]
      public FileSystemInfo[] GetFileSystemInfos(string searchPattern, SearchOption searchOption)
      {
         return Directory.EnumerateFileSystemEntryInfosCore<FileSystemInfo>(null, Transaction, LongFullName, searchPattern, searchOption, null, null, PathFormat.LongFullPath).ToArray();
      }

      #endregion // .NET
   }
}
