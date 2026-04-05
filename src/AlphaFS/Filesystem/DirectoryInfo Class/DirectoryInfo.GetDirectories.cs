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
      #region .NET

      /// <summary>現在のディレクトリのサブディレクトリを返します。</summary>
      /// <returns><see cref="DirectoryInfo"/> オブジェクトの配列。</returns>
      /// <remarks>サブディレクトリがない場合、このメソッドは空の配列を返します。このメソッドは再帰的ではありません。</remarks>
      /// <remarks>
      /// EnumerateDirectories メソッドと GetDirectories メソッドの違い: EnumerateDirectories を使用すると、コレクション全体が返される前に名前のコレクションの列挙を開始できます。
      /// GetDirectories を使用する場合は、配列にアクセスする前に名前の配列全体が返されるのを待つ必要があります。
      /// そのため、多数のファイルやディレクトリを操作する場合は、EnumerateDirectories の方が効率的です。
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      [SecurityCritical]
      public DirectoryInfo[] GetDirectories()
      {
         return Directory.EnumerateFileSystemEntryInfosCore<DirectoryInfo>(true, Transaction, LongFullName, Path.WildcardStarMatchAll, null, null, null, PathFormat.LongFullPath).ToArray();
      }


      /// <summary>指定された検索条件に一致する現在の <see cref="DirectoryInfo"/> 内のディレクトリの配列を返します。</summary>
      /// <returns><paramref name="searchPattern"/> に一致する <see cref="DirectoryInfo"/> 型の配列。</returns>
      /// <remarks>
      /// EnumerateDirectories メソッドと GetDirectories メソッドの違い: EnumerateDirectories を使用すると、コレクション全体が返される前に名前のコレクションの列挙を開始できます。
      /// GetDirectories を使用する場合は、配列にアクセスする前に名前の配列全体が返されるのを待つ必要があります。
      /// そのため、多数のファイルやディレクトリを操作する場合は、EnumerateDirectories の方が効率的です。
      /// </remarks>
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
      public DirectoryInfo[] GetDirectories(string searchPattern)
      {
         return Directory.EnumerateFileSystemEntryInfosCore<DirectoryInfo>(true, Transaction, LongFullName, searchPattern, null, null, null, PathFormat.LongFullPath).ToArray();
      }


      /// <summary>指定された検索条件に一致し、サブディレクトリを検索するかどうかを決定する値を使用して、現在の <see cref="DirectoryInfo"/> 内のディレクトリの配列を返します。</summary>
      /// <returns><paramref name="searchPattern"/> に一致する <see cref="DirectoryInfo"/> 型の配列。</returns>
      /// <remarks>サブディレクトリがない場合、または searchPattern パラメーターに一致するサブディレクトリがない場合、このメソッドは空の配列を返します。</remarks>
      /// <remarks>
      /// EnumerateDirectories メソッドと GetDirectories メソッドの違い: EnumerateDirectories を使用すると、コレクション全体が返される前に名前のコレクションの列挙を開始できます。
      /// GetDirectories を使用する場合は、配列にアクセスする前に名前の配列全体が返されるのを待つ必要があります。
      /// そのため、多数のファイルやディレクトリを操作する場合は、EnumerateDirectories の方が効率的です。
      /// </remarks>
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
      public DirectoryInfo[] GetDirectories(string searchPattern, SearchOption searchOption)
      {
         return Directory.EnumerateFileSystemEntryInfosCore<DirectoryInfo>(true, Transaction, LongFullName, searchPattern, searchOption, null, null, PathFormat.LongFullPath).ToArray();
      }

      #endregion // .NET
   }
}
