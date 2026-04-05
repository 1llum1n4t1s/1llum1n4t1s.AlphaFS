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

      /// <summary>現在のディレクトリからファイルリストを返します。</summary>
      /// <returns><see cref="FileInfo"/> 型の配列。</returns>
      /// <remarks>返されるファイル名の順序は保証されません。特定の並べ替え順序が必要な場合は、Sort() メソッドを使用してください。</remarks>
      /// <remarks><see cref="DirectoryInfo"/> にファイルがない場合、このメソッドは空の配列を返します。</remarks>
      /// <remarks>
      /// EnumerateFiles メソッドと GetFiles メソッドの違い: EnumerateFiles を使用すると、コレクション全体が返される前に名前のコレクションの列挙を開始できます。
      /// GetFiles を使用する場合は、配列にアクセスする前に名前の配列全体が返されるのを待つ必要があります。
      /// そのため、多数のファイルやディレクトリを操作する場合は、EnumerateFiles の方が効率的です。
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      [SecurityCritical]
      public FileInfo[] GetFiles()
      {
         return Directory.EnumerateFileSystemEntryInfosCore<FileInfo>(false, Transaction, LongFullName, Path.WildcardStarMatchAll, null, null, null, PathFormat.LongFullPath).ToArray();
      }


      /// <summary>指定された検索パターンに一致する現在のディレクトリからファイルリストを返します。</summary>
      /// <param name="searchPattern">
      ///   パス内のディレクトリ名と照合する検索文字列。
      ///   このパラメーターには、有効なリテラルパスとワイルドカード
      ///   （<see cref="Path.WildcardStarMatchAll"/> および <see cref="Path.WildcardQuestion"/>）文字の組み合わせを含めることができますが、正規表現はサポートされません。
      /// </param>
      /// <returns><see cref="FileInfo"/> 型の配列。</returns>
      /// <remarks>返されるファイル名の順序は保証されません。特定の並べ替え順序が必要な場合は、Sort() メソッドを使用してください。</remarks>
      /// <remarks><see cref="DirectoryInfo"/> にファイルがない場合、このメソッドは空の配列を返します。</remarks>
      /// <remarks>
      /// EnumerateFiles メソッドと GetFiles メソッドの違い: EnumerateFiles を使用すると、コレクション全体が返される前に名前のコレクションの列挙を開始できます。
      /// GetFiles を使用する場合は、配列にアクセスする前に名前の配列全体が返されるのを待つ必要があります。
      /// そのため、多数のファイルやディレクトリを操作する場合は、EnumerateFiles の方が効率的です。
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      [SecurityCritical]
      public FileInfo[] GetFiles(string searchPattern)
      {
         return Directory.EnumerateFileSystemEntryInfosCore<FileInfo>(false, Transaction, LongFullName, searchPattern, null, null, null, PathFormat.LongFullPath).ToArray();
      }


      /// <summary>指定された検索パターンに一致し、サブディレクトリを検索するかどうかを決定する値を使用して、現在のディレクトリからファイルリストを返します。</summary>
      /// <param name="searchPattern">
      ///   パス内のディレクトリ名と照合する検索文字列。
      ///   このパラメーターには、有効なリテラルパスとワイルドカード
      ///   （<see cref="Path.WildcardStarMatchAll"/> および <see cref="Path.WildcardQuestion"/>）文字の組み合わせを含めることができますが、正規表現はサポートされません。
      /// </param>
      /// <param name="searchOption">
      ///   <paramref name="searchOption"/> が現在のディレクトリのみを含むか、すべてのサブディレクトリを含むかを指定する
      ///   <see cref="SearchOption"/> 列挙値の 1 つ。
      /// </param>
      /// <returns><see cref="FileInfo"/> 型の配列。</returns>
      /// <remarks>返されるファイル名の順序は保証されません。特定の並べ替え順序が必要な場合は、Sort() メソッドを使用してください。</remarks>
      /// <remarks><see cref="DirectoryInfo"/> にファイルがない場合、このメソッドは空の配列を返します。</remarks>
      /// <remarks>
      /// EnumerateFiles メソッドと GetFiles メソッドの違い: EnumerateFiles を使用すると、コレクション全体が返される前に名前のコレクションの列挙を開始できます。
      /// GetFiles を使用する場合は、配列にアクセスする前に名前の配列全体が返されるのを待つ必要があります。
      /// そのため、多数のファイルやディレクトリを操作する場合は、EnumerateFiles の方が効率的です。
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      [SecurityCritical]
      public FileInfo[] GetFiles(string searchPattern, SearchOption searchOption)
      {
         return Directory.EnumerateFileSystemEntryInfosCore<FileInfo>(false, Transaction, LongFullName, searchPattern, searchOption, null, null, PathFormat.LongFullPath).ToArray();
      }

      #endregion // .NET
   }
}
