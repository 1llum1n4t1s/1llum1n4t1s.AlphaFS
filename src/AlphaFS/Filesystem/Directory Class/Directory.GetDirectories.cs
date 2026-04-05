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
using SearchOption = System.IO.SearchOption;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class Directory
   {
      #region .NET

      /// <summary>Returns the names of subdirectories (including their paths) in the specified directory.</summary>
      /// <returns>An array of the full names (including paths) of subdirectories in the specified path, or an empty array if no directories are found.</returns>
      /// <remarks>
      ///   <para>このメソッドで返される名前には、pathで指定されたディレクトリ情報がプレフィックスとして付けられます。</para>
      ///   <para>EnumerateDirectoriesとGetDirectoriesメソッドは次のように異なります: EnumerateDirectoriesを使用すると、 you can start enumerating the collection of names
      ///     before the whole collection is returned; when you use GetDirectories, 配列にアクセスする前に名前の配列全体が返されるのを待つ必要があります。
      ///     したがって、多くのファイルとディレクトリを操作する場合、EnumerateDirectoriesの方が効率的です。
      ///   </para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="path">検索するディレクトリ。</param>
      [SecurityCritical]
      public static string[] GetDirectories(string path)
      {
         return EnumerateFileSystemEntryInfosCore<string>(true, null, path, Path.WildcardStarMatchAll, null, null, null, PathFormat.RelativePath).ToArray();
      }


      /// <summary>指定されたディレクトリ内で、指定された検索パターンに一致するサブディレクトリの名前（パスを含む）を返します。</summary>
      /// <returns>An array of the full names (including paths) of the subdirectories that match the search pattern in the specified directory, or an empty array if no directories are found.</returns>
      /// <remarks>
      ///   <para>このメソッドで返される名前には、pathで指定されたディレクトリ情報がプレフィックスとして付けられます。</para>
      ///   <para>EnumerateDirectoriesとGetDirectoriesメソッドは次のように異なります: EnumerateDirectoriesを使用すると、 you can start enumerating the collection of names
      ///     before the whole collection is returned; when you use GetDirectories, 配列にアクセスする前に名前の配列全体が返されるのを待つ必要があります。
      ///     したがって、多くのファイルとディレクトリを操作する場合、EnumerateDirectoriesの方が効率的です。
      ///   </para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="path">検索するディレクトリ。</param>
      /// <param name="searchPattern">
      ///   ディレクトリ名と照合する検索文字列。対象: <paramref name="path"/>.
      ///   このパラメータには、有効なリテラルパスとワイルドカードの組み合わせを含めることができますが、
      ///   (<see cref="Path.WildcardStarMatchAll"/> and <see cref="Path.WildcardQuestion"/>) characters, but does not support regular expressions.
      /// </param>
      [SecurityCritical]
      public static string[] GetDirectories(string path, string searchPattern)
      {
         return EnumerateFileSystemEntryInfosCore<string>(true, null, path, searchPattern, null, null, null, PathFormat.RelativePath).ToArray();
      }


      /// <summary>Returns the names of the subdirectories (including their paths) that match the specified search pattern in the specified directory, and optionally searches subdirectories.</summary>
      /// <returns>An array of the full names (including paths) of the subdirectories that match the specified criteria, or an empty array if no directories are found.</returns>
      /// <remarks>
      ///   <para>このメソッドで返される名前には、pathで指定されたディレクトリ情報がプレフィックスとして付けられます。</para>
      ///   <para>EnumerateDirectoriesとGetDirectoriesメソッドは次のように異なります: EnumerateDirectoriesを使用すると、 you can start enumerating the collection of names
      ///     before the whole collection is returned; when you use GetDirectories, 配列にアクセスする前に名前の配列全体が返されるのを待つ必要があります。
      ///     したがって、多くのファイルとディレクトリを操作する場合、EnumerateDirectoriesの方が効率的です。
      ///   </para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="path">検索するディレクトリ。</param>
      /// <param name="searchPattern">
      ///   ディレクトリ名と照合する検索文字列。対象: <paramref name="path"/>.
      ///   このパラメータには、有効なリテラルパスとワイルドカードの組み合わせを含めることができますが、
      ///   (<see cref="Path.WildcardStarMatchAll"/> and <see cref="Path.WildcardQuestion"/>) characters, but does not support regular expressions.
      /// </param>
      /// <param name="searchOption">
      ///   One of the <see cref="SearchOption"/> enumeration values that specifies whether the <paramref name="searchOption"/>
      ///   現在のディレクトリのみを含むか、すべてのサブディレクトリを含むかを指定します。
      /// </param>
      [SecurityCritical]
      public static string[] GetDirectories(string path, string searchPattern, SearchOption searchOption)
      {
         return EnumerateFileSystemEntryInfosCore<string>(true, null, path, searchPattern, searchOption, null, null, PathFormat.RelativePath).ToArray();
      }

      #endregion // .NET
   }
}
