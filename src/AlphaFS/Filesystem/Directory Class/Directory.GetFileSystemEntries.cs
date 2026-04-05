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

      /// <summary>Returns the names of all files and subdirectories in the specified directory.</summary>
      /// <returns>An string[] array of the names of files and subdirectories in the specified directory.</returns>
      /// <remarks>
      ///   <para>EnumerateFileSystemEntriesとGetFileSystemEntriesメソッドは次のように異なります:  EnumerateFileSystemEntries,
      ///     コレクション全体が返される前にエントリの列挙を開始できます。 GetFileSystemEntries,
      ///     配列にアクセスする前にエントリの配列全体が返されるのを待つ必要があります。
      ///     したがって、多くのファイルとディレクトリを操作する場合、EnumerateFilesの方が効率的です。
      ///   </para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="path">The directory for which file and subdirectory names are returned.</param>
      [SecurityCritical]
      public static string[] GetFileSystemEntries(string path)
      {
         return EnumerateFileSystemEntryInfosCore<string>(null, null, path, Path.WildcardStarMatchAll, null, null, null, PathFormat.RelativePath).ToArray();
      }


      /// <summary>Returns an array of file system entries that match the specified search criteria.</summary>
      /// <returns>指定された検索条件に一致するファイルシステムエントリの string[] 配列。</returns>
      /// <remarks>
      ///   <para>EnumerateFileSystemEntriesとGetFileSystemEntriesメソッドは次のように異なります:  EnumerateFileSystemEntries,
      ///     コレクション全体が返される前にエントリの列挙を開始できます。 GetFileSystemEntries,
      ///     配列にアクセスする前にエントリの配列全体が返されるのを待つ必要があります。
      ///     したがって、多くのファイルとディレクトリを操作する場合、EnumerateFilesの方が効率的です。
      ///   </para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="path">The path to be searched.</param>
      /// <param name="searchPattern">
      ///   ディレクトリ名と照合する検索文字列。対象: <paramref name="path"/>.
      ///   このパラメータには、有効なリテラルパスとワイルドカードの組み合わせを含めることができますが、
      ///   (<see cref="Path.WildcardStarMatchAll"/> and <see cref="Path.WildcardQuestion"/>) characters, but does not support regular expressions.
      /// </param>
      [SecurityCritical]
      public static string[] GetFileSystemEntries(string path, string searchPattern)
      {
         return EnumerateFileSystemEntryInfosCore<string>(null, null, path, searchPattern, null, null, null, PathFormat.RelativePath).ToArray();
      }


      /// <summary>Gets an array of all the file names and directory names that match a <paramref name="searchPattern"/> in a specified path, and optionally searches subdirectories.</summary>
      /// <returns>指定された検索条件に一致するファイルシステムエントリの string[] 配列。</returns>
      /// <remarks>
      ///   <para>EnumerateFileSystemEntriesとGetFileSystemEntriesメソッドは次のように異なります:  EnumerateFileSystemEntries,
      ///     コレクション全体が返される前にエントリの列挙を開始できます。 GetFileSystemEntries,
      ///     配列にアクセスする前にエントリの配列全体が返されるのを待つ必要があります。
      ///     したがって、多くのファイルとディレクトリを操作する場合、EnumerateFilesの方が効率的です。
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
      public static string[] GetFileSystemEntries(string path, string searchPattern, SearchOption searchOption)
      {
         return EnumerateFileSystemEntryInfosCore<string>(null, null, path, searchPattern, searchOption, null, null, PathFormat.RelativePath).ToArray();
      }

      #endregion // .NET
   }
}
