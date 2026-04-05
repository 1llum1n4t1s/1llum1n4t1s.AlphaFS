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
using SearchOption = System.IO.SearchOption;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class Directory
   {
      #region .NET

      /// <summary>の列挙可能なコレクションを返します。 directory names in a specified <paramref name="path"/>.</summary>
      /// <returns><paramref name="path"/> で指定されたディレクトリ内のディレクトリのフルネーム（パスを含む）の列挙可能なコレクション。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="path">検索するディレクトリ。</param>
      [SecurityCritical]
      public static IEnumerable<string> EnumerateDirectories(string path)
      {
         return EnumerateFileSystemEntryInfosCore<string>(true, null, path, Path.WildcardStarMatchAll, null, null, null, PathFormat.RelativePath);
      }


      /// <summary>の列挙可能なコレクションを返します。 directory names that match a <paramref name="searchPattern"/> in a specified <paramref name="path"/>.</summary>
      /// <returns>指定されたディレクトリ内のディレクトリのフルネーム（パスを含む）の列挙可能なコレクション。対象: <paramref name="path"/> and that match the specified <paramref name="searchPattern"/>.</returns>
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
      public static IEnumerable<string> EnumerateDirectories(string path, string searchPattern)
      {
         return EnumerateFileSystemEntryInfosCore<string>(true, null, path, searchPattern, null, null, null, PathFormat.RelativePath);
      }


      /// <summary>の列挙可能なコレクションを返します。 directory names that match a <paramref name="searchPattern"/> in a specified <paramref name="path"/>, and optionally searches subdirectories.</summary>
      /// <returns>指定されたディレクトリ内のディレクトリのフルネーム（パスを含む）の列挙可能なコレクション。対象: <paramref name="path"/> and that match the specified <paramref name="searchPattern"/> and <paramref name="searchOption"/>.</returns>
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
      public static IEnumerable<string> EnumerateDirectories(string path, string searchPattern, SearchOption searchOption)
      {
         return EnumerateFileSystemEntryInfosCore<string>(true, null, path, searchPattern, searchOption, null, null, PathFormat.RelativePath);
      }

      #endregion // .NET


      /// <summary>[AlphaFS] の列挙可能なコレクションを返します。 directory names in a specified <paramref name="path"/>.</summary>
      /// <returns><paramref name="path"/> で指定されたディレクトリ内のディレクトリのフルネーム（パスを含む）の列挙可能なコレクション。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="path">検索するディレクトリ。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static IEnumerable<string> EnumerateDirectories(string path, PathFormat pathFormat)
      {
         return EnumerateFileSystemEntryInfosCore<string>(true, null, path, Path.WildcardStarMatchAll, null, null, null, pathFormat);
      }


      /// <summary>[AlphaFS] の列挙可能なコレクションを返します。 directory names that match a <paramref name="searchPattern"/> in a specified <paramref name="path"/>.</summary>
      /// <returns>指定されたディレクトリ内のディレクトリのフルネーム（パスを含む）の列挙可能なコレクション。対象: <paramref name="path"/> and that match the specified <paramref name="searchPattern"/>.</returns>
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
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      [Obsolete("Argument searchPattern is obsolete. The DirectoryEnumerationFilters argument provides better filter criteria.")]
      public static IEnumerable<string> EnumerateDirectories(string path, string searchPattern, PathFormat pathFormat)
      {
         return EnumerateFileSystemEntryInfosCore<string>(true, null, path, searchPattern, null, null, null, pathFormat);
      }


      /// <summary>[AlphaFS] の列挙可能なコレクションを返します。 directory names that match a <paramref name="searchPattern"/> in a specified <paramref name="path"/>, and optionally searches subdirectories.</summary>
      /// <returns>指定されたディレクトリ内のディレクトリのフルネーム（パスを含む）の列挙可能なコレクション。対象: <paramref name="path"/> and that match the specified <paramref name="searchPattern"/> and <paramref name="searchOption"/>.</returns>
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
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      [Obsolete("Argument searchPattern is obsolete. The DirectoryEnumerationFilters argument provides better filter criteria.")]
      public static IEnumerable<string> EnumerateDirectories(string path, string searchPattern, SearchOption searchOption, PathFormat pathFormat)
      {
         return EnumerateFileSystemEntryInfosCore<string>(true, null, path, searchPattern, searchOption, null, null, pathFormat);
      }


      /// <summary>[AlphaFS] の列挙可能なコレクションを返します。 directory names in a specified <paramref name="path"/>.</summary>
      /// <returns><paramref name="path"/> で指定されたディレクトリ内のディレクトリのフルネーム（パスを含む）の列挙可能なコレクション。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="path">検索するディレクトリ。</param>
      /// <param name="options">ディレクトリの列挙方法を指定する <see cref="DirectoryEnumerationOptions"/> フラグ。</param>
      [SecurityCritical]
      public static IEnumerable<string> EnumerateDirectories(string path, DirectoryEnumerationOptions options)
      {
         return EnumerateFileSystemEntryInfosCore<string>(true, null, path, Path.WildcardStarMatchAll, null, options, null, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] の列挙可能なコレクションを返します。 directory names in a specified <paramref name="path"/>.</summary>
      /// <returns><paramref name="path"/> で指定されたディレクトリ内のディレクトリのフルネーム（パスを含む）の列挙可能なコレクション。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="path">検索するディレクトリ。</param>
      /// <param name="options">ディレクトリの列挙方法を指定する <see cref="DirectoryEnumerationOptions"/> フラグ。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static IEnumerable<string> EnumerateDirectories(string path, DirectoryEnumerationOptions options, PathFormat pathFormat)
      {
         return EnumerateFileSystemEntryInfosCore<string>(true, null, path, Path.WildcardStarMatchAll, null, options, null, pathFormat);
      }


      /// <summary>[AlphaFS] の列挙可能なコレクションを返します。 directory names in a specified <paramref name="path"/>.</summary>
      /// <returns><paramref name="path"/> で指定されたディレクトリ内のディレクトリのフルネーム（パスを含む）の列挙可能なコレクション。</returns>
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
      /// <param name="options">ディレクトリの列挙方法を指定する <see cref="DirectoryEnumerationOptions"/> フラグ。</param>
      [SecurityCritical]
      [Obsolete("Argument searchPattern is obsolete. The DirectoryEnumerationFilters argument provides better filter criteria.")]
      public static IEnumerable<string> EnumerateDirectories(string path, string searchPattern, DirectoryEnumerationOptions options)
      {
         return EnumerateFileSystemEntryInfosCore<string>(true, null, path, searchPattern, null, options, null, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] の列挙可能なコレクションを返します。 directory names in a specified <paramref name="path"/>.</summary>
      /// <returns><paramref name="path"/> で指定されたディレクトリ内のディレクトリのフルネーム（パスを含む）の列挙可能なコレクション。</returns>
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
      /// <param name="options">ディレクトリの列挙方法を指定する <see cref="DirectoryEnumerationOptions"/> フラグ。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      [Obsolete("Argument searchPattern is obsolete. The DirectoryEnumerationFilters argument provides better filter criteria.")]
      public static IEnumerable<string> EnumerateDirectories(string path, string searchPattern, DirectoryEnumerationOptions options, PathFormat pathFormat)
      {
         return EnumerateFileSystemEntryInfosCore<string>(true, null, path, searchPattern, null, options, null, pathFormat);
      }


      /// <summary>[AlphaFS] の列挙可能なコレクションを返します。 directory names in a specified <paramref name="path"/>.</summary>
      /// <returns><paramref name="path"/> で指定されたディレクトリ内のディレクトリのフルネーム（パスを含む）の列挙可能なコレクション。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="path">検索するディレクトリ。</param>
      /// <param name="filters">処理で使用するカスタムフィルタの指定。</param>
      [SecurityCritical]
      public static IEnumerable<string> EnumerateDirectories(string path, DirectoryEnumerationFilters filters)
      {
         return EnumerateFileSystemEntryInfosCore<string>(true, null, path, Path.WildcardStarMatchAll, null, null, filters, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] の列挙可能なコレクションを返します。 directory names in a specified <paramref name="path"/>.</summary>
      /// <returns><paramref name="path"/> で指定されたディレクトリ内のディレクトリのフルネーム（パスを含む）の列挙可能なコレクション。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="path">検索するディレクトリ。</param>
      /// <param name="filters">処理で使用するカスタムフィルタの指定。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static IEnumerable<string> EnumerateDirectories(string path, DirectoryEnumerationFilters filters, PathFormat pathFormat)
      {
         return EnumerateFileSystemEntryInfosCore<string>(true, null, path, Path.WildcardStarMatchAll, null, null, filters, pathFormat);
      }


      /// <summary>[AlphaFS] の列挙可能なコレクションを返します。 directory names in a specified <paramref name="path"/>.</summary>
      /// <returns><paramref name="path"/> で指定されたディレクトリ内のディレクトリのフルネーム（パスを含む）の列挙可能なコレクション。</returns>
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
      /// <param name="filters">処理で使用するカスタムフィルタの指定。</param>
      [SecurityCritical]
      [Obsolete("Argument searchPattern is obsolete. The DirectoryEnumerationFilters argument provides better filter criteria.")]
      public static IEnumerable<string> EnumerateDirectories(string path, string searchPattern, DirectoryEnumerationFilters filters)
      {
         return EnumerateFileSystemEntryInfosCore<string>(true, null, path, searchPattern, null, null, filters, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] の列挙可能なコレクションを返します。 directory names in a specified <paramref name="path"/>.</summary>
      /// <returns><paramref name="path"/> で指定されたディレクトリ内のディレクトリのフルネーム（パスを含む）の列挙可能なコレクション。</returns>
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
      /// <param name="filters">処理で使用するカスタムフィルタの指定。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      [Obsolete("Argument searchPattern is obsolete. The DirectoryEnumerationFilters argument provides better filter criteria.")]
      public static IEnumerable<string> EnumerateDirectories(string path, string searchPattern, DirectoryEnumerationFilters filters, PathFormat pathFormat)
      {
         return EnumerateFileSystemEntryInfosCore<string>(true, null, path, searchPattern, null, null, filters, pathFormat);
      }


      /// <summary>[AlphaFS] の列挙可能なコレクションを返します。 directory names in a specified <paramref name="path"/>.</summary>
      /// <returns><paramref name="path"/> で指定されたディレクトリ内のディレクトリのフルネーム（パスを含む）の列挙可能なコレクション。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="path">検索するディレクトリ。</param>
      /// <param name="options">ディレクトリの列挙方法を指定する <see cref="DirectoryEnumerationOptions"/> フラグ。</param>
      /// <param name="filters">処理で使用するカスタムフィルタの指定。</param>
      [SecurityCritical]
      public static IEnumerable<string> EnumerateDirectories(string path, DirectoryEnumerationOptions options, DirectoryEnumerationFilters filters)
      {
         return EnumerateFileSystemEntryInfosCore<string>(true, null, path, Path.WildcardStarMatchAll, null, options, filters, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] の列挙可能なコレクションを返します。 directory names in a specified <paramref name="path"/>.</summary>
      /// <returns><paramref name="path"/> で指定されたディレクトリ内のディレクトリのフルネーム（パスを含む）の列挙可能なコレクション。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="path">検索するディレクトリ。</param>
      /// <param name="options">ディレクトリの列挙方法を指定する <see cref="DirectoryEnumerationOptions"/> フラグ。</param>
      /// <param name="filters">処理で使用するカスタムフィルタの指定。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static IEnumerable<string> EnumerateDirectories(string path, DirectoryEnumerationOptions options, DirectoryEnumerationFilters filters, PathFormat pathFormat)
      {
         return EnumerateFileSystemEntryInfosCore<string>(true, null, path, Path.WildcardStarMatchAll, null, options, filters, pathFormat);
      }


      /// <summary>[AlphaFS] の列挙可能なコレクションを返します。 directory names in a specified <paramref name="path"/>.</summary>
      /// <returns><paramref name="path"/> で指定されたディレクトリ内のディレクトリのフルネーム（パスを含む）の列挙可能なコレクション。</returns>
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
      /// <param name="options">ディレクトリの列挙方法を指定する <see cref="DirectoryEnumerationOptions"/> フラグ。</param>
      /// <param name="filters">処理で使用するカスタムフィルタの指定。</param>
      [SecurityCritical]
      [Obsolete("Argument searchPattern is obsolete. The DirectoryEnumerationFilters argument provides better filter criteria.")]
      public static IEnumerable<string> EnumerateDirectories(string path, string searchPattern, DirectoryEnumerationOptions options, DirectoryEnumerationFilters filters)
      {
         return EnumerateFileSystemEntryInfosCore<string>(true, null, path, searchPattern, null, options, filters, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] の列挙可能なコレクションを返します。 directory names in a specified <paramref name="path"/>.</summary>
      /// <returns><paramref name="path"/> で指定されたディレクトリ内のディレクトリのフルネーム（パスを含む）の列挙可能なコレクション。</returns>
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
      /// <param name="filters">処理で使用するカスタムフィルタの指定。</param>
      /// <param name="options">ディレクトリの列挙方法を指定する <see cref="DirectoryEnumerationOptions"/> フラグ。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      [Obsolete("Argument searchPattern is obsolete. The DirectoryEnumerationFilters argument provides better filter criteria.")]
      public static IEnumerable<string> EnumerateDirectories(string path, string searchPattern, DirectoryEnumerationOptions options, DirectoryEnumerationFilters filters, PathFormat pathFormat)
      {
         return EnumerateFileSystemEntryInfosCore<string>(true, null, path, searchPattern, null, options, filters, pathFormat);
      }
   }
}
