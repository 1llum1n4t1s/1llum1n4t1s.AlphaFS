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
   public static partial class Directory
   {
      /// <summary>[AlphaFS] 指定されたパス内のファイルシステムエントリの列挙可能なコレクションを返します。</summary>
      /// <returns>一致するファイルシステムエントリ。項目の型は <typeparamref name="T"/> によって決定されます。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <typeparam name="T">The type to return. This may be one of the following types:
      ///    <list type="definition">
      ///    <item>
      ///       <term><see cref="FileSystemEntryInfo"/></term>
      ///       <description>This method will return instances of <see cref="FileSystemEntryInfo"/> instances.</description>
      ///    </item>
      ///    <item>
      ///       <term><see cref="FileSystemInfo"/></term>
      ///       <description>This method will return instances of <see cref="DirectoryInfo"/> and <see cref="FileInfo"/> instances.</description>
      ///    </item>
      ///    <item>
      ///       <term><see cref="string"/></term>
      ///       <description>This method will return the full path of each item.</description>
      ///    </item>
      /// </list>
      /// </typeparam>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">検索するディレクトリ。</param>
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Infos")]
      [SecurityCritical]
      public static IEnumerable<T> EnumerateFileSystemEntryInfosTransacted<T>(KernelTransaction transaction, string path)
      {
         return EnumerateFileSystemEntryInfosCore<T>(null, transaction, path, Path.WildcardStarMatchAll, null, null, null, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] 指定されたパス内のファイルシステムエントリの列挙可能なコレクションを返します。</summary>
      /// <returns>一致するファイルシステムエントリ。項目の型は <typeparamref name="T"/> によって決定されます。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <typeparam name="T">The type to return. This may be one of the following types:
      ///    <list type="definition">
      ///    <item>
      ///       <term><see cref="FileSystemEntryInfo"/></term>
      ///       <description>This method will return instances of <see cref="FileSystemEntryInfo"/> instances.</description>
      ///    </item>
      ///    <item>
      ///       <term><see cref="FileSystemInfo"/></term>
      ///       <description>This method will return instances of <see cref="DirectoryInfo"/> and <see cref="FileInfo"/> instances.</description>
      ///    </item>
      ///    <item>
      ///       <term><see cref="string"/></term>
      ///       <description>This method will return the full path of each item.</description>
      ///    </item>
      /// </list>
      /// </typeparam>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">検索するディレクトリ。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Infos")]
      [SecurityCritical]
      public static IEnumerable<T> EnumerateFileSystemEntryInfosTransacted<T>(KernelTransaction transaction, string path, PathFormat pathFormat)
      {
         return EnumerateFileSystemEntryInfosCore<T>(null, transaction, path, Path.WildcardStarMatchAll, null, null, null, pathFormat);
      }
      

      /// <summary>[AlphaFS] 指定されたパス内のファイルシステムエントリの列挙可能なコレクションを返します。</summary>
      /// <returns>一致するファイルシステムエントリ。項目の型は <typeparamref name="T"/> によって決定されます。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <typeparam name="T">The type to return. This may be one of the following types:
      ///    <list type="definition">
      ///    <item>
      ///       <term><see cref="FileSystemEntryInfo"/></term>
      ///       <description>This method will return instances of <see cref="FileSystemEntryInfo"/> instances.</description>
      ///    </item>
      ///    <item>
      ///       <term><see cref="FileSystemInfo"/></term>
      ///       <description>This method will return instances of <see cref="DirectoryInfo"/> and <see cref="FileInfo"/> instances.</description>
      ///    </item>
      ///    <item>
      ///       <term><see cref="string"/></term>
      ///       <description>This method will return the full path of each item.</description>
      ///    </item>
      /// </list>
      /// </typeparam>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">検索するディレクトリ。</param>
      /// <param name="options">ディレクトリの列挙方法を指定する <see cref="DirectoryEnumerationOptions"/> フラグ。</param>
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Infos")]
      [SecurityCritical]
      public static IEnumerable<T> EnumerateFileSystemEntryInfosTransacted<T>(KernelTransaction transaction, string path, DirectoryEnumerationOptions options)
      {
         return EnumerateFileSystemEntryInfosCore<T>(null, transaction, path, Path.WildcardStarMatchAll, null, options,  null, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] 指定されたパス内のファイルシステムエントリの列挙可能なコレクションを返します。</summary>
      /// <returns>一致するファイルシステムエントリ。項目の型は <typeparamref name="T"/> によって決定されます。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <typeparam name="T">The type to return. This may be one of the following types:
      ///    <list type="definition">
      ///    <item>
      ///       <term><see cref="FileSystemEntryInfo"/></term>
      ///       <description>This method will return instances of <see cref="FileSystemEntryInfo"/> instances.</description>
      ///    </item>
      ///    <item>
      ///       <term><see cref="FileSystemInfo"/></term>
      ///       <description>This method will return instances of <see cref="DirectoryInfo"/> and <see cref="FileInfo"/> instances.</description>
      ///    </item>
      ///    <item>
      ///       <term><see cref="string"/></term>
      ///       <description>This method will return the full path of each item.</description>
      ///    </item>
      /// </list>
      /// </typeparam>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">検索するディレクトリ。</param>
      /// <param name="options">ディレクトリの列挙方法を指定する <see cref="DirectoryEnumerationOptions"/> フラグ。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Infos")]
      [SecurityCritical]
      public static IEnumerable<T> EnumerateFileSystemEntryInfosTransacted<T>(KernelTransaction transaction, string path, DirectoryEnumerationOptions options, PathFormat pathFormat)
      {
         return EnumerateFileSystemEntryInfosCore<T>(null, transaction, path, Path.WildcardStarMatchAll, null, options,  null, pathFormat);
      }
      

      /// <summary>[AlphaFS] の列挙可能なコレクションを返します。 file system entries that match a <paramref name="searchPattern"/> 指定されたパス内の</summary>
      /// <returns>一致するファイルシステムエントリ。項目の型は <typeparamref name="T"/> によって決定されます。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <typeparam name="T">The type to return. This may be one of the following types:
      ///    <list type="definition">
      ///    <item>
      ///       <term><see cref="FileSystemEntryInfo"/></term>
      ///       <description>This method will return instances of <see cref="FileSystemEntryInfo"/> instances.</description>
      ///    </item>
      ///    <item>
      ///       <term><see cref="FileSystemInfo"/></term>
      ///       <description>This method will return instances of <see cref="DirectoryInfo"/> and <see cref="FileInfo"/> instances.</description>
      ///    </item>
      ///    <item>
      ///       <term><see cref="string"/></term>
      ///       <description>This method will return the full path of each item.</description>
      ///    </item>
      /// </list>
      /// </typeparam>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">検索するディレクトリ。</param>
      /// <param name="searchPattern">
      ///   ディレクトリ名と照合する検索文字列。対象: <paramref name="path"/>.
      ///   このパラメータには、有効なリテラルパスとワイルドカードの組み合わせを含めることができますが、
      ///   (<see cref="Path.WildcardStarMatchAll"/> and <see cref="Path.WildcardQuestion"/>) characters, but does not support regular expressions.
      /// </param>
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Infos")]
      [SecurityCritical]
      [Obsolete("Argument searchPattern is obsolete. The DirectoryEnumerationFilters argument provides better filter criteria.")]
      public static IEnumerable<T> EnumerateFileSystemEntryInfosTransacted<T>(KernelTransaction transaction, string path, string searchPattern)
      {
         return EnumerateFileSystemEntryInfosCore<T>(null, transaction, path, searchPattern, null, null, null, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] の列挙可能なコレクションを返します。 file system entries that match a <paramref name="searchPattern"/> 指定されたパス内の</summary>
      /// <returns>一致するファイルシステムエントリ。項目の型は <typeparamref name="T"/> によって決定されます。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <typeparam name="T">The type to return. This may be one of the following types:
      ///    <list type="definition">
      ///    <item>
      ///       <term><see cref="FileSystemEntryInfo"/></term>
      ///       <description>This method will return instances of <see cref="FileSystemEntryInfo"/> instances.</description>
      ///    </item>
      ///    <item>
      ///       <term><see cref="FileSystemInfo"/></term>
      ///       <description>This method will return instances of <see cref="DirectoryInfo"/> and <see cref="FileInfo"/> instances.</description>
      ///    </item>
      ///    <item>
      ///       <term><see cref="string"/></term>
      ///       <description>This method will return the full path of each item.</description>
      ///    </item>
      /// </list>
      /// </typeparam>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">検索するディレクトリ。</param>
      /// <param name="searchPattern">
      ///   ディレクトリ名と照合する検索文字列。対象: <paramref name="path"/>.
      ///   このパラメータには、有効なリテラルパスとワイルドカードの組み合わせを含めることができますが、
      ///   (<see cref="Path.WildcardStarMatchAll"/> and <see cref="Path.WildcardQuestion"/>) characters, but does not support regular expressions.
      /// </param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Infos")]
      [SecurityCritical]
      [Obsolete("Argument searchPattern is obsolete. The DirectoryEnumerationFilters argument provides better filter criteria.")]
      public static IEnumerable<T> EnumerateFileSystemEntryInfosTransacted<T>(KernelTransaction transaction, string path, string searchPattern, PathFormat pathFormat)
      {
         return EnumerateFileSystemEntryInfosCore<T>(null, transaction, path, searchPattern, null, null, null, pathFormat);
      }
      

      /// <summary>[AlphaFS] の列挙可能なコレクションを返します。 file system entries that match a <paramref name="searchPattern"/> in a specified path using <see cref="DirectoryEnumerationOptions"/>.</summary>
      /// <returns>一致するファイルシステムエントリ。項目の型は <typeparamref name="T"/> によって決定されます。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <typeparam name="T">The type to return. This may be one of the following types:
      ///    <list type="definition">
      ///    <item>
      ///       <term><see cref="FileSystemEntryInfo"/></term>
      ///       <description>This method will return instances of <see cref="FileSystemEntryInfo"/> instances.</description>
      ///    </item>
      ///    <item>
      ///       <term><see cref="FileSystemInfo"/></term>
      ///       <description>This method will return instances of <see cref="DirectoryInfo"/> and <see cref="FileInfo"/> instances.</description>
      ///    </item>
      ///    <item>
      ///       <term><see cref="string"/></term>
      ///       <description>This method will return the full path of each item.</description>
      ///    </item>
      /// </list>
      /// </typeparam>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">検索するディレクトリ。</param>
      /// <param name="searchPattern">
      ///   ディレクトリ名と照合する検索文字列。対象: <paramref name="path"/>.
      ///   このパラメータには、有効なリテラルパスとワイルドカードの組み合わせを含めることができますが、
      ///   (<see cref="Path.WildcardStarMatchAll"/> and <see cref="Path.WildcardQuestion"/>) characters, but does not support regular expressions.
      /// </param>
      /// <param name="options">ディレクトリの列挙方法を指定する <see cref="DirectoryEnumerationOptions"/> フラグ。</param>
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Infos")]
      [SecurityCritical]
      [Obsolete("Argument searchPattern is obsolete. The DirectoryEnumerationFilters argument provides better filter criteria.")]
      public static IEnumerable<T> EnumerateFileSystemEntryInfosTransacted<T>(KernelTransaction transaction, string path, string searchPattern, DirectoryEnumerationOptions options)
      {
         return EnumerateFileSystemEntryInfosCore<T>(null, transaction, path, searchPattern, null, options,  null, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] の列挙可能なコレクションを返します。 file system entries that match a <paramref name="searchPattern"/> in a specified path using <see cref="DirectoryEnumerationOptions"/>.</summary>
      /// <returns>一致するファイルシステムエントリ。項目の型は <typeparamref name="T"/> によって決定されます。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <typeparam name="T">The type to return. This may be one of the following types:
      ///    <list type="definition">
      ///    <item>
      ///       <term><see cref="FileSystemEntryInfo"/></term>
      ///       <description>This method will return instances of <see cref="FileSystemEntryInfo"/> instances.</description>
      ///    </item>
      ///    <item>
      ///       <term><see cref="FileSystemInfo"/></term>
      ///       <description>This method will return instances of <see cref="DirectoryInfo"/> and <see cref="FileInfo"/> instances.</description>
      ///    </item>
      ///    <item>
      ///       <term><see cref="string"/></term>
      ///       <description>This method will return the full path of each item.</description>
      ///    </item>
      /// </list>
      /// </typeparam>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">検索するディレクトリ。</param>
      /// <param name="searchPattern">
      ///   ディレクトリ名と照合する検索文字列。対象: <paramref name="path"/>.
      ///   このパラメータには、有効なリテラルパスとワイルドカードの組み合わせを含めることができますが、
      ///   (<see cref="Path.WildcardStarMatchAll"/> and <see cref="Path.WildcardQuestion"/>) characters, but does not support regular expressions.
      /// </param>
      /// <param name="options">ディレクトリの列挙方法を指定する <see cref="DirectoryEnumerationOptions"/> フラグ。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Infos")]
      [SecurityCritical]
      [Obsolete("Argument searchPattern is obsolete. The DirectoryEnumerationFilters argument provides better filter criteria.")]
      public static IEnumerable<T> EnumerateFileSystemEntryInfosTransacted<T>(KernelTransaction transaction, string path, string searchPattern, DirectoryEnumerationOptions options, PathFormat pathFormat)
      {
         return EnumerateFileSystemEntryInfosCore<T>(null, transaction, path, searchPattern, null, options,  null, pathFormat);
      }


      /// <summary>[AlphaFS] 指定されたパス内のファイルシステムエントリの列挙可能なコレクションを返します。</summary>
      /// <returns>一致するファイルシステムエントリ。項目の型は <typeparamref name="T"/> によって決定されます。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <typeparam name="T">The type to return. This may be one of the following types:
      ///    <list type="definition">
      ///    <item>
      ///       <term><see cref="FileSystemEntryInfo"/></term>
      ///       <description>This method will return instances of <see cref="FileSystemEntryInfo"/> instances.</description>
      ///    </item>
      ///    <item>
      ///       <term><see cref="FileSystemInfo"/></term>
      ///       <description>This method will return instances of <see cref="DirectoryInfo"/> and <see cref="FileInfo"/> instances.</description>
      ///    </item>
      ///    <item>
      ///       <term><see cref="string"/></term>
      ///       <description>This method will return the full path of each item.</description>
      ///    </item>
      /// </list>
      /// </typeparam>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">検索するディレクトリ。</param>
      /// <param name="filters">処理で使用するカスタムフィルタの指定。</param>
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Infos")]
      [SecurityCritical]
      public static IEnumerable<T> EnumerateFileSystemEntryInfosTransacted<T>(KernelTransaction transaction, string path, DirectoryEnumerationFilters filters)
      {
         return EnumerateFileSystemEntryInfosCore<T>(null, transaction, path, Path.WildcardStarMatchAll, null, null, filters, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] 指定されたパス内のファイルシステムエントリの列挙可能なコレクションを返します。</summary>
      /// <returns>一致するファイルシステムエントリ。項目の型は <typeparamref name="T"/> によって決定されます。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <typeparam name="T">The type to return. This may be one of the following types:
      ///    <list type="definition">
      ///    <item>
      ///       <term><see cref="FileSystemEntryInfo"/></term>
      ///       <description>This method will return instances of <see cref="FileSystemEntryInfo"/> instances.</description>
      ///    </item>
      ///    <item>
      ///       <term><see cref="FileSystemInfo"/></term>
      ///       <description>This method will return instances of <see cref="DirectoryInfo"/> and <see cref="FileInfo"/> instances.</description>
      ///    </item>
      ///    <item>
      ///       <term><see cref="string"/></term>
      ///       <description>This method will return the full path of each item.</description>
      ///    </item>
      /// </list>
      /// </typeparam>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">検索するディレクトリ。</param>
      /// <param name="filters">処理で使用するカスタムフィルタの指定。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Infos")]
      [SecurityCritical]
      public static IEnumerable<T> EnumerateFileSystemEntryInfosTransacted<T>(KernelTransaction transaction, string path, DirectoryEnumerationFilters filters, PathFormat pathFormat)
      {
         return EnumerateFileSystemEntryInfosCore<T>(null, transaction, path, Path.WildcardStarMatchAll, null, null, filters, pathFormat);
      }


      /// <summary>[AlphaFS] 指定されたパス内のファイルシステムエントリの列挙可能なコレクションを返します。</summary>
      /// <returns>一致するファイルシステムエントリ。項目の型は <typeparamref name="T"/> によって決定されます。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <typeparam name="T">The type to return. This may be one of the following types:
      ///    <list type="definition">
      ///    <item>
      ///       <term><see cref="FileSystemEntryInfo"/></term>
      ///       <description>This method will return instances of <see cref="FileSystemEntryInfo"/> instances.</description>
      ///    </item>
      ///    <item>
      ///       <term><see cref="FileSystemInfo"/></term>
      ///       <description>This method will return instances of <see cref="DirectoryInfo"/> and <see cref="FileInfo"/> instances.</description>
      ///    </item>
      ///    <item>
      ///       <term><see cref="string"/></term>
      ///       <description>This method will return the full path of each item.</description>
      ///    </item>
      /// </list>
      /// </typeparam>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">検索するディレクトリ。</param>
      /// <param name="options">ディレクトリの列挙方法を指定する <see cref="DirectoryEnumerationOptions"/> フラグ。</param>
      /// <param name="filters">処理で使用するカスタムフィルタの指定。</param>
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Infos")]
      [SecurityCritical]
      public static IEnumerable<T> EnumerateFileSystemEntryInfosTransacted<T>(KernelTransaction transaction, string path, DirectoryEnumerationOptions options, DirectoryEnumerationFilters filters)
      {
         return EnumerateFileSystemEntryInfosCore<T>(null, transaction, path, Path.WildcardStarMatchAll, null, options,  filters, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] 指定されたパス内のファイルシステムエントリの列挙可能なコレクションを返します。</summary>
      /// <returns>一致するファイルシステムエントリ。項目の型は <typeparamref name="T"/> によって決定されます。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <typeparam name="T">The type to return. This may be one of the following types:
      ///    <list type="definition">
      ///    <item>
      ///       <term><see cref="FileSystemEntryInfo"/></term>
      ///       <description>This method will return instances of <see cref="FileSystemEntryInfo"/> instances.</description>
      ///    </item>
      ///    <item>
      ///       <term><see cref="FileSystemInfo"/></term>
      ///       <description>This method will return instances of <see cref="DirectoryInfo"/> and <see cref="FileInfo"/> instances.</description>
      ///    </item>
      ///    <item>
      ///       <term><see cref="string"/></term>
      ///       <description>This method will return the full path of each item.</description>
      ///    </item>
      /// </list>
      /// </typeparam>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">検索するディレクトリ。</param>
      /// <param name="options">ディレクトリの列挙方法を指定する <see cref="DirectoryEnumerationOptions"/> フラグ。</param>
      /// <param name="filters">処理で使用するカスタムフィルタの指定。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Infos")]
      [SecurityCritical]
      public static IEnumerable<T> EnumerateFileSystemEntryInfosTransacted<T>(KernelTransaction transaction, string path, DirectoryEnumerationOptions options, DirectoryEnumerationFilters filters, PathFormat pathFormat)
      {
         return EnumerateFileSystemEntryInfosCore<T>(null, transaction, path, Path.WildcardStarMatchAll, null, options,  filters, pathFormat);
      }


      /// <summary>[AlphaFS] の列挙可能なコレクションを返します。 file system entries that match a <paramref name="searchPattern"/> 指定されたパス内の</summary>
      /// <returns>一致するファイルシステムエントリ。項目の型は <typeparamref name="T"/> によって決定されます。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <typeparam name="T">The type to return. This may be one of the following types:
      ///    <list type="definition">
      ///    <item>
      ///       <term><see cref="FileSystemEntryInfo"/></term>
      ///       <description>This method will return instances of <see cref="FileSystemEntryInfo"/> instances.</description>
      ///    </item>
      ///    <item>
      ///       <term><see cref="FileSystemInfo"/></term>
      ///       <description>This method will return instances of <see cref="DirectoryInfo"/> and <see cref="FileInfo"/> instances.</description>
      ///    </item>
      ///    <item>
      ///       <term><see cref="string"/></term>
      ///       <description>This method will return the full path of each item.</description>
      ///    </item>
      /// </list>
      /// </typeparam>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">検索するディレクトリ。</param>
      /// <param name="searchPattern">
      ///   ディレクトリ名と照合する検索文字列。対象: <paramref name="path"/>.
      ///   このパラメータには、有効なリテラルパスとワイルドカードの組み合わせを含めることができますが、
      ///   (<see cref="Path.WildcardStarMatchAll"/> and <see cref="Path.WildcardQuestion"/>) characters, but does not support regular expressions.
      /// </param>
      /// <param name="filters">処理で使用するカスタムフィルタの指定。</param>
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Infos")]
      [SecurityCritical]
      [Obsolete("Argument searchPattern is obsolete. The DirectoryEnumerationFilters argument provides better filter criteria.")]
      public static IEnumerable<T> EnumerateFileSystemEntryInfosTransacted<T>(KernelTransaction transaction, string path, string searchPattern, DirectoryEnumerationFilters filters)
      {
         return EnumerateFileSystemEntryInfosCore<T>(null, transaction, path, searchPattern, null, null, filters, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] の列挙可能なコレクションを返します。 file system entries that match a <paramref name="searchPattern"/> 指定されたパス内の</summary>
      /// <returns>一致するファイルシステムエントリ。項目の型は <typeparamref name="T"/> によって決定されます。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <typeparam name="T">The type to return. This may be one of the following types:
      ///    <list type="definition">
      ///    <item>
      ///       <term><see cref="FileSystemEntryInfo"/></term>
      ///       <description>This method will return instances of <see cref="FileSystemEntryInfo"/> instances.</description>
      ///    </item>
      ///    <item>
      ///       <term><see cref="FileSystemInfo"/></term>
      ///       <description>This method will return instances of <see cref="DirectoryInfo"/> and <see cref="FileInfo"/> instances.</description>
      ///    </item>
      ///    <item>
      ///       <term><see cref="string"/></term>
      ///       <description>This method will return the full path of each item.</description>
      ///    </item>
      /// </list>
      /// </typeparam>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">検索するディレクトリ。</param>
      /// <param name="searchPattern">
      ///   ディレクトリ名と照合する検索文字列。対象: <paramref name="path"/>.
      ///   このパラメータには、有効なリテラルパスとワイルドカードの組み合わせを含めることができますが、
      ///   (<see cref="Path.WildcardStarMatchAll"/> and <see cref="Path.WildcardQuestion"/>) characters, but does not support regular expressions.
      /// </param>
      /// <param name="filters">処理で使用するカスタムフィルタの指定。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Infos")]
      [SecurityCritical]
      [Obsolete("Argument searchPattern is obsolete. The DirectoryEnumerationFilters argument provides better filter criteria.")]
      public static IEnumerable<T> EnumerateFileSystemEntryInfosTransacted<T>(KernelTransaction transaction, string path, string searchPattern, DirectoryEnumerationFilters filters, PathFormat pathFormat)
      {
         return EnumerateFileSystemEntryInfosCore<T>(null, transaction, path, searchPattern, null, null, filters, pathFormat);
      }


      /// <summary>[AlphaFS] の列挙可能なコレクションを返します。 file system entries that match a <paramref name="searchPattern"/> 指定されたパス内の</summary>
      /// <returns>一致するファイルシステムエントリ。項目の型は <typeparamref name="T"/> によって決定されます。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <typeparam name="T">The type to return. This may be one of the following types:
      ///    <list type="definition">
      ///    <item>
      ///       <term><see cref="FileSystemEntryInfo"/></term>
      ///       <description>This method will return instances of <see cref="FileSystemEntryInfo"/> instances.</description>
      ///    </item>
      ///    <item>
      ///       <term><see cref="FileSystemInfo"/></term>
      ///       <description>This method will return instances of <see cref="DirectoryInfo"/> and <see cref="FileInfo"/> instances.</description>
      ///    </item>
      ///    <item>
      ///       <term><see cref="string"/></term>
      ///       <description>This method will return the full path of each item.</description>
      ///    </item>
      /// </list>
      /// </typeparam>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">検索するディレクトリ。</param>
      /// <param name="searchPattern">
      ///   ディレクトリ名と照合する検索文字列。対象: <paramref name="path"/>.
      ///   このパラメータには、有効なリテラルパスとワイルドカードの組み合わせを含めることができますが、
      ///   (<see cref="Path.WildcardStarMatchAll"/> and <see cref="Path.WildcardQuestion"/>) characters, but does not support regular expressions.
      /// </param>
      /// <param name="options">ディレクトリの列挙方法を指定する <see cref="DirectoryEnumerationOptions"/> フラグ。</param>
      /// <param name="filters">処理で使用するカスタムフィルタの指定。</param>
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Infos")]
      [SecurityCritical]
      [Obsolete("Argument searchPattern is obsolete. The DirectoryEnumerationFilters argument provides better filter criteria.")]
      public static IEnumerable<T> EnumerateFileSystemEntryInfosTransacted<T>(KernelTransaction transaction, string path, string searchPattern, DirectoryEnumerationOptions options, DirectoryEnumerationFilters filters)
      {
         return EnumerateFileSystemEntryInfosCore<T>(null, transaction, path, searchPattern, null, options,  filters, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] の列挙可能なコレクションを返します。 file system entries that match a <paramref name="searchPattern"/> 指定されたパス内の</summary>
      /// <returns>一致するファイルシステムエントリ。項目の型は <typeparamref name="T"/> によって決定されます。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <typeparam name="T">The type to return. This may be one of the following types:
      ///    <list type="definition">
      ///    <item>
      ///       <term><see cref="FileSystemEntryInfo"/></term>
      ///       <description>This method will return instances of <see cref="FileSystemEntryInfo"/> instances.</description>
      ///    </item>
      ///    <item>
      ///       <term><see cref="FileSystemInfo"/></term>
      ///       <description>This method will return instances of <see cref="DirectoryInfo"/> and <see cref="FileInfo"/> instances.</description>
      ///    </item>
      ///    <item>
      ///       <term><see cref="string"/></term>
      ///       <description>This method will return the full path of each item.</description>
      ///    </item>
      /// </list>
      /// </typeparam>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">検索するディレクトリ。</param>
      /// <param name="searchPattern">
      ///   ディレクトリ名と照合する検索文字列。対象: <paramref name="path"/>.
      ///   このパラメータには、有効なリテラルパスとワイルドカードの組み合わせを含めることができますが、
      ///   (<see cref="Path.WildcardStarMatchAll"/> and <see cref="Path.WildcardQuestion"/>) characters, but does not support regular expressions.
      /// </param>
      /// <param name="options">ディレクトリの列挙方法を指定する <see cref="DirectoryEnumerationOptions"/> フラグ。</param>
      /// <param name="filters">処理で使用するカスタムフィルタの指定。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Infos")]
      [SecurityCritical]
      [Obsolete("Argument searchPattern is obsolete. The DirectoryEnumerationFilters argument provides better filter criteria.")]
      public static IEnumerable<T> EnumerateFileSystemEntryInfosTransacted<T>(KernelTransaction transaction, string path, string searchPattern, DirectoryEnumerationOptions options, DirectoryEnumerationFilters filters, PathFormat pathFormat)
      {
         return EnumerateFileSystemEntryInfosCore<T>(null, transaction, path, searchPattern, null, options,  filters, pathFormat);
      }
   }
}
