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
using System.Security;

namespace Alphaleonis.Win32.Filesystem
{
   public sealed partial class DirectoryInfo
   {
      #region .NET

      /// <summary>この <see cref="DirectoryInfo"/> が空の場合、削除します。</summary>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      [SecurityCritical]
      public override void Delete()
      {
         Directory.DeleteDirectoryCore(Transaction, EntryInfo, null, false, false, false, PathFormat.LongFullPath);

         // System.IO は自身が行った変更でキャッシュ済みの状態を無効化する。同じ挙動にそろえる。
         Refresh();
      }


      /// <summary>サブディレクトリおよびファイルを削除するかどうかを指定して、この <see cref="DirectoryInfo"/> インスタンスを削除します。</summary>
      /// <remarks>
      ///   <para><see cref="DirectoryInfo"/> にファイルもサブディレクトリもない場合、recursive が <c>false</c> でもこのメソッドは <see cref="DirectoryInfo"/> を削除します。</para>
      ///   <para>recursive が false のときに空でない <see cref="DirectoryInfo"/> を削除しようとすると、<see cref="IOException"/> がスローされます。</para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="recursive">このディレクトリ、そのサブディレクトリ、およびすべてのファイルを削除する場合は <c>true</c>、それ以外の場合は <c>false</c>。</param>
      [SecurityCritical]
      public void Delete(bool recursive)
      {
         Directory.DeleteDirectoryCore(Transaction, EntryInfo, null, recursive, false, false, PathFormat.LongFullPath);

         // System.IO は自身が行った変更でキャッシュ済みの状態を無効化する。同じ挙動にそろえる。
         Refresh();
      }

      #endregion // .NET


      /// <summary>[AlphaFS] ファイルおよびサブディレクトリを削除するかどうかを指定して、この <see cref="DirectoryInfo"/> インスタンスを削除します。</summary>
      /// <remarks>
      ///   <para><see cref="DirectoryInfo"/> にファイルもサブディレクトリもない場合、recursive が <c>false</c> でもこのメソッドは <see cref="DirectoryInfo"/> を削除します。</para>
      ///   <para>recursive が false のときに空でない <see cref="DirectoryInfo"/> を削除しようとすると、<see cref="IOException"/> がスローされます。</para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="recursive">このディレクトリ、そのサブディレクトリ、およびすべてのファイルを削除する場合は <c>true</c>、それ以外の場合は <c>false</c>。</param>
      /// <param name="ignoreReadOnly"><c>true</c> の場合、ファイルおよびディレクトリの読み取り専用属性を無視します。</param>
      [SecurityCritical]
      public void Delete(bool recursive, bool ignoreReadOnly)
      {
         Directory.DeleteDirectoryCore(Transaction, EntryInfo, null, recursive, ignoreReadOnly, false, PathFormat.LongFullPath);

         // System.IO は自身が行った変更でキャッシュ済みの状態を無効化する。同じ挙動にそろえる。
         Refresh();
      }


      /// <summary>[AlphaFS] ファイルおよびサブディレクトリを削除するかどうかを指定して、この <see cref="DirectoryInfo"/> インスタンスを削除します。</summary>
      /// <remarks>
      ///   <para><see cref="DirectoryInfo"/> にファイルもサブディレクトリもない場合、recursive が <c>false</c> でもこのメソッドは <see cref="DirectoryInfo"/> を削除します。</para>
      ///   <para>recursive が false のときに空でない <see cref="DirectoryInfo"/> を削除しようとすると、<see cref="IOException"/> がスローされます。</para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="recursive">このディレクトリ、そのサブディレクトリ、およびすべてのファイルを削除する場合は <c>true</c>、それ以外の場合は <c>false</c>。</param>
      /// <param name="ignoreReadOnly"><c>true</c> の場合、ファイルおよびディレクトリの読み取り専用属性を無視します。</param>
      /// <param name="continueOnNotFound"><c>true</c> の場合、ディレクトリが存在しないときに <see cref="DirectoryNotFoundException"/> をスローしません。</param>
      [SecurityCritical]
      public void Delete(bool recursive, bool ignoreReadOnly, bool continueOnNotFound)
      {
         Directory.DeleteDirectoryCore(Transaction, EntryInfo, null, recursive, ignoreReadOnly, continueOnNotFound, PathFormat.LongFullPath);

         // System.IO は自身が行った変更でキャッシュ済みの状態を無効化する。同じ挙動にそろえる。
         Refresh();
      }
   }
}
