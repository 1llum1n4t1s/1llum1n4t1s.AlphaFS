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

using System.Diagnostics.CodeAnalysis;
using System.Security;
using System.Security.AccessControl;

namespace Alphaleonis.Win32.Filesystem
{
   public sealed partial class DirectoryInfo
   {
      #region .NET

      /// <summary>ディレクトリを作成します。</summary>
      /// <remarks>ディレクトリが既に存在する場合、このメソッドは何もしません。</remarks>
      [SecurityCritical]
      public void Create()
      {
         Directory.CreateDirectoryCore(true, Transaction, LongFullName, null, null, false, PathFormat.LongFullPath);

         // System.IO は自身が行った変更でキャッシュ済みの状態を無効化する。同じ挙動にそろえる。
         Refresh();
      }


      /// <summary><see cref="DirectorySecurity"/> オブジェクトを使用してディレクトリを作成します。</summary>
      /// <param name="directorySecurity">ディレクトリに適用するアクセス制御。</param>
      /// <remarks>ディレクトリが既に存在する場合、このメソッドは何もしません。</remarks>
      [SuppressMessage("Microsoft.Design", "CA1011:ConsiderPassingBaseTypesAsParameters")]
      [SecurityCritical]
      public void Create(DirectorySecurity directorySecurity)
      {
         Directory.CreateDirectoryCore(true, Transaction, LongFullName, null, directorySecurity, false, PathFormat.LongFullPath);

         // System.IO は自身が行った変更でキャッシュ済みの状態を無効化する。同じ挙動にそろえる。
         Refresh();
      }

      #endregion // .NET


      /// <summary>[AlphaFS] ディレクトリを作成します。</summary>
      /// <param name="compress"><c>true</c> の場合、NTFS 圧縮を使用してディレクトリを圧縮します。</param>
      /// <remarks>ディレクトリが既に存在する場合、このメソッドは何もしません。</remarks>
      [SuppressMessage("Microsoft.Design", "CA1011:ConsiderPassingBaseTypesAsParameters")]
      [SecurityCritical]
      public DirectoryInfo Create(bool compress)
      {
         var directoryInfo = Directory.CreateDirectoryCore(true, Transaction, LongFullName, null, null, compress, PathFormat.LongFullPath);

         // System.IO は自身が行った変更でキャッシュ済みの状態を無効化する。同じ挙動にそろえる。
         Refresh();

         return directoryInfo;
      }


      /// <summary>[AlphaFS] <see cref="DirectorySecurity"/> オブジェクトを使用してディレクトリを作成します。</summary>
      /// <param name="directorySecurity">ディレクトリに適用するアクセス制御。</param>
      /// <param name="compress"><c>true</c> の場合、NTFS 圧縮を使用してディレクトリを圧縮します。</param>
      /// <remarks>ディレクトリが既に存在する場合、このメソッドは何もしません。</remarks>
      [SuppressMessage("Microsoft.Design", "CA1011:ConsiderPassingBaseTypesAsParameters")]
      [SecurityCritical]
      public DirectoryInfo Create(DirectorySecurity directorySecurity, bool compress)
      {
         var directoryInfo = Directory.CreateDirectoryCore(true, Transaction, LongFullName, null, directorySecurity, compress, PathFormat.LongFullPath);

         // System.IO は自身が行った変更でキャッシュ済みの状態を無効化する。同じ挙動にそろえる。
         Refresh();

         return directoryInfo;
      }
   }
}
