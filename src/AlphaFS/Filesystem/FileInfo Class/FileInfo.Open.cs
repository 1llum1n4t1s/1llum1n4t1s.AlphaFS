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

using System.IO;
using System.Security;
using System.Security.AccessControl;

namespace Alphaleonis.Win32.Filesystem
{
   partial class FileInfo
   {
      #region .NET

      /// <summary>指定されたモードでファイルを開きます。</summary>
      /// <returns>指定されたモードで開かれた読み取り/書き込みアクセスの非共有 <see cref="FileStream"/> ファイル。</returns>
      /// <param name="mode">ファイルを開くモード（Open や Append など）を指定する <see cref="FileMode"/> 定数。</param>
      [SecurityCritical]
      public FileStream Open(FileMode mode)
      {
         return File.OpenCore(Transaction, LongFullName, mode, FileAccess.Read, FileShare.None, ExtendedFileAttributes.Normal, null, null, PathFormat.LongFullPath);
      }


      /// <summary>読み取り、書き込み、または読み取り/書き込みアクセスで指定されたモードでファイルを開きます。</summary>
      /// <returns>指定されたモードとアクセスで開かれた非共有 <see cref="FileStream"/> オブジェクト。</returns>
      /// <param name="mode">ファイルを開くモード（Open や Append など）を指定する <see cref="FileMode"/> 定数。</param>
      /// <param name="access">Read、Write、または ReadWrite のいずれのファイルアクセスでファイルを開くかを指定する <see cref="FileAccess"/> 定数。</param>
      [SecurityCritical]
      public FileStream Open(FileMode mode, FileAccess access)
      {
         return File.OpenCore(Transaction, LongFullName, mode, access, FileShare.None, ExtendedFileAttributes.Normal, null, null, PathFormat.LongFullPath);
      }


      /// <summary>読み取り、書き込み、または読み取り/書き込みアクセスと指定された共有オプションで指定されたモードでファイルを開きます。</summary>
      /// <returns>指定されたモード、アクセス、および共有オプションで開かれた <see cref="FileStream"/> オブジェクト。</returns>
      /// <param name="mode">ファイルを開くモード（Open や Append など）を指定する <see cref="FileMode"/> 定数。</param>
      /// <param name="access">Read、Write、または ReadWrite のいずれのファイルアクセスでファイルを開くかを指定する <see cref="FileAccess"/> 定数。</param>
      /// <param name="share">他の <see cref="FileStream"/> オブジェクトがこのファイルに持つアクセスの種類を指定する <see cref="FileShare"/> 定数。</param>
      [SecurityCritical]
      public FileStream Open(FileMode mode, FileAccess access, FileShare share)
      {
         return File.OpenCore(Transaction, LongFullName, mode, access, share, ExtendedFileAttributes.Normal, null, null, PathFormat.LongFullPath);
      }

      #endregion // .NET


      /// <summary>[AlphaFS] 読み取り、書き込み、または読み取り/書き込みアクセスで指定されたモードでファイルを開きます。</summary>
      /// <returns>指定されたモードとアクセスで開かれた非共有 <see cref="FileStream"/> オブジェクト。</returns>
      /// <param name="mode">ファイルを開くモード（Open や Append など）を指定する <see cref="FileMode"/> 定数。</param>
      /// <param name="rights">ファイルが存在しない場合にファイルを作成するかどうかを指定し、既存のファイルの内容を保持するか上書きするかを追加オプションと共に決定する <see cref="FileSystemRights"/> 値。</param>
      [SecurityCritical]
      public FileStream Open(FileMode mode, FileSystemRights rights)
      {
         return File.OpenCore(Transaction, LongFullName, mode, rights, FileShare.None, ExtendedFileAttributes.Normal, null, null, PathFormat.LongFullPath);
      }


      /// <summary>[AlphaFS] 読み取り、書き込み、または読み取り/書き込みアクセスと指定された共有オプションで指定されたモードでファイルを開きます。</summary>
      /// <returns>指定されたモード、アクセス、および共有オプションで開かれた <see cref="FileStream"/> オブジェクト。</returns>
      /// <param name="mode">ファイルを開くモード（Open や Append など）を指定する <see cref="FileMode"/> 定数。</param>
      /// <param name="rights">ファイルが存在しない場合にファイルを作成するかどうかを指定し、既存のファイルの内容を保持するか上書きするかを追加オプションと共に決定する <see cref="FileSystemRights"/> 値。</param>
      /// <param name="share">他の <see cref="FileStream"/> オブジェクトがこのファイルに持つアクセスの種類を指定する <see cref="FileShare"/> 定数。</param>
      [SecurityCritical]
      public FileStream Open(FileMode mode, FileSystemRights rights, FileShare share)
      {
         return File.OpenCore(Transaction, LongFullName, mode, rights, share, ExtendedFileAttributes.Normal, null, null, PathFormat.LongFullPath);
      }
   }
}
