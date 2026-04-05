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

using Microsoft.Win32.SafeHandles;
using System.IO;
using System.Security.AccessControl;
using FileStream = System.IO.FileStream;
using System.Diagnostics.CodeAnalysis;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class File
   {
      /// <summary>指定されたパスで、指定されたモード、読み取り/書き込みまたは読み書きアクセス、 read/write access, the specified sharing option and additional options specified.</summary>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">開くファイル。</param>
      /// <param name="mode">ファイルが存在しない場合に作成するかどうか、および既存ファイルの内容を保持するか上書きするかを指定する<see cref="FileMode"/>値。</param>
      /// <param name="access">ファイルに対して実行できる操作を指定する<see cref="FileAccess"/>値。</param>
      /// <param name="share">他のスレッドがファイルに対して持つアクセスの種類を指定する<see cref="FileShare"/>値。</param>
      /// <param name="attributes">Advanced <see cref="ExtendedFileAttributes"/> options for this file.</param>
      /// <param name="bufferSize">0より大きい正の<see cref="System.Int32"/>値でバッファサイズを示します。 デフォルトのバッファサイズは4096です。</param>
      /// <param name="security">The security.</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      /// <returns>
      ///   <para>A <see cref="FileStream"/> instance on the specified path, having the specified mode with</para>
      ///   <para>read, write, or read/write および指定された共有オプションの<see cref="FileStream"/>。</para>
      /// </returns>
      internal static FileStream OpenCore(KernelTransaction transaction, string path, FileMode mode, FileAccess access, FileShare share, ExtendedFileAttributes attributes, int? bufferSize, FileSecurity security, PathFormat pathFormat)
      {
         var rights = access == FileAccess.Read ? FileSystemRights.Read : (access == FileAccess.Write ? FileSystemRights.Write : FileSystemRights.Read | FileSystemRights.Write);

         return OpenCore(transaction, path, mode, rights, share, attributes, bufferSize, security, pathFormat);
      }


      /// <summary>指定されたパスで、指定されたモード、読み取り/書き込みまたは読み書きアクセス、 read/write access, the specified sharing option and additional options specified.</summary>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">開くファイル。</param>
      /// <param name="mode">ファイルが存在しない場合に作成するかどうか、および既存ファイルの内容を保持するか上書きするかを指定する<see cref="FileMode"/>値。</param>
      /// <param name="rights">A <see cref="FileSystemRights"/> value that specifies whether a file is created if one does not exist, and determines whether the contents of existing files are retained or overwritten along with additional options.</param>
      /// <param name="share">他のスレッドがファイルに対して持つアクセスの種類を指定する<see cref="FileShare"/>値。</param>
      /// <param name="attributes">Advanced <see cref="ExtendedFileAttributes"/> options for this file.</param>
      /// <param name="bufferSize">0より大きい正の<see cref="System.Int32"/>値でバッファサイズを示します。 デフォルトのバッファサイズは4096です。</param>
      /// <param name="security">The security.</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      /// <returns>
      ///   <para>A <see cref="FileStream"/> instance on the specified path, having the specified mode with</para>
      ///   <para>read, write, or read/write および指定された共有オプションの<see cref="FileStream"/>。</para>
      /// </returns>
      [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope")]
      internal static FileStream OpenCore(KernelTransaction transaction, string path, FileMode mode, FileSystemRights rights, FileShare share, ExtendedFileAttributes attributes, int? bufferSize, FileSecurity security, PathFormat pathFormat)
      {
         var access = ((rights & FileSystemRights.ReadData) != 0 ? FileAccess.Read : 0) |
                      ((rights & FileSystemRights.WriteData) != 0 || (rights & FileSystemRights.AppendData) != 0 ? FileAccess.Write : 0);


         SafeFileHandle safeHandle = null;

         try
         {
            safeHandle = CreateFileCore(transaction, false, path, attributes, security, mode, rights, share, true, false, pathFormat);

            return new FileStream(safeHandle, access, bufferSize ?? NativeMethods.DefaultFileBufferSize, (attributes & ExtendedFileAttributes.Overlapped) != 0);
         }
         catch
         {
            if (null != safeHandle && !safeHandle.IsClosed)
            {
               safeHandle.Close();
            }

            throw;
         }
      }
   }
}
