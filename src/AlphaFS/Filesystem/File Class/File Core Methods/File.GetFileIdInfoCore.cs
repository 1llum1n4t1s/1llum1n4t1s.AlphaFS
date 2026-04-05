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
using System.Runtime.InteropServices;
using System.Security;
using System.Security.AccessControl;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class File
   {
      /// <summary>Gets the unique identifier for a file. The identifier is composed of a 64-bit volume serial number and 128-bit file system entry identifier.</summary>
      /// <returns>要求された情報を含む<see cref="FileIdInfo"/>インスタンス。</returns>
      /// <remarks>ファイルIDは時間の経過とともに一意であることが保証さ���ていません。ファイルシステムはそれらを再利用できるためです。場合によっては、ファイルのファイルIDが時間の経過とともに変更されることがあります。</remarks>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="isFolder"><paramref name="path"/>がファイルかディレクトリかを指定します。</param>
      /// <param name="path">ファイルへのパス。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      internal static FileIdInfo GetFileIdInfoCore(KernelTransaction transaction, bool isFolder, string path, PathFormat pathFormat)
      {
         using var handle = CreateFileCore(transaction, isFolder, path, ExtendedFileAttributes.BackupSemantics, null, FileMode.Open, FileSystemRights.ReadData, FileShare.ReadWrite, true, false, pathFormat);
         if (NativeMethods.IsAtLeastWindows8)
         {
            // ReFS is supported.
            using var safeBuffer = new SafeGlobalMemoryBufferHandle(Marshal.SizeOf<NativeMethods.FILE_ID_INFO>());
            var success = NativeMethods.GetFileInformationByHandleEx(handle, NativeMethods.FILE_INFO_BY_HANDLE_CLASS.FILE_ID_INFO, safeBuffer, (uint) safeBuffer.Capacity);

            var lastError = Marshal.GetLastWin32Error();
            if (!success)
            {
               NativeError.ThrowException(lastError, isFolder, path);
            }

            return new FileIdInfo(safeBuffer.PtrToStructure<NativeMethods.FILE_ID_INFO>(0));
         }


         // Only NTFS is supported.
         return GetFileIdInfo(handle);
      }
   }
}
