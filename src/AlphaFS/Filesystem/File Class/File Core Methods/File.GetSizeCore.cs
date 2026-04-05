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
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Runtime.InteropServices;
using System.Security;
using System.Security.AccessControl;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class File
   {
      /// <summary>指定されたファイルのサイズを取得します。</summary>
      /// <returns>最初のストリームまたはすべてのストリームのファイルサイズ(バイト単位)。</returns>
      /// <remarks><paramref name="path"/>または<paramref name="safeFileHandle"/>のいずれかを使用し、両方は使用しないでください。</remarks>
      /// <param name="safeFileHandle">ファイルへの<see cref="SafeFileHandle"/>。</param>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">ファイルへのパス。</param>
      /// <param name="sizeOfAllStreams">すべての代替データストリームのサイズを取得する場合は<c>true</c>、最初のストリームのサイズを取得する場合は<c>false</c>。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope")]
      [SecurityCritical]
      internal static long GetSizeCore(SafeFileHandle safeFileHandle, KernelTransaction transaction, string path, bool sizeOfAllStreams, PathFormat pathFormat)
      {
         var pathLp = Path.GetExtendedLengthPathCore(transaction, path, pathFormat, GetFullPathOptions.RemoveTrailingDirectorySeparator | GetFullPathOptions.FullCheck);

         if (sizeOfAllStreams)
         {
            return FindAllStreamsCore(transaction, pathLp);
         }


         var callerHandle = null != safeFileHandle;

         if (!callerHandle)
         {
            safeFileHandle = CreateFileCore(transaction, false, pathLp, ExtendedFileAttributes.Normal, null, FileMode.Open, FileSystemRights.ReadData, FileShare.ReadWrite, true, false, PathFormat.LongFullPath);
         }

         long fileSize;

         try
         {
            var success = NativeMethods.GetFileSizeEx(safeFileHandle, out fileSize);

            var lastError = Marshal.GetLastWin32Error();

            if (!success && lastError != Win32Errors.ERROR_SUCCESS)
            {
               NativeError.ThrowException(lastError, path);
            }
         }
         finally
         {
            // Handle is ours, dispose.
            if (!callerHandle && null != safeFileHandle && !safeFileHandle.IsClosed)
            {
               safeFileHandle.Close();
            }
         }

         return fileSize;
      }
   }
}
