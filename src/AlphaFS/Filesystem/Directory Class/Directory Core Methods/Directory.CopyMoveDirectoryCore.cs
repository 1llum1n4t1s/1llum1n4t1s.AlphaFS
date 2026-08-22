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
using System.Security;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class Directory
   {
      [SecurityCritical]
      internal static void CopyMoveDirectoryCore(bool retry, CopyMoveArguments cma, CopyMoveResult copyMoveResult)
      {
         var dirs = new Queue<string>(NativeMethods.DefaultDirectoryQueueCapacity);

         if (!File.ExistsCore(cma.Transaction, true, cma.SourcePathLp, PathFormat.LongFullPath))
         {
            NativeError.ThrowException(Win32Errors.ERROR_PATH_NOT_FOUND, cma.SourcePathLp);
         }

         // 空ディレクトリやファイルから始まる列挙でも、コピー先のルートを先に確保する。
         CreateDirectoryCore(true, cma.Transaction, cma.DestinationPathLp, null, null, false, PathFormat.LongFullPath);

         dirs.Enqueue(cma.SourcePathLp);


         while (dirs.Count > 0)
         {
            var srcLp = dirs.Dequeue();

            var dstLp = srcLp.StartsWith(cma.SourcePathLp, StringComparison.OrdinalIgnoreCase)
               ? cma.DestinationPathLp + srcLp.Substring(cma.SourcePathLp.Length)
               : srcLp.ReplaceIgnoreCase(cma.SourcePathLp, cma.DestinationPathLp);
            

            // ソースフォルダを走査し、ファイルとフォルダを処理する。
            // 再帰は使用せず、代わりにQueueを使用する。

            foreach (var fseiSource in EnumerateFileSystemEntryInfosCore<FileSystemEntryInfo>(null, cma.Transaction, srcLp, Path.WildcardStarMatchAll, null, null, cma.DirectoryEnumerationFilters, PathFormat.LongFullPath))
            {
               var fseiSourcePath = fseiSource.LongFullPath;

               var fseiDestinationPath = Path.CombineCore(false, dstLp, fseiSource.FileName);

               if (fseiSource.IsDirectory)
               {
                  if (fseiSource.IsSymbolicLink && File.HasCopySymbolicLink(cma.CopyOptions))
                  {
                     var linkTargetInfo = File.GetLinkTargetInfoCore(cma.Transaction, fseiSourcePath, false, PathFormat.LongFullPath);

                     File.CreateSymbolicLinkCore(cma.Transaction, fseiDestinationPath, linkTargetInfo.SubstituteName, SymbolicLinkTarget.Directory, PathFormat.LongFullPath);

                     copyMoveResult.TotalFolders++;
                     continue;
                  }

                  CreateDirectoryCore(true, cma.Transaction, fseiDestinationPath, null, null, false, PathFormat.LongFullPath);

                  copyMoveResult.TotalFolders++;

                  dirs.Enqueue(fseiSourcePath);
               }

               else
               {
                  // ファイルカウントはFile.CopyMoveCoreメソッドで行われる。

                  File.CopyMoveCore(retry, cma, true, false, fseiSourcePath, fseiDestinationPath, copyMoveResult);

                  if (copyMoveResult.IsCanceled)
                  {
                     // whileループを中断。
                     dirs.Clear();

                     // foreachループを中断。
                     break;
                  }
                  

                  if (copyMoveResult.ErrorCode == Win32Errors.NO_ERROR)
                  {
                     copyMoveResult.TotalBytes += fseiSource.FileSize;

                     if (cma.EmulateMove)
                     {
                        File.DeleteFileCore(cma.Transaction, fseiSourcePath, true, fseiSource.Attributes, PathFormat.LongFullPath);
                     }
                  }
               }
            }
         }


         if (!copyMoveResult.IsCanceled && copyMoveResult.ErrorCode == Win32Errors.NO_ERROR)
         {
            if (cma.CopyTimestamps)
            {
               CopyFolderTimestamps(cma);
            }

            if (cma.EmulateMove)
            {
               DeleteDirectoryCore(cma.Transaction, null, cma.SourcePathLp, true, true, true, PathFormat.LongFullPath);
            }
         }
      }
   }
}
