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
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Security;
using System.Threading;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class File
   {
      // Symbolic Link Effects on File Systems Functions: https://msdn.microsoft.com/en-us/library/windows/desktop/aa365682(v=vs.85).aspx


      /// <summary>非トランザクション/トランザクションファイルまたはディレクトリとその子要素を新しい場所にコピー/移動します。 <see cref="CopyOptions"/>または<see cref="MoveOptions"/>を指定でき、
      /// およびコールバック関数を通じてアプリケーショ��に進行状況を通知する可能性があります。
      /// </summary>
      /// <remarks>
      ///   <para>非常に大きなファイル転送には<see cref="CopyOptions.NoBuffering"/>オプションが推奨されます。</para>
      ///   <para>You cannot use the Move method to overwrite an existing file, unless
      ///   <paramref name="cma.MoveOptions"/> contains <see cref="MoveOptions.ReplaceExisting"/>.</para>
      ///   <para>This Move method works across disk volumes, and it does not throw an exception if the
      ///   ソースとコピー先が同じ場合にも同様です。</para>
      ///   <para>Note that if you attempt to replace a file by moving a file of the same name into
      ///   that directory, you get an IOException.</para>
      /// </remarks>
      /// <returns>コピーまたは移動操作のステータスを含む<see cref="CopyMoveResult"/>クラスを返します。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="FileNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      [SuppressMessage("Microsoft.Maintainability", "CA1502:AvoidExcessiveComplexity")]
      [SecurityCritical]
      internal static CopyMoveResult CopyMoveCore(bool retry, CopyMoveArguments cma, bool driveChecked, bool isFolder, string sourceFilePath, string destinationFilePath, CopyMoveResult copyMoveResult)
      {
         #region Setup

         cma = ValidateFileOrDirectoryMoveArguments(cma, driveChecked, false, sourceFilePath, destinationFilePath, out sourceFilePath, out destinationFilePath);


         var copyMoveRes = copyMoveResult ?? new CopyMoveResult(cma, isFolder, sourceFilePath, destinationFilePath);

         var isSingleFileAction = null == copyMoveResult && copyMoveRes.IsFile;


         var attempts = 1;

         var retryTimeout = 0;

         var errorFilter = null != cma.DirectoryEnumerationFilters && null != cma.DirectoryEnumerationFilters.ErrorFilter ? cma.DirectoryEnumerationFilters.ErrorFilter : null;

         if (retry)
         {
            if (null != errorFilter)
            {
               attempts += cma.DirectoryEnumerationFilters.ErrorRetry;

               retryTimeout = cma.DirectoryEnumerationFilters.ErrorRetryTimeout;
            }

            else
            {
               if (cma.Retry <= 0)
               {
                  cma.Retry = 2;
               }

               if (cma.RetryTimeout <= 0)
               {
                  cma.RetryTimeout = 10;
               }

               attempts += cma.Retry;

               retryTimeout = cma.RetryTimeout;
            }
         }


         // 実行中のStopwatchでstartを呼び出しても何も起こらない。
         copyMoveRes.Stopwatch.Start();

         #endregion // Setup


         while (attempts-- > 0)
         {
            // MSDN: コピー/移動操作中にこのフラグがTRUEに設定されると、操作はキャンセルされます。
            // それ以外の場合、コピー/移動操作は完了まで続行されます。

            copyMoveRes.ErrorCode = (int) Win32Errors.NO_ERROR;

            copyMoveRes.IsCanceled = false;


            if (!cma.DelayUntilReboot)
            {
               // Ensure the file's parent directory exists.

               var parentFolder = Directory.GetParentCore(cma.Transaction, destinationFilePath, PathFormat.LongFullPath);

               if (null != parentFolder)
               {
                  parentFolder.Create();
               }
            }


            if (CopyMoveNative(cma, !cma.IsCopy, sourceFilePath, destinationFilePath, out var cancel, out var lastError))
            {
               // We take an extra hit by getting the file size for a single file Copy or Move action.

               if (isSingleFileAction)

               {
                  copyMoveRes.TotalBytes = GetSizeCore(null, cma.Transaction, destinationFilePath, true, PathFormat.LongFullPath);
               }


               if (!isFolder)
               {
                  copyMoveRes.TotalFiles++;

                  // Only set timestamps for files.

                  if (cma.CopyTimestamps)

                  {
                     CopyTimestampsCore(cma.Transaction, false, sourceFilePath, destinationFilePath, false, PathFormat.LongFullPath);
                  }
               }
               
               break;
            }


            // The Copy/Move action failed or is canceled.

            copyMoveRes.ErrorCode = lastError;

            copyMoveRes.IsCanceled = cancel;

            
            // 呼び出し元に例外を報告する。
            if (null != errorFilter)
            {
               var continueCopyMove = errorFilter(lastError, new Win32Exception(lastError).Message, Path.GetCleanExceptionPath(destinationFilePath));

               if (!continueCopyMove)
               {
                  copyMoveRes.IsCanceled = true;
                  break;
               }
            }


            if (!cancel)
            {
               if (retry)
               {
                  copyMoveRes.Retries++;
               }

               retry = attempts > 0 && retryTimeout > 0;


               // Remove any read-only/hidden attribute, which might also fail.

               RestartMoveOrThrowException(retry, lastError, isFolder, !cma.IsCopy, cma, sourceFilePath, destinationFilePath);

               if (retry)
               {
                  if (null != errorFilter && cma.DirectoryEnumerationFilters.CancellationToken.CanBeCanceled)
                  {
                     if (cma.DirectoryEnumerationFilters.CancellationToken.WaitHandle.WaitOne(retryTimeout * 1000))
                     {
                        copyMoveRes.IsCanceled = true;
                        break;
                     }
                  }

                  else
                  {
                     System.Threading.Thread.Sleep(retryTimeout * 1000);
                  }
               }
            }
         }

         
         if (isSingleFileAction)
         {
            copyMoveRes.Stopwatch.Stop();
         }

         copyMoveResult = copyMoveRes;

         return copyMoveResult;
      }
   }
}
