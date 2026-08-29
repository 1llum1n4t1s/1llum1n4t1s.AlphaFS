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
   public static partial class Directory
   {
      /// <summary>Copy/move a Non-/Transacted file or directory including its children to a new location, <see cref="CopyOptions"/> or <see cref="MoveOptions"/> can be specified,
      /// and the possibility of notifying the application of its progress through a callback function.
      /// </summary>
      /// <returns>A <see cref="CopyMoveResult"/> class with the status of the Copy or Move action.</returns>
      /// <remarks>
      ///   <para>Option <see cref="CopyOptions.NoBuffering"/> is recommended for very large file transfers.</para>
      ///   <para>You cannot use the Move method to overwrite an existing file, unless <paramref name="cma.moveOptions"/> contains <see cref="MoveOptions.ReplaceExisting"/>.</para>
      ///   <para>Note that if you attempt to replace a file by moving a file of the same name into that directory, you get an IOException.</para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <exception cref="PlatformNotSupportedException">The operating system is older than Windows Vista.</exception>
      [SecurityCritical]
      internal static CopyMoveResult CopyMoveCore(CopyMoveArguments cma)
      {
         #region Setup
         
         var fsei = File.GetFileSystemEntryInfoCore(cma.Transaction, false, cma.SourcePath, true, cma.PathFormat);

         var isFolder = null == fsei || fsei.IsDirectory;

         // Directory.Move はファイルとフォルダの両方に適用可能。

         cma = File.ValidateFileOrDirectoryMoveArguments(cma, false, isFolder);


         if (isFolder && cma.IsCopy && IsSameOrDescendantDirectory(cma.SourcePathLp, cma.DestinationPathLp))
         {
            throw new IOException("コピー先ディレクトリをコピー元ディレクトリ自身、またはその配下にすることはできません。");
         }


         var copyMoveResult = new CopyMoveResult(cma, isFolder);

         var errorFilter = null != cma.DirectoryEnumerationFilters && null != cma.DirectoryEnumerationFilters.ErrorFilter ? cma.DirectoryEnumerationFilters.ErrorFilter : null;

         var retry = null != errorFilter && (cma.DirectoryEnumerationFilters.ErrorRetry > 0 || cma.DirectoryEnumerationFilters.ErrorRetryTimeout > 0);

         if (retry)
         {
            if (cma.DirectoryEnumerationFilters.ErrorRetry <= 0)
            {
               cma.DirectoryEnumerationFilters.ErrorRetry = 2;
            }

            if (cma.DirectoryEnumerationFilters.ErrorRetryTimeout <= 0)
            {
               cma.DirectoryEnumerationFilters.ErrorRetryTimeout = 10;
            }
         }


         // 実行中のStopwatchでstartを呼び出しても何も起こらない。
         copyMoveResult.Stopwatch.Start();

         #endregion // Setup


         var replacedDestinationPath = PrepareDirectoryReplacement(cma, isFolder);


         try
         {
            if (cma.IsCopy)
            {
               var sourceEntryInfo = isFolder
                  ? File.GetFileSystemEntryInfoCore(cma.Transaction, true, cma.SourcePathLp, false, PathFormat.LongFullPath)
                  : null;

               // ルート自体がジャンクションの場合も、リンク先を列挙せずジャンクションを複製する。
               if (null != sourceEntryInfo && sourceEntryInfo.IsMountPoint)
               {
                  var linkTargetInfo = File.GetLinkTargetInfoCore(cma.Transaction, cma.SourcePathLp, false, PathFormat.LongFullPath);
                  var linkTargetPath = Path.GetRegularPathCore(linkTargetInfo.SubstituteName, GetFullPathOptions.RemoveTrailingDirectorySeparator, false);

                  CreateJunctionCore(cma.Transaction, cma.DestinationPathLp, linkTargetPath, false, false, PathFormat.LongFullPath);
                  copyMoveResult.TotalFolders = 1;
               }

               // フォルダのシンボリックリンクをコピーする。
               // CopyFileEx() では実行できないため、エミュレートする。

               else if (File.HasCopySymbolicLink(cma.CopyOptions) && null != sourceEntryInfo && sourceEntryInfo.IsSymbolicLink)
               {
                  var lvi = File.GetLinkTargetInfoCore(cma.Transaction, cma.SourcePathLp, true, PathFormat.LongFullPath);

                  if (null != lvi)
                  {
                     File.CreateSymbolicLinkCore(cma.Transaction, cma.DestinationPathLp, lvi.SubstituteName, SymbolicLinkTarget.Directory, PathFormat.LongFullPath);

                     copyMoveResult.TotalFolders = 1;
                  }
               }

               else
               {
                  if (isFolder)
                  {
                     CopyMoveDirectoryCore(retry, cma, copyMoveResult);
                  }

                  else
                  {
                     File.CopyMoveCore(retry, cma, true, false, cma.SourcePathLp, cma.DestinationPathLp, copyMoveResult);
                  }
               }
            }


            // Move

            else
            {
               // ファイルまたはディレクトリとその子要素を移動します。
               // 既存のディレクトリとその子要素を新しいディレクトリにコピーします。

               File.CopyMoveCore(retry, cma, true, isFolder, cma.SourcePathLp, cma.DestinationPathLp, copyMoveResult);


               // 同じドライブ上で移動が行われた場合、ファイル/フォルダの数は不明。
               // ただし、1つのフォルダが正常に移動されたことは分かっている。

               if (copyMoveResult.ErrorCode == Win32Errors.NO_ERROR && isFolder)
               {
                  copyMoveResult.TotalFolders = 1;
               }
            }
         }
         catch (Exception moveException)
         {
            try
            {
               RestoreReplacedDirectory(cma, replacedDestinationPath);
            }
            catch (Exception restoreException)
            {
               throw new AggregateException("ディレクトリの移動と、置換先ディレクトリの復元の両方に失敗しました。", moveException, restoreException);
            }

            throw;
         }


         if (null != replacedDestinationPath)
         {
            if (copyMoveResult.IsCanceled || copyMoveResult.ErrorCode != Win32Errors.NO_ERROR)
            {
               RestoreReplacedDirectory(cma, replacedDestinationPath);
            }
            else
            {
               DeleteDirectoryCore(cma.Transaction, null, replacedDestinationPath, true, true, true, PathFormat.LongFullPath);
            }
         }


         copyMoveResult.Stopwatch.Stop();

         return copyMoveResult;
      }


      private static string PrepareDirectoryReplacement(CopyMoveArguments cma, bool isFolder)
      {
         if (!isFolder || cma.DelayUntilReboot || !File.ExistsCore(cma.Transaction, true, cma.DestinationPathLp, PathFormat.LongFullPath))
         {
            return null;
         }

         if (!File.HasReplaceExisting(cma.MoveOptions))
         {
            if (cma.EmulateMove)
            {
               throw new AlreadyExistsException(cma.DestinationPathLp, true);
            }

            return null;
         }

         string replacementPath;
         // 置換先を先に削除せず、同じ親の一時名へ退避する。ソース移動に失敗した場合は元へ戻す。
         do
         {
            replacementPath = cma.DestinationPathLp + ".alphafs-" + Path.GetRandomFileName();
         }
         while (File.ExistsCore(cma.Transaction, true, replacementPath, PathFormat.LongFullPath) ||
                File.ExistsCore(cma.Transaction, false, replacementPath, PathFormat.LongFullPath));

         MoveDirectoryWithoutOverwrite(cma.Transaction, cma.DestinationPathLp, replacementPath);
         return replacementPath;
      }


      private static void RestoreReplacedDirectory(CopyMoveArguments cma, string replacedDestinationPath)
      {
         if (null == replacedDestinationPath ||
             !File.ExistsCore(cma.Transaction, true, replacedDestinationPath, PathFormat.LongFullPath))
         {
            return;
         }

         // エミュレート移動ではソース削除を全コピー成功後まで遅延しているため、
         // 失敗途中のコピー先を破棄しても元データは失われない。
         if (File.ExistsCore(cma.Transaction, true, cma.DestinationPathLp, PathFormat.LongFullPath))
         {
            DeleteDirectoryCore(cma.Transaction, null, cma.DestinationPathLp, true, true, true, PathFormat.LongFullPath);
         }

         MoveDirectoryWithoutOverwrite(cma.Transaction, replacedDestinationPath, cma.DestinationPathLp);
      }


      private static bool IsSameOrDescendantDirectory(string sourcePath, string destinationPath)
      {
         var sourcePathWithoutSeparator = sourcePath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
         var destinationPathWithoutSeparator = destinationPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

         return destinationPathWithoutSeparator.Equals(sourcePathWithoutSeparator, StringComparison.OrdinalIgnoreCase) ||
                destinationPathWithoutSeparator.StartsWith(sourcePathWithoutSeparator + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase);
      }


      private static void MoveDirectoryWithoutOverwrite(KernelTransaction transaction, string sourcePath, string destinationPath)
      {
         var moveArguments = new CopyMoveArguments
         {
            Transaction = transaction,
            SourcePath = sourcePath,
            DestinationPath = destinationPath,
            MoveOptions = MoveOptions.None,
            PathFormat = PathFormat.LongFullPath
         };

         File.CopyMoveCore(false, moveArguments, true, true, sourcePath, destinationPath, null);
      }
   }
}
