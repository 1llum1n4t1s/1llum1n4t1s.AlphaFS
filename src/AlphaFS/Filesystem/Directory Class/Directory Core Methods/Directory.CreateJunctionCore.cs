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
using System.Globalization;
using System.IO;
using System.Security;
using System.Security.AccessControl;
using Microsoft.Win32.SafeHandles;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class Directory
   {
      /// <summary>NTFSディレクトリジャンクションを作成します（CMDコマンド: "MKLINK /J" に類似）。同名のジャンクションポイントの上書きが許可されます。</summary>
      /// <returns>Returns the long path to the directory junction.</returns>
      /// <remarks>
      /// ディレクトリは空であり、ローカルボリュームに存在する必要があります。
      /// <paramref name="directoryPath"/>（ターゲット）のディレクトリの日付と時刻スタンプが the directory junction.
      /// <para>
      ///   MSDN: ジャンクション（ソフトリンクとも呼ばれる）は、参照するストレージオブジェクトが別個のディレクトリである点でハードリンクと異なり、
      ///   ジャンクションは同じコンピュータ上の異なるローカルボリュームにあるディレクトリをリンクできます。
      ///   それ以外の点では、ジャンクションはハードリンクと同様に動作します。ジャンクションはリパースポイントを通じて実装されます。
      /// </para>
      /// </remarks>
      /// <exception cref="AlreadyExistsException"/>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="junctionPath">作成するジャンクションポイントのパス。</param>
      /// <param name="directoryPath">The path to the directory. If the directory does not exist it will be created.</param>
      /// <param name="overwrite"><c>true</c> to overwrite an existing junction point. The directory is removed and recreated.</param>
      /// <param name="copyTargetTimestamps"><c>true</c> to copy the target date and time stamps to the directory junction.</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      internal static string CreateJunctionCore(KernelTransaction transaction, string junctionPath, string directoryPath, bool overwrite, bool copyTargetTimestamps, PathFormat pathFormat)
      {
         if (pathFormat != PathFormat.LongFullPath)
         {
            Path.CheckSupportedPathFormat(directoryPath, true, true);
            Path.CheckSupportedPathFormat(junctionPath, true, true);

            directoryPath = Path.GetExtendedLengthPathCore(transaction, directoryPath, pathFormat, GetFullPathOptions.RemoveTrailingDirectorySeparator);
            junctionPath = Path.GetExtendedLengthPathCore(transaction, junctionPath, pathFormat, GetFullPathOptions.RemoveTrailingDirectorySeparator);

            pathFormat = PathFormat.LongFullPath;
         }


         // Directory Junction logic.


         // Check if drive letter is a mapped network drive.
         if (new DriveInfo(directoryPath).IsUnc)
         {
            throw new ArgumentException(string.Format(CultureInfo.InvariantCulture, Resources.Network_Path_Not_Allowed, directoryPath), "directoryPath");
         }

         if (new DriveInfo(junctionPath).IsUnc)
         {
            throw new ArgumentException(string.Format(CultureInfo.InvariantCulture, Resources.Network_Path_Not_Allowed, junctionPath), "junctionPath");
         }


         // Check for existing file.
         File.ThrowIOExceptionIfFsoExist(transaction, false, directoryPath, pathFormat);
         File.ThrowIOExceptionIfFsoExist(transaction, false, junctionPath, pathFormat);


         // Check for existing directory junction folder.
         if (File.ExistsCore(transaction, true, junctionPath, pathFormat))
         {
            if (overwrite)
            {
               // 上書き対象は既存ジャンクションだけに限定する。リパースポイントとして検証してから
               // 解除することで、通常のディレクトリツリーを誤って削除しない。
               DeleteJunctionCore(transaction, null, junctionPath, true, pathFormat);
            }

            else
            {
               // Ensure the folder is empty.
               if (!IsEmptyCore(transaction, junctionPath, pathFormat))
               {
                  throw new DirectoryNotEmptyException(junctionPath, true);
               }

               throw new AlreadyExistsException(junctionPath, true);
            }
         }


         // 公開契約どおり、存在しないターゲットディレクトリを先に作成する。
         CreateDirectoryCore(true, transaction, directoryPath, null, null, false, pathFormat);

         // Create the folder and convert it to a directory junction.
         CreateDirectoryCore(true, transaction, junctionPath, null, null, false, pathFormat);

         using (var safeHandle = OpenDirectoryJunction(transaction, junctionPath, pathFormat))
            Device.CreateDirectoryJunction(safeHandle, directoryPath);


         // Copy the target date and time stamps to the directory junction.
         if (copyTargetTimestamps)
         {
            File.CopyTimestampsCore(transaction, true, directoryPath, junctionPath, true, pathFormat);
         }


         return junctionPath;
      }


      private static SafeFileHandle OpenDirectoryJunction(KernelTransaction transaction, string junctionPath, PathFormat pathFormat)
      {
         return File.CreateFileCore(transaction, true, junctionPath, ExtendedFileAttributes.BackupSemantics | ExtendedFileAttributes.OpenReparsePoint, null, FileMode.Open, FileSystemRights.WriteData, FileShare.ReadWrite, false, false, pathFormat);
      }
   }
}
