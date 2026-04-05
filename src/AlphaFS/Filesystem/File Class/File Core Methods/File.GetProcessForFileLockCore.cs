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
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Runtime.InteropServices;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class File
   {
      /// <summary>[AlphaFS] <paramref name="filePaths"/>で指定されたファイルにロックを持つプロセスのリストを取得します。</summary>
      /// <returns>
      /// <c>null</c> when no processes found that are locking the file(s) specified by <paramref name="filePaths"/>.
      /// <paramref name="filePaths"/>で指定されたファイルをロックしているプロセスのリスト。
      /// </returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="ArgumentOutOfRangeException"/>
      /// <exception cref="InvalidOperationException"/>
      /// <exception cref="PlatformNotSupportedException">The operating system is older than Windows Vista.</exception>
      /// <exception cref="PlatformNotSupportedException">The operating system is older than Windows Vista.</exception>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="filePaths">1つ以上のファイルパスのリスト。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SuppressMessage("Microsoft.Usage", "CA1806:DoNotIgnoreMethodResults", MessageId = "Alphaleonis.Win32.Filesystem.NativeMethods.RmEndSession(System.UInt32)")]
      internal static Collection<Process> GetProcessForFileLockCore(KernelTransaction transaction, Collection<string> filePaths, PathFormat pathFormat)
      {
         if (!NativeMethods.IsAtLeastWindowsVista)
         {
            throw new PlatformNotSupportedException(new Win32Exception((int) Win32Errors.ERROR_OLD_WIN_VERSION).Message);
         }

         if (null == filePaths)
         {
            throw new ArgumentNullException("filePaths");
         }

         if (filePaths.Count == 0)
         {
            throw new ArgumentOutOfRangeException("filePaths", "No paths specified.");
         }


         var isLongPath = pathFormat == PathFormat.LongFullPath;
         var allPaths = isLongPath ? new Collection<string>(filePaths) : new Collection<string>();

         if (!isLongPath)
         {
            foreach (var path in filePaths)
               allPaths.Add(Path.GetExtendedLengthPathCore(transaction, path, pathFormat, GetFullPathOptions.RemoveTrailingDirectorySeparator | GetFullPathOptions.FullCheck));
         }




         var success = NativeMethods.RmStartSession(out var sessionHandle, 0, Guid.NewGuid().ToString()) == Win32Errors.ERROR_SUCCESS;

         var lastError = Marshal.GetLastWin32Error();
         if (!success)
         {
            NativeError.ThrowException(lastError);
         }




         var processes = new Collection<Process>();

         try
         {
            // すべての実行中プロセスのスナップショットカウント。
            var processesFound = (uint) Process.GetProcesses().Length;
            uint lpdwRebootReasons = 0;


            success = NativeMethods.RmRegisterResources(sessionHandle, (uint) allPaths.Count, allPaths.ToArray(), 0, null, 0, null) == Win32Errors.ERROR_SUCCESS;

            lastError = Marshal.GetLastWin32Error();
            if (!success)
            {
               NativeError.ThrowException(lastError);
            }


            GetList:

            var processInfo = new NativeMethods.RM_PROCESS_INFO[processesFound];
            var processesTotal = processesFound;


            lastError = NativeMethods.RmGetList(sessionHandle, out processesFound, ref processesTotal, processInfo, ref lpdwRebootReasons);


            // 既に実行中プロセスの総数があるため、これは不要です。
            if (lastError == Win32Errors.ERROR_MORE_DATA)
            {
               goto GetList;
            }


            if (lastError != Win32Errors.ERROR_SUCCESS)
            {
               NativeError.ThrowException(lastError);
            }


            for (var i = 0; i < processesTotal; i++)
            {
               try
               {
                  processes.Add(Process.GetProcessById(processInfo[i].Process.dwProcessId));
               }

               // MSDN: processIdパラメータで指定されたプロセスは実行されていません。識別子が期限切れの可能性があります。
               catch (ArgumentException) {}
            }
         }
         finally
         {
            NativeMethods.RmEndSession(sessionHandle);
         }


         return processes.Count == 0 ? null : processes;
      }
   }
}
