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
using System.Runtime.InteropServices;
using System.Security;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class File
   {
      // Symbolic Link Effects on File Systems Functions: https://msdn.microsoft.com/en-us/library/windows/desktop/aa365682(v=vs.85).aspx


      // MSDN: lpProgressRoutineがユーザーによる操作のキャンセルによりPROGRESS_CANCELを返した場合、
      // CopyFileExはゼロを返し、GetLastErrorはERROR_REQUEST_ABORTEDを返します。
      // この場合、部分的にコピーされたコピー先ファイルは削除されます。
      //
      // lpProgressRoutineがユーザーによる操作の停止によりPROGRESS_STOPを返した場合、
      // CopyFileExはゼロを返し、GetLastErrorはERROR_REQUEST_ABORTEDを返します。
      // この場合、部分的にコピーされたコピー先ファイルはそのまま残されます。


      // 注: 両方のパスが同じボリュームを参照していても、パスの1つがUNCパスの場合、MoveFileXxxは失敗します。
      // 例: src = C:\TempSrc, dst = \\localhost\C$\TempDst

      // MoveFileXxxはレジストリにアクセスできない場合に失敗します。この関数は再起動時に名前変更されるファイルの場所を次のレジストリ値に格納します:
      //
      //    HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\PendingFileRenameOperations
      //
      // このレジストリ値はREG_MULTI_SZ型です。各名前変更操作は、削除かどうかに応じて次のNULL終端文字列の1つを格納します:
      //
      //    szDstFile\0\0              : 再起動時にszDstFileが削除されることを示します。
      //    szSrcFile\0szDstFile\0     : 再起動時にszSrcFileがszDstFileに名前変更されることを示します。


      [SecurityCritical]
      private static bool CopyMoveNative(CopyMoveArguments cma, bool isMove, string sourcePathLp, string destinationPathLp, out bool cancel, out int lastError)
      {
         cancel = false;

         var success = null == cma.Transaction || !NativeMethods.IsAtLeastWindowsVista

            // CopyFileEx() / CopyFileTransacted() / MoveFileWithProgress() / MoveFileTransacted()
            // 2013-04-15: MSDNはLongPathの使用を確認しています。


            ? isMove
               ? NativeMethods.MoveFileWithProgress(sourcePathLp, destinationPathLp, cma.Routine, IntPtr.Zero, (MoveOptions) cma.MoveOptions)

               : NativeMethods.CopyFileEx(sourcePathLp, destinationPathLp, cma.Routine, IntPtr.Zero, out cancel, (CopyOptions) cma.CopyOptions)

            : isMove
               ? NativeMethods.MoveFileTransacted(sourcePathLp, destinationPathLp, cma.Routine, IntPtr.Zero, (MoveOptions) cma.MoveOptions, cma.Transaction.SafeHandle)

               : NativeMethods.CopyFileTransacted(sourcePathLp, destinationPathLp, cma.Routine, IntPtr.Zero, out cancel, (CopyOptions) cma.CopyOptions, cma.Transaction.SafeHandle);


         lastError = Marshal.GetLastWin32Error();


         return success;
      }
   }
}
