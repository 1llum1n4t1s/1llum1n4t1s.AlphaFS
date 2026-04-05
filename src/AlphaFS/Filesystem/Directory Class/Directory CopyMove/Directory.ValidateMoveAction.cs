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
using System.Security;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class Directory
   {
      [SecurityCritical]
      internal static CopyMoveArguments ValidateMoveAction(CopyMoveArguments cma)
      {
         // 移動操作またはコピー操作フォールバックが可能かどうかを判定する。

         cma.IsCopy = false;
         cma.EmulateMove = false;


         // 両方のパスのルート部分を比較する。

         var equalRootPaths = Path.GetPathRoot(cma.SourcePathLp, false).Equals(Path.GetPathRoot(cma.DestinationPathLp, false), StringComparison.OrdinalIgnoreCase);


         // Volume.IsSameVolume() メソッドは、一方のパスがUNCパスであっても、両方のパスが同じボリュームを参照する場合にtrueを返す。
         // 例: src = C:\TempSrc、dst = \\localhost\C$\TempDst

         var isSameVolume = equalRootPaths || Volume.IsSameVolume(cma.SourcePathLp, cma.DestinationPathLp);

         var isMove = isSameVolume && equalRootPaths;

         if (!isMove)
         {
            // Move() は Copy() と Delete() を使ってエミュレートできるが、MoveOptions.CopyAllowed フラグが設定されている場合のみ。

            isMove = File.HasCopyAllowed(cma.MoveOptions);


            // MSDN: .NET3.5+: IOException: 異なるボリュームにディレクトリを移動しようとした。

            if (!isMove)
            {
               NativeError.ThrowException(Win32Errors.ERROR_NOT_SAME_DEVICE, cma.SourcePathLp, cma.DestinationPathLp);
            }
         }


         // MoveFileXxx メソッドは以下の場合に失敗する:
         // - ディレクトリが移動される場合
         // - 両方のパスが同じボリュームを参照していても、一方のパスがUNCパスの場合。
         //   例: src = C:\TempSrc、dst = \\localhost\C$\TempDst

         if (isMove)
         {
            var srcIsUncPath = Path.IsUncPathCore(cma.SourcePathLp, false, false);
            var dstIsUncPath = Path.IsUncPathCore(cma.DestinationPathLp, false, false);

            isMove = srcIsUncPath == dstIsUncPath;
         }


         isMove = isMove && isSameVolume && equalRootPaths;


         // Move() をエミュレートする。
         if (!isMove)
         {
            cma.MoveOptions = null;

            cma.IsCopy = true;
            cma.EmulateMove = true;
            cma.CopyOptions = CopyOptions.None;
         }


         return cma;
      }
   }
}
