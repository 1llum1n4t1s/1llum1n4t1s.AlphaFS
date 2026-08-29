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

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class Directory
   {
      private static void CopyFolderTimestamps(CopyMoveArguments cma)
      {
         // TODO 2018-01-09: ローカル + UNCパスの組み合わせでまだ100%ではない。
         // ソースフォルダを走査し、フォルダのみを処理する。

         foreach (var fseiSource in EnumerateFileSystemEntryInfosCore<FileSystemEntryInfo>(true, cma.Transaction, cma.SourcePathLp, Path.WildcardStarMatchAll, null,
                     DirectoryEnumerationOptions.Recursive | DirectoryEnumerationOptions.SkipReparsePoints, cma.DirectoryEnumerationFilters, PathFormat.LongFullPath))
         {
            var destinationPath = cma.DestinationPathLp + fseiSource.LongFullPath.Substring(cma.SourcePathLp.Length);

            File.CopyTimestampsCore(cma.Transaction, true, fseiSource.LongFullPath, destinationPath, false, PathFormat.LongFullPath);
         }


         // ルートディレクトリ（指定されたパス）を処理する。

         File.CopyTimestampsCore(cma.Transaction, true, cma.SourcePathLp, cma.DestinationPathLp, false, PathFormat.LongFullPath);


         // TODO: コンピュータで有効にすると、FindFirstFileが最終アクセス日時を変更してしまう。
      }
   }
}
