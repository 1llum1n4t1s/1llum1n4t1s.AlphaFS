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
   partial class FileInfo
   {
      /// <summary>既存のファイルを新しいファイルにコピー/移動します。既存のファイルの上書きを許可します。</summary>
      /// <returns>コピーまたは移動操作のステータスを含む <see cref="CopyMoveResult"/> クラス。</returns>
      /// <remarks>
      ///   <para>非常に大きなファイル転送には、<see cref="CopyOptions.NoBuffering"/> オプションが推奨されます。</para>
      ///   <para>可能な限り、このメソッドでは短いファイル名（<c>XXXXXX~1.XXX</c> など）の使用を避けてください。</para>
      ///   <para>2 つのファイルに同等の短いファイル名がある場合、このメソッドが失敗して例外がスローされるか、望ましくない動作が発生する可能性があります。</para>
      /// </remarks>
      /// <param name="destinationPath"><para>コピー先ディレクトリの完全パス文字列。</para></param>
      /// <param name="copyOptions"><para>このパラメーターは <c>null</c> にできます。ファイルのコピー方法を指定するには <see cref="CopyOptions"/> を使用します。</para></param>
      /// <param name="moveOptions"><para>このパラメーターは <c>null</c> にできます。ファイルの移動方法を指定するには <see cref="MoveOptions"/> を使用します。</para></param>
      /// <param name="preserveDates">元のタイムスタンプを保持する場合は <c>true</c>、それ以外の場合は <c>false</c>。</param>
      /// <param name="progressHandler"><para>このパラメーターは <c>null</c> にできます。ファイルの別の部分がコピーされるたびに呼び出されるコールバック関数。</para></param>
      /// <param name="userProgressData"><para>このパラメーターは <c>null</c> にできます。コールバック関数に渡される引数。</para></param>
      /// <param name="longFullPath">[out] 取得された長い完全パスを返します。</param>
      /// <param name="pathFormat">パスパラメーターの形式を示します。</param>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      [SecurityCritical]
      private CopyMoveResult CopyToMoveToCore(string destinationPath, CopyOptions? copyOptions, MoveOptions? moveOptions, bool preserveDates, CopyMoveProgressRoutine progressHandler, object userProgressData, out string longFullPath, PathFormat pathFormat)
      {
         longFullPath = Path.GetExtendedLengthPathCore(Transaction, destinationPath, pathFormat, GetFullPathOptions.TrimEnd | GetFullPathOptions.RemoveTrailingDirectorySeparator | GetFullPathOptions.FullCheck);

         return File.CopyMoveCore(false, new CopyMoveArguments
         {
            Transaction = Transaction,
            CopyOptions = preserveDates ? copyOptions | CopyOptions.CopyTimestamp : copyOptions & ~CopyOptions.CopyTimestamp,
            MoveOptions = moveOptions,
            ProgressHandler = progressHandler,
            UserProgressData = userProgressData,
            PathFormat = PathFormat.LongFullPath

         }, false, false, LongFullName, longFullPath, null);
      }


      /// <summary>現在の <see cref="FileInfo"/> インスタンスを新しいコピー先パスで更新します。</summary>
      private void UpdateDestinationPath(string destinationPath, string destinationPathLp)
      {
         _name = Path.GetFileName(destinationPathLp, true);

         UpdateSourcePath(destinationPath, destinationPathLp);
      }
   }
}
