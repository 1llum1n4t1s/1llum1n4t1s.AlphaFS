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
   public sealed partial class DirectoryInfo
   {
      /// <summary>非トランザクション/トランザクションファイルまたはディレクトリとその子をコピー/移動して新しい場所に配置します。
      /// <see cref="CopyOptions"/> または <see cref="MoveOptions"/> を指定でき、コールバック関数を通じてアプリケーションに進行状況を通知できます。
      /// </summary>
      /// <returns>コピーまたは移動操作の詳細を含む <see cref="CopyMoveResult"/> クラス。</returns>
      /// <remarks>
      ///   <para>非常に大きなファイル転送には、<see cref="CopyOptions.NoBuffering"/> オプションが推奨されます。</para>
      ///   <para><paramref name="moveOptions"/> に <see cref="MoveOptions.ReplaceExisting"/> が含まれていない限り、Move メソッドを使用して既存のファイルを上書きすることはできません。</para>
      ///   <para>同じ名前のファイルをそのディレクトリに移動してファイルを置き換えようとすると、IOException が発生することに注意してください。</para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="destinationPath">コピー先のディレクトリパス。</param>
      /// <param name="preserveDates">元のタイムスタンプを保持する場合は <c>true</c>、それ以外の場合は <c>false</c>。</param>
      /// <param name="copyOptions">ファイルのコピー方法を指定する <see cref="CopyOptions"/>。このパラメーターは <c>null</c> にできます。</param>
      /// <param name="moveOptions">ファイルの移動方法を指定する <see cref="MoveOptions"/>。このパラメーターは <c>null</c> にできます。</param>
      /// <param name="filters">処理で使用するカスタムフィルターの仕様。</param>
      /// <param name="progressHandler">ファイルの別の部分がコピーされるたびに呼び出されるコールバック関数。このパラメーターは <c>null</c> にできます。</param>
      /// <param name="userProgressData">コールバック関数に渡される引数。このパラメーターは <c>null</c> にできます。</param>
      /// <param name="longFullPath">取得された長い完全パスを返します。</param>
      /// <param name="pathFormat">パスパラメーターの形式を示します。</param>
      [SecurityCritical]
      private CopyMoveResult CopyToMoveToCore(string destinationPath, bool preserveDates, CopyOptions? copyOptions, MoveOptions? moveOptions, DirectoryEnumerationFilters filters, CopyMoveProgressRoutine progressHandler, object userProgressData, out string longFullPath, PathFormat pathFormat)
      {
         longFullPath = Path.GetExtendedLengthPathCore(Transaction, destinationPath, pathFormat, GetFullPathOptions.TrimEnd | GetFullPathOptions.RemoveTrailingDirectorySeparator | GetFullPathOptions.FullCheck);

         return Directory.CopyMoveCore(new CopyMoveArguments
         {
            Transaction = Transaction,
            SourcePathLp = LongFullName,
            SourcePath = LongFullName,
            DestinationPathLp = longFullPath,
            DestinationPath = longFullPath,
            CopyOptions = preserveDates ? copyOptions | CopyOptions.CopyTimestamp : copyOptions & ~CopyOptions.CopyTimestamp,
            DirectoryEnumerationFilters = filters,
            MoveOptions = moveOptions,
            ProgressHandler = progressHandler,
            UserProgressData = userProgressData,
            PathFormat = PathFormat.LongFullPath
         });
      }
   }
}
