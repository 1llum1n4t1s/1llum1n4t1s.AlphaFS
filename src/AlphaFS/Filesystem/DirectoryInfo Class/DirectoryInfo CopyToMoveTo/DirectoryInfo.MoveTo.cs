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
      #region .NET

      /// <summary><see cref="DirectoryInfo"/> インスタンスとその内容を新しいパスに移動します。</summary>
      /// <remarks>
      ///   <para>デフォルトで既存のディレクトリの上書きを防止するには、このメソッドを使用します。</para>
      ///   <para>このメソッドはディスクボリュームをまたいで動作しません。</para>
      ///   <para>可能な限り、このメソッドでは短いファイル名（<c>XXXXXX~1.XXX</c> など）の使用を避けてください。</para>
      ///   <para>2 つのディレクトリに同等の短いファイル名がある場合、このメソッドが失敗して例外がスローされるか、望ましくない動作が発生する可能性があります。</para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="destinationPath">
      ///   <para>このディレクトリの移動先の名前とパス。</para>
      ///   <para>移動先は別のディスクボリュームまたは同一の名前のディレクトリにすることはできません。</para>
      ///   <para>このディレクトリをサブディレクトリとして追加する既存のディレクトリにすることができます。</para>
      /// </param>
      [SecurityCritical]
      public void MoveTo(string destinationPath)
      {

         CopyToMoveToCore(destinationPath, false, null, MoveOptions.None, null, null, null, out var destinationPathLp, PathFormat.RelativePath);

         UpdateSourcePath(destinationPath, destinationPathLp);
      }

      #endregion // .NET


      /// <summary>[AlphaFS] <see cref="DirectoryInfo"/> インスタンスとその内容を新しいパスに移動します。</summary>
      /// <remarks>
      ///   <para>デフォルトで既存のディレクトリの上書きを防止するには、このメソッドを使用します。</para>
      ///   <para>このメソッドはディスクボリュームをまたいで動作しません。</para>
      ///   <para>可能な限り、このメソッドでは短いファイル名（<c>XXXXXX~1.XXX</c> など）の使用を避けてください。</para>
      ///   <para>2 つのディレクトリに同等の短いファイル名がある場合、このメソッドが失敗して例外がスローされるか、望ましくない動作が発生する可能性があります。</para>
      /// </remarks>
      /// <returns>ディレクトリが完全に移動された場合の新しい <see cref="DirectoryInfo"/> インスタンス。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="destinationPath">
      ///   <para>このディレクトリの移動先の名前とパス。</para>
      ///   <para>移動先は別のディスクボリュームまたは同一の名前のディレクトリにすることはできません。</para>
      ///   <para>このディレクトリをサブディレクトリとして追加する既存のディレクトリにすることができます。</para>
      /// </param>
      /// <param name="pathFormat">パスパラメーターの形式を示します。</param>
      [SecurityCritical]
      public DirectoryInfo MoveTo(string destinationPath, PathFormat pathFormat)
      {

         CopyToMoveToCore(destinationPath, false, null, MoveOptions.None, null, null, null, out var destinationPathLp, pathFormat);

         UpdateSourcePath(destinationPath, destinationPathLp);

         return new DirectoryInfo(Transaction, destinationPathLp, PathFormat.LongFullPath);
      }


      /// <summary>[AlphaFS] <see cref="DirectoryInfo"/> インスタンスとその内容を新しいパスに移動します。<see cref="MoveOptions"/> を指定できます。</summary>
      /// <remarks>
      ///   <para>Use this method to allow or prevent overwriting of an existing directory.</para>
      ///   <para>This method does not work across disk volumes unless <paramref name="moveOptions"/> contains <see cref="MoveOptions.CopyAllowed"/>.</para>
      ///   <para>Whenever possible, avoid using short file names (such as <c>XXXXXX~1.XXX</c>) with this method.</para>
      ///   <para>If two directories have equivalent short file names then this method may fail and raise an exception and/or result in undesirable behavior.</para>
      /// </remarks>
      /// <returns>A new <see cref="DirectoryInfo"/> instance if the directory was completely moved.</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="destinationPath">
      ///   <para>The name and path to which to move this directory.</para>
      ///   <para>The destination cannot be another disk volume unless <paramref name="moveOptions"/> contains <see cref="MoveOptions.CopyAllowed"/>, or a directory with the identical name.</para>
      ///   <para>It can be an existing directory to which you want to add this directory as a subdirectory.</para>
      /// </param>
      /// <param name="moveOptions">ディレクトリの移動方法を指定する <see cref="MoveOptions"/>。このパラメーターは <c>null</c> にできます。</param>
      [SecurityCritical]
      public DirectoryInfo MoveTo(string destinationPath, MoveOptions moveOptions)
      {

         CopyToMoveToCore(destinationPath, false, null, moveOptions, null, null, null, out var destinationPathLp, PathFormat.RelativePath);

         UpdateSourcePath(destinationPath, destinationPathLp);

         return null != destinationPathLp ? new DirectoryInfo(Transaction, destinationPathLp, PathFormat.LongFullPath) : null;
      }


      /// <summary>[AlphaFS] <see cref="DirectoryInfo"/> インスタンスとその内容を新しいパスに移動します。<see cref="MoveOptions"/> を指定できます。</summary>
      /// <remarks>
      ///   <para>Use this method to allow or prevent overwriting of an existing directory.</para>
      ///   <para>This method does not work across disk volumes unless <paramref name="moveOptions"/> contains <see cref="MoveOptions.CopyAllowed"/>.</para>
      ///   <para>Whenever possible, avoid using short file names (such as <c>XXXXXX~1.XXX</c>) with this method.</para>
      ///   <para>If two directories have equivalent short file names then this method may fail and raise an exception and/or result in undesirable behavior.</para>
      /// </remarks>
      /// <returns>A new <see cref="DirectoryInfo"/> instance if the directory was completely moved.</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="destinationPath">
      ///   <para>The name and path to which to move this directory.</para>
      ///   <para>The destination cannot be another disk volume unless <paramref name="moveOptions"/> contains <see cref="MoveOptions.CopyAllowed"/>, or a directory with the identical name.</para>
      ///   <para>It can be an existing directory to which you want to add this directory as a subdirectory.</para>
      /// </param>
      /// <param name="moveOptions">ディレクトリの移動方法を指定する <see cref="MoveOptions"/>。このパラメーターは <c>null</c> にできます。</param>
      /// <param name="pathFormat">パスパラメーターの形式を示します。</param>
      [SecurityCritical]
      public DirectoryInfo MoveTo(string destinationPath, MoveOptions moveOptions, PathFormat pathFormat)
      {

         CopyToMoveToCore(destinationPath, false, null, moveOptions, null, null, null, out var destinationPathLp, pathFormat);

         UpdateSourcePath(destinationPath, destinationPathLp);

         return null != destinationPathLp ? new DirectoryInfo(Transaction, destinationPathLp, PathFormat.LongFullPath) : null;
      }


      /// <summary>[AlphaFS] <see cref="DirectoryInfo"/> インスタンスとその内容を新しいパスに移動します。<see cref="MoveOptions"/> を指定でき、
      /// コールバック関数を通じてアプリケーションに進行状況を通知できます。
      /// </summary>
      /// <remarks>
      ///   <para>Use this method to allow or prevent overwriting of an existing directory.</para>
      ///   <para>This method does not work across disk volumes unless <paramref name="moveOptions"/> contains <see cref="MoveOptions.CopyAllowed"/>.</para>
      ///   <para>Whenever possible, avoid using short file names (such as <c>XXXXXX~1.XXX</c>) with this method.</para>
      ///   <para>If two directories have equivalent short file names then this method may fail and raise an exception and/or result in undesirable behavior.</para>
      /// </remarks>
      /// <returns>移動操作の詳細を含む <see cref="CopyMoveResult"/> クラス。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="destinationPath">
      ///   <para>The name and path to which to move this directory.</para>
      ///   <para>The destination cannot be another disk volume unless <paramref name="moveOptions"/> contains <see cref="MoveOptions.CopyAllowed"/>, or a directory with the identical name.</para>
      ///   <para>It can be an existing directory to which you want to add this directory as a subdirectory.</para>
      /// </param>
      /// <param name="moveOptions">ディレクトリの移動方法を指定する <see cref="MoveOptions"/>。このパラメーターは <c>null</c> にできます。</param>
      /// <param name="progressHandler">ディレクトリの別の部分が移動されるたびに呼び出されるコールバック関数。このパラメーターは <c>null</c> にできます。</param>
      /// <param name="userProgressData">コールバック関数に渡される引数。このパラメーターは <c>null</c> にできます。</param>
      [SecurityCritical]
      public CopyMoveResult MoveTo(string destinationPath, MoveOptions moveOptions, CopyMoveProgressRoutine progressHandler, object userProgressData)
      {
         // DelayUntilReboot を拒否する。
         if ((moveOptions & MoveOptions.DelayUntilReboot) != 0)
         {
            throw new ArgumentException("The DelayUntilReboot flag is invalid for this method.", "moveOptions");
         }

         var cmr = CopyToMoveToCore(destinationPath, false, null, moveOptions, null, progressHandler, userProgressData, out var destinationPathLp, PathFormat.RelativePath);

         UpdateSourcePath(destinationPath, destinationPathLp);

         return cmr;
      }


      /// <summary>[AlphaFS] <see cref="DirectoryInfo"/> インスタンスとその内容を新しいパスに移動します。<see cref="MoveOptions"/> を指定でき、
      ///   <para>コールバック関数を通じてアプリケーションに進行状況を通知できます。</para>
      /// </summary>
      /// <remarks>
      ///   <para>Use this method to allow or prevent overwriting of an existing directory.</para>
      ///   <para>This method does not work across disk volumes unless <paramref name="moveOptions"/> contains <see cref="MoveOptions.CopyAllowed"/>.</para>
      ///   <para>Whenever possible, avoid using short file names (such as <c>XXXXXX~1.XXX</c>) with this method.</para>
      ///   <para>If two directories have equivalent short file names then this method may fail and raise an exception and/or result in undesirable behavior.</para>
      /// </remarks>
      /// <returns>移動操作の詳細を含む <see cref="CopyMoveResult"/> クラス。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="destinationPath">
      ///   <para>The name and path to which to move this directory.</para>
      ///   <para>The destination cannot be another disk volume unless <paramref name="moveOptions"/> contains <see cref="MoveOptions.CopyAllowed"/>, or a directory with the identical name.</para>
      ///   <para>It can be an existing directory to which you want to add this directory as a subdirectory.</para>
      /// </param>
      /// <param name="moveOptions">ディレクトリの移動方法を指定する <see cref="MoveOptions"/>。このパラメーターは <c>null</c> にできます。</param>
      /// <param name="progressHandler">ディレクトリの別の部分が移動されるたびに呼び出されるコールバック関数。このパラメーターは <c>null</c> にできます。</param>
      /// <param name="userProgressData">コールバック関数に渡される引数。このパラメーターは <c>null</c> にできます。</param>
      /// <param name="pathFormat">パスパラメーターの形式を示します。</param>
      [SecurityCritical]
      public CopyMoveResult MoveTo(string destinationPath, MoveOptions moveOptions, CopyMoveProgressRoutine progressHandler, object userProgressData, PathFormat pathFormat)
      {
         // DelayUntilReboot を拒否する。
         if ((moveOptions & MoveOptions.DelayUntilReboot) != 0)
         {
            throw new ArgumentException("The DelayUntilReboot flag is invalid for this method.", "moveOptions");
         }

         var cmr = CopyToMoveToCore(destinationPath, false, null, moveOptions, null, progressHandler, userProgressData, out var destinationPathLp, pathFormat);

         UpdateSourcePath(destinationPath, destinationPathLp);

         return cmr;
      }
   }
}
