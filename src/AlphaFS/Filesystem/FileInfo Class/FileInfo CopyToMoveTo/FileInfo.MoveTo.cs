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
      #region .NET

      /// <summary>指定されたファイルを新しい場所に移動します。新しいファイル名を指定するオプションを提供します。</summary>
      /// <returns>移動操作の詳細を含む <see cref="CopyMoveResult"/> クラス。</returns>
      /// <remarks>
      ///   <para>デフォルトで既存のファイルの上書きを防止するには、このメソッドを使用します。</para>
      ///   <para>このメソッドはディスクボリュームをまたいで動作します。</para>
      ///   <para>たとえば、ファイル c:\MyFile.txt を d:\public に移動して NewFile.txt に名前を変更できます。</para>
      ///   <para>可能な限り、このメソッドでは短いファイル名（<c>XXXXXX~1.XXX</c> など）の使用を避けてください。</para>
      ///   <para>2 つのファイルに同等の短いファイル名がある場合、このメソッドが失敗して例外がスローされるか、望ましくない動作が発生する可能性があります。</para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="FileNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="destinationPath">ファイルの移動先のパス。別のファイル名を指定できます。</param>
      [SecurityCritical]
      public void MoveTo(string destinationPath)
      {

         CopyToMoveToCore(destinationPath, null, MoveOptions.CopyAllowed, false, null, null, out var destinationPathLp, PathFormat.RelativePath);

         UpdateDestinationPath(destinationPath, destinationPathLp);
      }

      #endregion // .NET


      /// <summary>[AlphaFS] 指定されたファイルを新しい場所に移動します。新しいファイル名を指定するオプションを提供します。</summary>
      /// <returns>正常に移動された場合、完全修飾パスを持つ新しい <see cref="FileInfo"/> インスタンスを返します。</returns>
      /// <remarks>
      ///   <para>Use this method to prevent overwriting of an existing file by default.</para>
      ///   <para>This method works across disk volumes.</para>
      ///   <para>For example, the file c:\MyFile.txt can be moved to d:\public and renamed NewFile.txt.</para>
      ///   <para>Whenever possible, avoid using short file names (such as <c>XXXXXX~1.XXX</c>) with this method.</para>
      ///   <para>If two files have equivalent short file names then this method may fail and raise an exception and/or result in undesirable behavior.</para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="FileNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="destinationPath">ファイルの移動先のパス。別のファイル名を指定できます。</param>
      /// <param name="pathFormat">パスパラメーターの形式を示します。</param>
      [SecurityCritical]
      public FileInfo MoveTo(string destinationPath, PathFormat pathFormat)
      {

         CopyToMoveToCore(destinationPath, null, MoveOptions.CopyAllowed, false, null, null, out var destinationPathLp, pathFormat);

         UpdateDestinationPath(destinationPath, destinationPathLp);

         return new FileInfo(Transaction, destinationPathLp, PathFormat.LongFullPath);
      }
      

      /// <summary>[AlphaFS] 指定されたファイルを新しい場所に移動します。新しいファイル名を指定するオプションを提供し、<see cref="MoveOptions"/> を指定できます。</summary>
      /// <returns>Returns a new <see cref="FileInfo"/> instance with a fully qualified path when successfully moved.</returns>
      /// <remarks>
      ///   <para>Use this method to allow or prevent overwriting of an existing file.</para>
      ///   <para>This method works across disk volumes.</para>
      ///   <para>For example, the file c:\MyFile.txt can be moved to d:\public and renamed NewFile.txt.</para>
      ///   <para>Whenever possible, avoid using short file names (such as <c>XXXXXX~1.XXX</c>) with this method.</para>
      ///   <para>If two files have equivalent short file names then this method may fail and raise an exception and/or result in undesirable behavior.</para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="FileNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="destinationPath">ファイルの移動先のパス。別のファイル名を指定できます。</param>
      /// <param name="moveOptions">ディレクトリの移動方法を指定する <see cref="MoveOptions"/>。このパラメーターは <c>null</c> にできます。</param>
      [SecurityCritical]
      public FileInfo MoveTo(string destinationPath, MoveOptions moveOptions)
      {

         CopyToMoveToCore(destinationPath, null, moveOptions, false, null, null, out var destinationPathLp, PathFormat.RelativePath);

         UpdateDestinationPath(destinationPath, destinationPathLp);

         return null != destinationPathLp ? new FileInfo(Transaction, destinationPathLp, PathFormat.LongFullPath) : null;
      }


      /// <summary>[AlphaFS] 指定されたファイルを新しい場所に移動します。新しいファイル名を指定するオプションを提供し、<see cref="MoveOptions"/> を指定できます。</summary>
      /// <returns>Returns a new <see cref="FileInfo"/> instance with a fully qualified path when successfully moved.</returns>
      /// <remarks>
      ///   <para>Use this method to allow or prevent overwriting of an existing file.</para>
      ///   <para>This method works across disk volumes.</para>
      ///   <para>For example, the file c:\MyFile.txt can be moved to d:\public and renamed NewFile.txt.</para>
      ///   <para>Whenever possible, avoid using short file names (such as <c>XXXXXX~1.XXX</c>) with this method.</para>
      ///   <para>If two files have equivalent short file names then this method may fail and raise an exception and/or result in undesirable behavior.</para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="FileNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="destinationPath">ファイルの移動先のパス。別のファイル名を指定できます。</param>
      /// <param name="moveOptions">ディレクトリの移動方法を指定する <see cref="MoveOptions"/>。このパラメーターは <c>null</c> にできます。</param>
      /// <param name="pathFormat">パスパラメーターの形式を示します。</param>
      [SecurityCritical]
      public FileInfo MoveTo(string destinationPath, MoveOptions moveOptions, PathFormat pathFormat)
      {

         CopyToMoveToCore(destinationPath, null, moveOptions, false, null, null, out var destinationPathLp, pathFormat);

         UpdateDestinationPath(destinationPath, destinationPathLp);

         return null != destinationPathLp ? new FileInfo(Transaction, destinationPathLp, PathFormat.LongFullPath) : null;
      }
      

      /// <summary>[AlphaFS] 指定されたファイルを新しい場所に移動します。新しいファイル名を指定するオプションを提供し、<see cref="MoveOptions"/> を指定でき、
      /// コールバック関数を通じてアプリケーションに進行状況を通知できます。
      /// </summary>
      /// <returns>移動操作のステータスを含む <see cref="CopyMoveResult"/> クラス。</returns>
      /// <remarks>
      ///   <para>Use this method to allow or prevent overwriting of an existing file.</para>
      ///   <para>This method works across disk volumes.</para>
      ///   <para>For example, the file c:\MyFile.txt can be moved to d:\public and renamed NewFile.txt.</para>
      ///   <para>Whenever possible, avoid using short file names (such as <c>XXXXXX~1.XXX</c>) with this method.</para>
      ///   <para>If two files have equivalent short file names then this method may fail and raise an exception and/or result in undesirable behavior.</para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="FileNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="destinationPath">ファイルの移動先のパス。別のファイル名を指定できます。</param>
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


         var cmr = CopyToMoveToCore(destinationPath, null, moveOptions, false, progressHandler, userProgressData, out var destinationPathLp, PathFormat.RelativePath);

         UpdateDestinationPath(destinationPath, destinationPathLp);

         return cmr;
      }


      /// <summary>[AlphaFS] 指定されたファイルを新しい場所に移動します。新しいファイル名を指定するオプションを提供し、<see cref="MoveOptions"/> を指定できます。</summary>
      /// <returns>移動操作のステータスを含む <see cref="CopyMoveResult"/> クラス。</returns>
      /// <remarks>
      ///   <para>Use this method to allow or prevent overwriting of an existing file.</para>
      ///   <para>This method works across disk volumes.</para>
      ///   <para>For example, the file c:\MyFile.txt can be moved to d:\public and renamed NewFile.txt.</para>
      ///   <para>Whenever possible, avoid using short file names (such as <c>XXXXXX~1.XXX</c>) with this method.</para>
      ///   <para>If two files have equivalent short file names then this method may fail and raise an exception and/or result in undesirable behavior.</para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="FileNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="destinationPath">ファイルの移動先のパス。別のファイル名を指定できます。</param>
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


         var cmr = CopyToMoveToCore(destinationPath, null, moveOptions, false, progressHandler, userProgressData, out var destinationPathLp, pathFormat);

         UpdateDestinationPath(destinationPath, destinationPathLp);

         return cmr;
      }
   }
}
