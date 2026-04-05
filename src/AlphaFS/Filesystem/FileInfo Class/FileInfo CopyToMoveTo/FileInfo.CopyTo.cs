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

      /// <summary>既存のファイルを新しいファイルにコピーします。既存のファイルの上書きは許可しません。</summary>
      /// <returns>完全修飾パスを持つ新しい <see cref="FileInfo"/> インスタンス。</returns>
      /// <remarks>
      ///   <para>デフォルトで既存のファイルの上書きを防止するには、このメソッドを使用します。</para>
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
      /// <param name="destinationPath">コピー先の新しいファイルの名前。</param>
      [SecurityCritical]
      public FileInfo CopyTo(string destinationPath)
      {

         CopyToMoveToCore(destinationPath, CopyOptions.FailIfExists, null, false, null, null, out var destinationPathLp, PathFormat.RelativePath);

         UpdateDestinationPath(destinationPath, destinationPathLp);

         return new FileInfo(Transaction, destinationPathLp, PathFormat.LongFullPath);
      }


      /// <summary>既存のファイルを新しいファイルにコピーします。既存のファイルの上書きを許可します。</summary>
      /// <returns>完全修飾パスを持つ新しい <see cref="FileInfo"/> インスタンス。</returns>
      /// <remarks>
      ///   <para>このメソッドを使用して、既存のファイルの上書きを許可または防止します。</para>
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
      /// <param name="destinationPath">コピー先の新しいファイルの名前。</param>
      /// <param name="overwrite">既存のファイルの上書きを許可する場合は <c>true</c>、それ以外の場合は <c>false</c>。</param>
      [SecurityCritical]
      public FileInfo CopyTo(string destinationPath, bool overwrite)
      {

         CopyToMoveToCore(destinationPath, overwrite ? CopyOptions.None : CopyOptions.FailIfExists, null, false, null, null, out var destinationPathLp, PathFormat.RelativePath);

         UpdateDestinationPath(destinationPath, destinationPathLp);

         return new FileInfo(Transaction, destinationPathLp, PathFormat.LongFullPath);
      }

      #endregion // .NET


      /// <summary>[AlphaFS] Copies an existing file to a new file, disallowing the overwriting of an existing file.</summary>
      /// <returns>A new <see cref="FileInfo"/> instance with a fully qualified path.</returns>
      /// <remarks>
      ///   <para>Use this method to prevent overwriting of an existing file by default.</para>
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
      /// <param name="destinationPath">コピー先の新しいファイルの名前。</param>
      /// <param name="pathFormat">パスパラメーターの形式を示します。</param>
      [SecurityCritical]
      public FileInfo CopyTo(string destinationPath, PathFormat pathFormat)
      {

         CopyToMoveToCore(destinationPath, CopyOptions.FailIfExists, null, false, null, null, out var destinationPathLp, pathFormat);

         UpdateDestinationPath(destinationPath, destinationPathLp);

         return new FileInfo(Transaction, destinationPathLp, PathFormat.LongFullPath);
      }


      /// <summary>[AlphaFS] Copies an existing file to a new file, allowing the overwriting of an existing file.</summary>
      /// <returns>A new <see cref="FileInfo"/> instance with a fully qualified path.</returns>
      /// <remarks>
      ///   <para>Use this method to allow or prevent overwriting of an existing file.</para>
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
      /// <param name="destinationPath">コピー先の新しいファイルの名前。</param>
      /// <param name="overwrite">既存のファイルの上書きを許可する場合は <c>true</c>、それ以外の場合は <c>false</c>。</param>
      /// <param name="pathFormat">パスパラメーターの形式を示します。</param>
      [SecurityCritical]
      public FileInfo CopyTo(string destinationPath, bool overwrite, PathFormat pathFormat)
      {

         CopyToMoveToCore(destinationPath, overwrite ? CopyOptions.None : CopyOptions.FailIfExists, null, false, null, null, out var destinationPathLp, pathFormat);

         UpdateDestinationPath(destinationPath, destinationPathLp);

         return new FileInfo(Transaction, destinationPathLp, PathFormat.LongFullPath);
      }


      /// <summary>[AlphaFS] 既存のファイルを新しいファイルにコピーします。既存のファイルの上書きを許可し、<see cref="CopyOptions"/> を指定できます。</summary>
      /// <returns>完全修飾パスを持つ新しい <see cref="FileInfo"/> インスタンス。</returns>
      /// <remarks>
      ///   <para>非常に大きなファイル転送には、<see cref="CopyOptions.NoBuffering"/> オプションが推奨されます。</para>
      ///   <para>このメソッドを使用して、既存のファイルの上書きを許可または防止します。</para>
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
      /// <param name="destinationPath">コピー先の新しいファイルの名前。</param>
      /// <param name="copyOptions">ファイルのコピー方法を指定する <see cref="CopyOptions"/>。</param>
      [SecurityCritical]
      public FileInfo CopyTo(string destinationPath, CopyOptions copyOptions)
      {

         CopyToMoveToCore(destinationPath, copyOptions, null, false, null, null, out var destinationPathLp, PathFormat.RelativePath);

         UpdateDestinationPath(destinationPath, destinationPathLp);

         return new FileInfo(Transaction, destinationPathLp, PathFormat.LongFullPath);
      }

      
      /// <summary>[AlphaFS] 既存のファイルを新しいファイルにコピーします。既存のファイルの上書きを許可し、<see cref="CopyOptions"/> を指定できます。</summary>
      /// <returns>完全修飾パスを持つ新しい <see cref="FileInfo"/> インスタンス。</returns>
      /// <remarks>
      ///   <para>非常に大きなファイル転送には、<see cref="CopyOptions.NoBuffering"/> オプションが推奨されます。</para>
      ///   <para>このメソッドを使用して、既存のファイルの上書きを許可または防止します。</para>
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
      /// <param name="destinationPath">コピー先の新しいファイルの名前。</param>
      /// <param name="copyOptions">ファイルのコピー方法を指定する <see cref="CopyOptions"/>。</param>
      /// <param name="pathFormat">パスパラメーターの形式を示します。</param>
      [SecurityCritical]
      public FileInfo CopyTo(string destinationPath, CopyOptions copyOptions, PathFormat pathFormat)
      {

         CopyToMoveToCore(destinationPath, copyOptions, null, false, null, null, out var destinationPathLp, pathFormat);

         UpdateDestinationPath(destinationPath, destinationPathLp);

         return new FileInfo(Transaction, destinationPathLp, PathFormat.LongFullPath);
      }
      

      /// <summary>[AlphaFS] 既存のファイルを新しいファイルにコピーします。既存のファイルの上書きを許可し、<see cref="CopyOptions"/> を指定できます。</summary>
      /// <returns>完全修飾パスを持つ新しい <see cref="FileInfo"/> インスタンス。</returns>
      /// <remarks>
      ///   <para>非常に大きなファイル転送には、<see cref="CopyOptions.NoBuffering"/> オプションが推奨されます。</para>
      ///   <para>このメソッドを使用して、既存のファイルの上書きを許可または防止します。</para>
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
      /// <param name="destinationPath">コピー先の新しいファイルの名前。</param>
      /// <param name="copyOptions">ファイルのコピー方法を指定する <see cref="CopyOptions"/>。</param>
      /// <param name="preserveDates">元のタイムスタンプを保持する場合は <c>true</c>、それ以外の場合は <c>false</c>。</param>
      [SecurityCritical]
      public FileInfo CopyTo(string destinationPath, CopyOptions copyOptions, bool preserveDates)
      {

         CopyToMoveToCore(destinationPath, copyOptions, null, preserveDates, null, null, out var destinationPathLp, PathFormat.RelativePath);

         UpdateDestinationPath(destinationPath, destinationPathLp);

         return new FileInfo(Transaction, destinationPathLp, PathFormat.LongFullPath);
      }


      /// <summary>[AlphaFS] 既存のファイルを新しいファイルにコピーします。既存のファイルの上書きを許可し、<see cref="CopyOptions"/> を指定できます。</summary>
      /// <returns>完全修飾パスを持つ新しい <see cref="FileInfo"/> インスタンス。</returns>
      /// <remarks>
      ///   <para>非常に大きなファイル転送には、<see cref="CopyOptions.NoBuffering"/> オプションが推奨されます。</para>
      ///   <para>このメソッドを使用して、既存のファイルの上書きを許可または防止します。</para>
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
      /// <param name="destinationPath">コピー先の新しいファイルの名前。</param>
      /// <param name="copyOptions">ファイルのコピー方法を指定する <see cref="CopyOptions"/>。</param>
      /// <param name="preserveDates">元のタイムスタンプを保持する場合は <c>true</c>、それ以外の場合は <c>false</c>。</param>
      /// <param name="pathFormat">パスパラメーターの形式を示します。</param>
      [SecurityCritical]
      public FileInfo CopyTo(string destinationPath, CopyOptions copyOptions, bool preserveDates, PathFormat pathFormat)
      {

         CopyToMoveToCore(destinationPath, copyOptions, null, preserveDates, null, null, out var destinationPathLp, pathFormat);

         UpdateDestinationPath(destinationPath, destinationPathLp);

         return new FileInfo(Transaction, destinationPathLp, PathFormat.LongFullPath);
      }

      
      /// <summary>[AlphaFS] 既存のファイルを新しいファイルにコピーします。既存のファイルの上書きを許可し、<see cref="CopyOptions"/> を指定できます。</summary>
      /// <returns>コピー操作の詳細を含む <see cref="CopyMoveResult"/> クラス。</returns>
      ///   <para>コールバック関数を通じてアプリケーションに進行状況を通知できます。</para>
      /// <remarks>
      ///   <para>Option <see cref="CopyOptions.NoBuffering"/> is recommended for very large file transfers.</para>
      ///   <para>Use this method to allow or prevent overwriting of an existing file.</para>
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
      /// <param name="destinationPath">コピー先の新しいファイルの名前。</param>
      /// <param name="copyOptions">ファイルのコピー方法を指定する <see cref="CopyOptions"/>。</param>
      /// <param name="progressHandler">ファイルの別の部分がコピーされるたびに呼び出されるコールバック関数。このパラメーターは <c>null</c> にできます。</param>
      /// <param name="userProgressData">コールバック関数に渡される引数。このパラメーターは <c>null</c> にできます。</param>
      [SecurityCritical]
      public CopyMoveResult CopyTo(string destinationPath, CopyOptions copyOptions, CopyMoveProgressRoutine progressHandler, object userProgressData)
      {

         var cmr = CopyToMoveToCore(destinationPath, copyOptions, null, false, progressHandler, userProgressData, out var destinationPathLp, PathFormat.RelativePath);

         UpdateDestinationPath(destinationPath, destinationPathLp);

         return cmr;
      }


      /// <summary>[AlphaFS] 既存のファイルを新しいファイルにコピーします。既存のファイルの上書きを許可し、<see cref="CopyOptions"/> を指定できます。</summary>
      /// <returns>コピー操作の詳細を含む <see cref="CopyMoveResult"/> クラス。</returns>
      /// <remarks>
      ///   <para>Option <see cref="CopyOptions.NoBuffering"/> is recommended for very large file transfers.</para>
      ///   <para>Use this method to allow or prevent overwriting of an existing file.</para>
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
      /// <param name="destinationPath">コピー先の新しいファイルの名前。</param>
      /// <param name="copyOptions">ファイルのコピー方法を指定する <see cref="CopyOptions"/>。</param>
      /// <param name="progressHandler">ファイルの別の部分がコピーされるたびに呼び出されるコールバック関数。このパラメーターは <c>null</c> にできます。</param>
      /// <param name="userProgressData">コールバック関数に渡される引数。このパラメーターは <c>null</c> にできます。</param>
      /// <param name="pathFormat">パスパラメーターの形式を示します。</param>
      [SecurityCritical]
      public CopyMoveResult CopyTo(string destinationPath, CopyOptions copyOptions, CopyMoveProgressRoutine progressHandler, object userProgressData, PathFormat pathFormat)
      {

         var cmr = CopyToMoveToCore(destinationPath, copyOptions, null, false, progressHandler, userProgressData, out var destinationPathLp, pathFormat);

         UpdateDestinationPath(destinationPath, destinationPathLp);

         return cmr;
      }
      

      /// <summary>[AlphaFS] 既存のファイルを新しいファイルにコピーします。既存のファイルの上書きを許可し、<see cref="CopyOptions"/> を指定できます。</summary>
      /// <returns>コピー操作の詳細を含む <see cref="CopyMoveResult"/> クラス。</returns>
      ///   <para>コールバック関数を通じてアプリケーションに進行状況を通知できます。</para>
      /// <remarks>
      ///   <para>Option <see cref="CopyOptions.NoBuffering"/> is recommended for very large file transfers.</para>
      ///   <para>Use this method to allow or prevent overwriting of an existing file.</para>
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
      /// <param name="destinationPath">コピー先の新しいファイルの名前。</param>
      /// <param name="copyOptions">ファイルのコピー方法を指定する <see cref="CopyOptions"/>。</param>
      /// <param name="preserveDates">元のタイムスタンプを保持する場合は <c>true</c>、それ以外の場合は <c>false</c>。</param>
      /// <param name="progressHandler">ファイルの別の部分がコピーされるたびに呼び出されるコールバック関数。このパラメーターは <c>null</c> にできます。</param>
      /// <param name="userProgressData">コールバック関数に渡される引数。このパラメーターは <c>null</c> にできます。</param>
      [SecurityCritical]
      public CopyMoveResult CopyTo(string destinationPath, CopyOptions copyOptions, bool preserveDates, CopyMoveProgressRoutine progressHandler, object userProgressData)
      {

         var cmr = CopyToMoveToCore(destinationPath, copyOptions, null, preserveDates, progressHandler, userProgressData, out var destinationPathLp, PathFormat.RelativePath);

         UpdateDestinationPath(destinationPath, destinationPathLp);

         return cmr;
      }


      /// <summary>[AlphaFS] 既存のファイルを新しいファイルにコピーします。既存のファイルの上書きを許可し、<see cref="CopyOptions"/> を指定できます。</summary>
      /// <returns>コピー操作の詳細を含む <see cref="CopyMoveResult"/> クラス。</returns>
      ///   <para>コールバック関数を通じてアプリケーションに進行状況を通知できます。</para>
      /// <remarks>
      ///   <para>Option <see cref="CopyOptions.NoBuffering"/> is recommended for very large file transfers.</para>
      ///   <para>Use this method to allow or prevent overwriting of an existing file.</para>
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
      /// <param name="destinationPath">コピー先の新しいファイルの名前。</param>
      /// <param name="copyOptions">ファイルのコピー方法を指定する <see cref="CopyOptions"/>。</param>
      /// <param name="preserveDates">元のタイムスタンプを保持する場合は <c>true</c>、それ以外の場合は <c>false</c>。</param>
      /// <param name="progressHandler">ファイルの別の部分がコピーされるたびに呼び出されるコールバック関数。このパラメーターは <c>null</c> にできます。</param>
      /// <param name="userProgressData">コールバック関数に渡される引数。このパラメーターは <c>null</c> にできます。</param>
      /// <param name="pathFormat">パスパラメーターの形式を示します。</param>
      [SecurityCritical]
      public CopyMoveResult CopyTo(string destinationPath, CopyOptions copyOptions, bool preserveDates, CopyMoveProgressRoutine progressHandler, object userProgressData, PathFormat pathFormat)
      {

         var cmr = CopyToMoveToCore(destinationPath, copyOptions, null, preserveDates, progressHandler, userProgressData, out var destinationPathLp, pathFormat);

         UpdateDestinationPath(destinationPath, destinationPathLp);

         return cmr;
      }
   }
}
