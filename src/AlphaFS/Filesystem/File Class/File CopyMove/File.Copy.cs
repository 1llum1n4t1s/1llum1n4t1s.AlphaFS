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
   public static partial class File
   {
      #region Obsolete

      /// <summary>[AlphaFS] 既存のファイルを新しいファイルにコピーします。同名ファイルの上書きが許可されます。</summary>
      /// <remarks>
      ///   <para>元のファイルの属性はコピーされたファイルに保持されます。</para>
      ///   <para>可能な限り、このメソッドでは短いファイル名(<c>XXXXXX~1.XXX</c>など)の使用を避けてください。</para>
      ///   <para>2つのファイルの短いファイル名が同等の場合、このメソッドは失敗して例外を発生させるか、望ましくない動作になる可能性があります。</para>
      /// </remarks>
      /// <returns>コピー操作の詳細を含む<see cref="CopyMoveResult"/>クラスを返します。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="FileNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="sourcePath">コピーするファイル。 </param>
      /// <param name="destinationPath">コピー先ファイルの名前。ディレクトリは指定できません。</param>
      /// <param name="overwrite"><c>true</c>の場合、コピー先ファイルの読み取り専用属性と隠し属性を無視して上書きします。それ以外の場合は<c>false</c>。</param>
      /// <param name="progressHandler">ファイルの別の部分がコピーされるたびに呼び出されるコールバック関数。このパラメータは<c>null</c>にできます。</param>
      /// <param name="userProgressData">コールバック関数に渡される引数。このパラメータは<c>null</c>にできます。</param>
      [Obsolete("To disable/enable overwrite, use other overload and use CopyOptions.None enum flag or remove CopyOptions.FailIfExists enum flag.")]
      [SecurityCritical]
      public static CopyMoveResult Copy(string sourcePath, string destinationPath, bool overwrite, CopyMoveProgressRoutine progressHandler, object userProgressData)
      {
         return CopyMoveCore(false, new CopyMoveArguments
         {
            CopyOptions = overwrite ? CopyOptions.None : CopyOptions.FailIfExists,
            ProgressHandler = progressHandler,
            UserProgressData = userProgressData

         }, false, false, sourcePath, destinationPath, null);
      }


      /// <summary>[AlphaFS] 既存のファイルを新しいファイルにコピーします。同名ファイルの上書きが許可されます。</summary>
      /// <remarks>
      ///   <para>元のファイルの属性はコピーされたファイルに保持されます。</para>
      ///   <para>可能な限り、このメソッドでは短いファイル名(<c>XXXXXX~1.XXX</c>など)の使用を避けてください。</para>
      ///   <para>2つのファイルの短いファイル名が同等の場合、このメソッドは失敗して例外を発生させるか、望ましくない動作になる可能性があります。</para>
      /// </remarks>
      /// <returns>コピー操作の詳細を含む<see cref="CopyMoveResult"/>クラスを返します。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="FileNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="sourcePath">コピーするファイル。 </param>
      /// <param name="destinationPath">コピー先ファイルの名前。ディレクトリは指定できません。</param>
      /// <param name="overwrite"><c>true</c>の場合、コピー先ファイルの読み取り専用属性と隠し属性を無視して上書きします。それ以外の場合は<c>false</c>。</param>
      /// <param name="progressHandler">ファイルの別の部分がコピーされるたびに呼び出されるコールバック関数。このパラメータは<c>null</c>にできます。</param>
      /// <param name="userProgressData">コールバック関数に渡される引数。このパラメータは<c>null</c>にできます。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [Obsolete("To disable/enable overwrite, use other overload and use CopyOptions.None enum flag or remove CopyOptions.FailIfExists enum flag.")]
      [SecurityCritical]
      public static CopyMoveResult Copy(string sourcePath, string destinationPath, bool overwrite, CopyMoveProgressRoutine progressHandler, object userProgressData, PathFormat pathFormat)
      {
         return CopyMoveCore(false, new CopyMoveArguments
         {
            CopyOptions = overwrite ? CopyOptions.None : CopyOptions.FailIfExists,
            ProgressHandler = progressHandler,
            UserProgressData = userProgressData,
            PathFormat = pathFormat

         }, false, false, sourcePath, destinationPath, null);
      }
      

      /// <summary>[AlphaFS] 既存のファイルを新しいファイルにコピーします。同名ファイルの上書きが許可されます。 <see cref="CopyOptions"/> can be specified.</summary>
      /// <remarks>
      ///   <para>非常に大きなファイル転送には<see cref="CopyOptions.NoBuffering"/>オプションが推奨されます。</para>
      ///   <para>元のファイルの属性はコピーされたファイルに保持されます。</para>
      ///   <para>可能な限り、このメソッドでは短いファイル名(<c>XXXXXX~1.XXX</c>など)の使用を避けてください。</para>
      ///   <para>2つのファイルの短いファイル名が同等の場合、このメソッドは失敗して例外を発生させるか、望ましくない動作になる可能性があります。</para>
      /// </remarks>
      /// <returns>コピー操作の詳細を含む<see cref="CopyMoveResult"/>クラスを返します。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="FileNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="sourcePath">コピーするファイル。</param>
      /// <param name="destinationPath">コピー先ファイルの名前。ディレクトリは指定できません。</param>
      /// <param name="copyOptions">ファイルのコピー方法を指定する<see cref="CopyOptions"/>。このパラメータは<c>null</c>にできます。</param>
      /// <param name="preserveDates">元のタイムスタンプを保持する必要がある場合は<c>true</c>。それ以外の場合は<c>false</c>。</param>
      [SecurityCritical]
      [Obsolete("Use other overload and add CopyOptions.CopyTimestamp enum flag.")]
      public static CopyMoveResult Copy(string sourcePath, string destinationPath, CopyOptions copyOptions, bool preserveDates)
      {
         return CopyMoveCore(false, new CopyMoveArguments
         {
            CopyOptions = preserveDates ? copyOptions | CopyOptions.CopyTimestamp : copyOptions & ~CopyOptions.CopyTimestamp

         }, false, false, sourcePath, destinationPath, null);
      }


      /// <summary>[AlphaFS] 既存のファイルを新しいファイルにコピーします。同名ファイルの上書きが許可されます。 <see cref="CopyOptions"/> can be specified.</summary>
      /// <remarks>
      ///   <para>非常に大きなファイル転送には<see cref="CopyOptions.NoBuffering"/>オプションが推奨されます。</para>
      ///   <para>元のファイルの属性はコピーされたファイルに保持されます。</para>
      ///   <para>可能な限り、このメソッドでは短いファイル名(<c>XXXXXX~1.XXX</c>など)の使用を避けてください。</para>
      ///   <para>2つのファイルの短いファイル名が同等の場合、このメソッドは失敗して例外を発生させるか、望ましくない動作になる可能性があります。</para>
      /// </remarks>
      /// <returns>コピー操作の詳細を含む<see cref="CopyMoveResult"/>クラスを返します。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="FileNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="sourcePath">コピーするファイル。</param>
      /// <param name="destinationPath">コピー先ファイルの名前。ディレクトリは指定できません。</param>
      /// <param name="copyOptions">ファイルのコピー方法を指定する<see cref="CopyOptions"/>。このパラメータは<c>null</c>にできます。</param>
      /// <param name="preserveDates">元のタイムスタンプを保持する必要がある場合は<c>true</c>。それ以外の場合は<c>false</c>。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      [Obsolete("Use other overload and add CopyOptions.CopyTimestamp enum flag.")]
      public static CopyMoveResult Copy(string sourcePath, string destinationPath, CopyOptions copyOptions, bool preserveDates, PathFormat pathFormat)
      {
         return CopyMoveCore(false, new CopyMoveArguments
         {
            CopyOptions = preserveDates ? copyOptions | CopyOptions.CopyTimestamp : copyOptions & ~CopyOptions.CopyTimestamp,
            PathFormat = pathFormat

         }, false, false, sourcePath, destinationPath, null);
      }


      /// <summary>[AlphaFS] 既存のファイルを新しいファイルにコピーします。同名ファイルの上書きが許可されます。  <see cref="CopyOptions"/> can be specified,
      /// およびコールバック関数を通じてアプリケーションに進行状況を通知する可能性があります。
      /// </summary>
      /// <remarks>
      ///   <para>非常に大きなファイル転送には<see cref="CopyOptions.NoBuffering"/>オプションが推奨されます。</para>
      ///   <para>元のファイルの属性はコピーされたファイルに保持されます。</para>
      ///   <para>可能な限り、このメソッドでは短いファイル名(<c>XXXXXX~1.XXX</c>など)の使用を避けてください。</para>
      ///   <para>2つのファイルの短いファイル名が同等の場合、このメソッドは失敗して例外を発生させるか、望ましくない動作になる可能性があります。</para>
      /// </remarks>
      /// <returns>コピー操作の詳細を含む<see cref="CopyMoveResult"/>クラスを返します。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="FileNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="sourcePath">コピーするファイル。</param>
      /// <param name="destinationPath">コピー先ファイルの名前。ディレクトリは指定できません。</param>
      /// <param name="copyOptions">ファイルのコピー方法を指定する<see cref="CopyOptions"/>。このパラメータは<c>null</c>にできます。</param>
      /// <param name="preserveDates">元のタイムスタンプを保持する必要がある場合は<c>true</c>。それ以外の場合は<c>false</c>。</param>
      /// <param name="progressHandler">ファイルの別の部分がコピーされるたびに呼び出されるコールバック関数。このパラメータは<c>null</c>にできます。</param>
      /// <param name="userProgressData">コールバック関数に渡される引数。このパラメータは<c>null</c>にできます。</param>
      [SecurityCritical]
      [Obsolete("Use other overload and add CopyOptions.CopyTimestamp enum flag.")]
      public static CopyMoveResult Copy(string sourcePath, string destinationPath, CopyOptions copyOptions, bool preserveDates, CopyMoveProgressRoutine progressHandler, object userProgressData)
      {
         return CopyMoveCore(false, new CopyMoveArguments
         {
            CopyOptions = preserveDates ? copyOptions | CopyOptions.CopyTimestamp : copyOptions & ~CopyOptions.CopyTimestamp,
            ProgressHandler = progressHandler,
            UserProgressData = userProgressData

         }, false, false, sourcePath, destinationPath, null);
      }


      /// <summary>[AlphaFS] 既存のファイルを新しいファイルにコピーします。同名ファイルの上書きが許可されます。  <see cref="CopyOptions"/> can be specified,
      /// およびコールバック関数を通じてアプリケーションに進行状況を通知する可能性があります。
      /// </summary>
      /// <remarks>
      ///   <para>非常に大きなファイル転送には<see cref="CopyOptions.NoBuffering"/>オプションが推奨されます。</para>
      ///   <para>元のファイルの属性はコピーされたファイルに保持されます。</para>
      ///   <para>可能な限り、このメソッドでは短いファイル名(<c>XXXXXX~1.XXX</c>など)の使用を避けてください。</para>
      ///   <para>2つのファイルの短いファイル名が同等の場合、このメソッドは失敗して例外を発生させるか、望ましくない動作になる可能性があります。</para>
      /// </remarks>
      /// <returns>コピー操作の詳細を含む<see cref="CopyMoveResult"/>クラスを返します。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="FileNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="sourcePath">コピーするファイル。</param>
      /// <param name="destinationPath">コピー先ファイルの名前。ディレクトリは指定できません。</param>
      /// <param name="copyOptions">ファイルのコピー方法を指定する<see cref="CopyOptions"/>。このパラメータは<c>null</c>にできます。</param>
      /// <param name="preserveDates">元のタイムスタンプを保持する必要がある場合は<c>true</c>。それ以外の場合は<c>false</c>。</param>
      /// <param name="progressHandler">ファイルの別の部分がコピーされるたびに呼び出されるコールバック関数。このパラメータは<c>null</c>にできます。</param>
      /// <param name="userProgressData">コールバック関数に渡される引数。このパラメータは<c>null</c>にできます。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      [Obsolete("Use other overload and add CopyOptions.CopyTimestamp enum flag.")]
      public static CopyMoveResult Copy(string sourcePath, string destinationPath, CopyOptions copyOptions, bool preserveDates, CopyMoveProgressRoutine progressHandler, object userProgressData, PathFormat pathFormat)
      {
         return CopyMoveCore(false, new CopyMoveArguments
         {
            CopyOptions = preserveDates ? copyOptions | CopyOptions.CopyTimestamp : copyOptions & ~CopyOptions.CopyTimestamp,
            ProgressHandler = progressHandler,
            UserProgressData = userProgressData,
            PathFormat = pathFormat

         }, false, false, sourcePath, destinationPath, null);
      }

      #endregion // Obsolete

      
      #region .NET

      /// <summary>既存のファイルを新しいファイルにコピーします。同名ファイルの上書きは許可されません。</summary>
      /// <remarks>
      ///   <para>元のファイルの属性はコピーされたファイルに保持されます。</para>
      ///   <para>可能な限り、このメソッドでは短いファイル名(<c>XXXXXX~1.XXX</c>など)の使用を避けてください。</para>
      ///   <para>2つのファイルの短いファイル名が同等の場合、このメソッドは失敗して例外を発生させるか、望ましくない動作になる可能性があります。</para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="FileNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="sourcePath">コピーするファイル。</param>
      /// <param name="destinationPath">コピー先ファイルの名前。ディレクトリまたは既存のファイルは指定できません。</param>
      [SecurityCritical]
      public static void Copy(string sourcePath, string destinationPath)
      {
         CopyMoveCore(false, new CopyMoveArguments
         {
            CopyOptions = CopyOptions.FailIfExists

         }, false, false, sourcePath, destinationPath, null);
      }


      /// <summary>Copies an existing file to a new file. Overwriting a file of the same name is allowed.</summary>
      /// <remarks>
      ///   <para>元のファイルの属性はコピーされたファイルに保持されます。</para>
      ///   <para>可能な限り、このメソッドでは短いファイル名(<c>XXXXXX~1.XXX</c>など)の使用を避けてください。</para>
      ///   <para>2つのファイルの短いファイル名が同等の場合、このメソッドは失敗して例外を発生させるか、望ましくない動作になる可能性があります。</para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="FileNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="sourcePath">コピーするファイル。 </param>
      /// <param name="destinationPath">コピー先ファイルの名前。ディレクトリは指定できません。</param>
      /// <param name="overwrite"><c>true</c>の場合、コピー先ファイルの読み取り専用属性と隠し属性を無視して上書きします。それ以外の場合は<c>false</c>。</param>      
      [SecurityCritical]
      public static void Copy(string sourcePath, string destinationPath, bool overwrite)
      {
         CopyMoveCore(false, new CopyMoveArguments
         {
            CopyOptions = overwrite ? CopyOptions.None : CopyOptions.FailIfExists

         }, false, false, sourcePath, destinationPath, null);
      }

      #endregion // .NET


      /// <summary>[AlphaFS] 既存のファイルを新しいファイルにコピーします。同名ファイルの上書きは許可されません。</summary>
      /// <remarks>
      ///   <para>元のファイルの属性はコピーされたファイルに保持されます。</para>
      ///   <para>可能な限り、このメソッドでは短いファイル名(<c>XXXXXX~1.XXX</c>など)の使用を避けてください。</para>
      ///   <para>2つのファイルの短いファイル名が同等の場合、このメソッドは失敗して例外を発生させるか、望ましくない動作になる可能性があります。</para>
      /// </remarks>
      /// <returns>コピー操作の詳細を含む<see cref="CopyMoveResult"/>クラスを返します。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="FileNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="sourcePath">コピーするファイル。 </param>
      /// <param name="destinationPath">コピー先ファイルの名前。ディレクトリは指定できません。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static CopyMoveResult Copy(string sourcePath, string destinationPath, PathFormat pathFormat)
      {
         return CopyMoveCore(false, new CopyMoveArguments
         {
            CopyOptions = CopyOptions.FailIfExists,
            PathFormat = pathFormat

         }, false, false, sourcePath, destinationPath, null);
      }

      /// <summary>
      /// [AlphaFS] 既存のファイルを新しいファイルにコピーします。 Overwriting a file of the same name is
      /// 許可されます。
      /// </summary>
      /// <remarks>
      /// <para>元のファイルの属性はコピーされたファイルに保持されます。</para>
      /// <para>Whenever possible, avoid using short file names (such as <c>XXXXXX~1.XXX</c>) with this
      /// method.</para>
      /// <para>If two files have equivalent short file names then this method may fail and raise an
      /// exception and/or result in undesirable behavior.</para>
      /// </remarks>
      /// <exception cref="ArgumentException">.</exception>
      /// <exception cref="ArgumentNullException">.</exception>
      /// <exception cref="DirectoryNotFoundException">.</exception>
      /// <exception cref="FileNotFoundException">.</exception>
      /// <exception cref="IOException">.</exception>
      /// <exception cref="NotSupportedException">.</exception>
      /// <exception cref="UnauthorizedAccessException">.</exception>
      /// <param name="sourcePath">コピーするファイル。</param>
      /// <param name="destinationPath">コピー先ファイルの名前。ディレクトリは指定できません。</param>
      /// <param name="overwrite"><c>true</c> if the destination file should ignoring the read-only and
      /// hidden attributes and overwrite; otherwise, <c>false</c>.</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      /// <returns>
      /// コピー操作の詳細を含む<see cref="CopyMoveResult"/>クラスを返します。
      /// </returns>
      [SecurityCritical]
      public static CopyMoveResult Copy(string sourcePath, string destinationPath, bool overwrite, PathFormat pathFormat)
      {
         return CopyMoveCore(true, new CopyMoveArguments
         {
            CopyOptions = overwrite ? CopyOptions.None : CopyOptions.FailIfExists,
            PathFormat = pathFormat

         }, false, false, sourcePath, destinationPath, null);
      }


      /// <summary>[AlphaFS] 既存のファイルを新しいファイルにコピーします。同名ファイルの上書きは許可されません。</summary>
      /// <remarks>
      ///   <para>元のファイルの属性はコピーされたファイルに保持されます。</para>
      ///   <para>可能な限り、このメソッドでは短いファイル名(<c>XXXXXX~1.XXX</c>など)の使用を避けてください。</para>
      ///   <para>2つのファイルの短いファイル名が同等の場合、このメソッドは失敗して例外を発生させるか、望ましくない動作になる可能性があります。</para>
      /// </remarks>
      /// <returns>コピー操作の詳細を含む<see cref="CopyMoveResult"/>クラスを返します。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="FileNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="sourcePath">コピーするファイル。</param>
      /// <param name="destinationPath">コピー先ファイルの名前。ディレクトリまたは既存のファイルは指定できません。</param>
      /// <param name="retry">コピー失敗時のリトライ回数。</param>
      /// <param name="retryTimeout">リトライ間の待機時間(秒)。</param>
      [SecurityCritical]
      public static CopyMoveResult Copy(string sourcePath, string destinationPath, int retry, int retryTimeout)
      {
         return CopyMoveCore(true, new CopyMoveArguments
         {
            Retry = retry,
            RetryTimeout = retryTimeout,
            CopyOptions = CopyOptions.FailIfExists

         }, false, false, sourcePath, destinationPath, null);
      }


      /// <summary>[AlphaFS] 既存のファイルを新しいファイルにコピーします。同名ファイルの上書きは許可されません。</summary>
      /// <remarks>
      ///   <para>元のファイルの属性はコピーされたファイルに保持されます。</para>
      ///   <para>可能な限り、このメソッドでは短いファイル名(<c>XXXXXX~1.XXX</c>など)の使用を避けてください。</para>
      ///   <para>2つのファイルの短いファイル名が同等の場合、このメソッドは失敗して例外を発生させるか、望ましくない動作になる可能性があります。</para>
      /// </remarks>
      /// <returns>コピー操作の詳細を含む<see cref="CopyMoveResult"/>クラスを返します。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="FileNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="sourcePath">コピーするファイル。 </param>
      /// <param name="destinationPath">コピー先ファイルの名前。ディレクトリは指定できません。</param>
      /// <param name="retry">コピー失敗時のリトライ回数。</param>
      /// <param name="retryTimeout">リトライ間の待機時間(秒)。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static CopyMoveResult Copy(string sourcePath, string destinationPath, int retry, int retryTimeout, PathFormat pathFormat)
      {
         return CopyMoveCore(true, new CopyMoveArguments
         {
            Retry = retry,
            RetryTimeout = retryTimeout,
            CopyOptions = CopyOptions.FailIfExists,
            PathFormat = pathFormat

         }, false, false, sourcePath, destinationPath, null);
      }


      /// <summary>[AlphaFS] 既存のファイルを新しいファイルにコピーします。同名ファイルの上書きは許可されません。 Possibility of notifying the application of its progress through a callback function.</summary>
      /// <returns>コピー操作の詳細を含む<see cref="CopyMoveResult"/>クラスを返します。</returns>
      /// <remarks>
      ///   <para>元のファイルの属性はコピーされたファイルに保持されます。</para>
      ///   <para>可能な限り、このメソッドでは短いファイル名(<c>XXXXXX~1.XXX</c>など)の使用を避けてください。</para>
      ///   <para>2つのファイルの短いファイル名が同等の場合、このメソッドは失敗して例外を発生させるか、望ましくない動作になる可能性があります。</para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="FileNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="sourcePath">コピーするファイル。 </param>
      /// <param name="destinationPath">コピー先ファイルの名前。ディレクトリは指定できません。</param>
      /// <param name="progressHandler">ファイルの別の部分がコピーされるたびに呼び出されるコールバック関数。このパラメータは<c>null</c>にできます。</param>
      /// <param name="userProgressData">コールバック関数に渡される引数。このパラメータは<c>null</c>にできます。</param>
      [SecurityCritical]
      public static CopyMoveResult Copy(string sourcePath, string destinationPath, CopyMoveProgressRoutine progressHandler, object userProgressData)
      {
         return CopyMoveCore(false, new CopyMoveArguments
         {
            CopyOptions = CopyOptions.FailIfExists,
            ProgressHandler = progressHandler,
            UserProgressData = userProgressData

         }, false, false, sourcePath, destinationPath, null);
      }


      /// <summary>[AlphaFS] 既存のファイルを新しいファイルにコピーします。同名ファイルの上書きは許可されません。 Possibility of notifying the application of its progress through a callback function.</summary>
      /// <returns>コピー操作の詳細を含む<see cref="CopyMoveResult"/>クラスを返します。</returns>
      /// <remarks>
      ///   <para>元のファイルの属性はコピーされたファイルに保持されます。</para>
      ///   <para>可能な限り、このメソッドでは短いファイル名(<c>XXXXXX~1.XXX</c>など)の使用を避けてください。</para>
      ///   <para>2つのファイルの短いファイル名が同等の場合、このメソッドは失敗して例外を発生させるか、望ましくない動作になる可能性があります。</para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="FileNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="sourcePath">コピーするファイル。 </param>
      /// <param name="destinationPath">コピー先ファイルの名前。ディレクトリは指定できません。</param>
      /// <param name="progressHandler">ファイルの別の部分がコピーされるたびに呼び出されるコールバック関数。このパラメータは<c>null</c>にできます。</param>
      /// <param name="userProgressData">コールバック関数に渡される引数。このパラメータは<c>null</c>にできます。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static CopyMoveResult Copy(string sourcePath, string destinationPath, CopyMoveProgressRoutine progressHandler, object userProgressData, PathFormat pathFormat)
      {
         return CopyMoveCore(false, new CopyMoveArguments
         {
            CopyOptions = CopyOptions.FailIfExists,
            ProgressHandler = progressHandler,
            UserProgressData = userProgressData,
            PathFormat = pathFormat

         }, false, false, sourcePath, destinationPath, null);
      }


      /// <summary>[AlphaFS] 既存のファイルを新しいファイルにコピーします。同名ファイルの上書きは許可されません。 Possibility of notifying the application of its progress through a callback function.</summary>
      /// <returns>コピー操作の詳細を含む<see cref="CopyMoveResult"/>クラスを返します。</returns>
      /// <remarks>
      ///   <para>元のファイルの属性はコピーされたファイルに保持されます。</para>
      ///   <para>可能な限り、このメソッドでは短いファイル名(<c>XXXXXX~1.XXX</c>など)の使用を避けてください。</para>
      ///   <para>2つのファイルの短いファイル名が同等の場合、このメソッドは失敗して例外を発生させるか、望ましくない動作になる可能性があります。</para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="FileNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="sourcePath">コピーするファイル。 </param>
      /// <param name="destinationPath">コピー先ファイルの名前。ディレクトリは指定できません。</param>
      /// <param name="retry">コピー失敗時のリトライ回数。</param>
      /// <param name="retryTimeout">リトライ間の待機時間(秒)。</param>
      /// <param name="progressHandler">ファイルの別の部分がコピーされるたびに呼び出されるコールバック関数。このパラメータは<c>null</c>にできます。</param>
      /// <param name="userProgressData">コールバック関数に渡される引数。このパラメータは<c>null</c>にできます。</param>
      [SecurityCritical]
      public static CopyMoveResult Copy(string sourcePath, string destinationPath, int retry, int retryTimeout, CopyMoveProgressRoutine progressHandler, object userProgressData)
      {
         return CopyMoveCore(false, new CopyMoveArguments
         {
            Retry = retry,
            RetryTimeout = retryTimeout,
            CopyOptions = CopyOptions.FailIfExists,
            ProgressHandler = progressHandler,
            UserProgressData = userProgressData

         }, false, false, sourcePath, destinationPath, null);
      }


      /// <summary>[AlphaFS] 既存のファイルを新しいファイルにコピーします。同名ファイルの上書きは許可されません。 Possibility of notifying the application of its progress through a callback function.</summary>
      /// <returns>コピー操作の詳細を含む<see cref="CopyMoveResult"/>クラスを返します。</returns>
      /// <remarks>
      ///   <para>元のファイルの属性はコピーされたファイルに保持されます。</para>
      ///   <para>可能な限り、このメソッドでは短いファイル名(<c>XXXXXX~1.XXX</c>など)の使用を避けてください。</para>
      ///   <para>2つのファイルの短いファイル名が同等の場合、このメソッドは失敗して例外を発生させるか、望ましくない動作になる可能性があります。</para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="FileNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="sourcePath">コピーするファイル。 </param>
      /// <param name="destinationPath">コピー先ファイルの名前。ディレクトリは指定できません。</param>
      /// <param name="retry">コピー失敗時のリトライ回数。</param>
      /// <param name="retryTimeout">リトライ間の待機時間(秒)。</param>
      /// <param name="progressHandler">ファイルの別の部分がコピーされるたびに呼び出されるコールバック関数。このパラメータは<c>null</c>にできます。</param>
      /// <param name="userProgressData">コールバック関数に渡される引数。このパラメータは<c>null</c>にできます。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static CopyMoveResult Copy(string sourcePath, string destinationPath, int retry, int retryTimeout, CopyMoveProgressRoutine progressHandler, object userProgressData, PathFormat pathFormat)
      {
         return CopyMoveCore(false, new CopyMoveArguments
         {
            Retry = retry,
            RetryTimeout = retryTimeout,
            CopyOptions = CopyOptions.FailIfExists,
            ProgressHandler = progressHandler,
            UserProgressData = userProgressData,
            PathFormat = pathFormat

         }, false, false, sourcePath, destinationPath, null);
      }
      



      /// <summary>[AlphaFS] 既存のファイルを新しいファイルにコピーします。同名ファイルの上書きが許可されます。 <see cref="CopyOptions"/> can be specified.</summary>
      /// <remarks>
      ///   <para>非常に大きなファイル転送には<see cref="CopyOptions.NoBuffering"/>オプションが推奨されます。</para>
      ///   <para>元のファイルの属性はコピーされたファイルに保持されます。</para>
      ///   <para>可能な限り、このメソッドでは短いファイル名(<c>XXXXXX~1.XXX</c>など)の使用を避けてください。</para>
      ///   <para>2つのファイルの短いファイル名が同等の場合、このメソッドは失敗して例外を発生させるか、望ましくない動作になる可能性があります。</para>
      /// </remarks>
      /// <returns>コピー操作の詳細を含む<see cref="CopyMoveResult"/>クラスを返します。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="FileNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="sourcePath">コピーするファイル。</param>
      /// <param name="destinationPath">コピー先ファイルの名前。ディレクトリは指定できません。</param>
      /// <param name="copyOptions">ファイルのコピー方法を指定する<see cref="CopyOptions"/>。このパラメータは<c>null</c>にできます。</param>
      [SecurityCritical]
      public static CopyMoveResult Copy(string sourcePath, string destinationPath, CopyOptions copyOptions)
      {
         return CopyMoveCore(false, new CopyMoveArguments
         {
            CopyOptions = copyOptions

         }, false, false, sourcePath, destinationPath, null);
      }


      /// <summary>[AlphaFS] 既存のファイルを新しいファイルにコピーします。同名ファイルの上書きが許可されます。 <see cref="CopyOptions"/> can be specified.</summary>
      /// <remarks>
      ///   <para>非常に大きなファイル転送には<see cref="CopyOptions.NoBuffering"/>オプションが推奨されます。</para>
      ///   <para>元のファイルの属性はコピーされたファイルに保持されます。</para>
      ///   <para>可能な限り、このメソッドでは短いファイル名(<c>XXXXXX~1.XXX</c>など)の使用を避けてください。</para>
      ///   <para>2つのファイルの短いファイル名が同等の場合、このメソッドは失敗して例外を発生させるか、望ましくない動作になる可能性があります。</para>
      /// </remarks>
      /// <returns>コピー操作の詳細を含む<see cref="CopyMoveResult"/>クラスを返します。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="FileNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="sourcePath">コピーするファイル。</param>
      /// <param name="destinationPath">コピー先ファイルの名前。ディレクトリは指定できません。</param>
      /// <param name="copyOptions">ファイルのコピー方法を指定する<see cref="CopyOptions"/>。このパラメータは<c>null</c>にできます。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static CopyMoveResult Copy(string sourcePath, string destinationPath, CopyOptions copyOptions, PathFormat pathFormat)
      {
         return CopyMoveCore(false, new CopyMoveArguments
         {
            CopyOptions = copyOptions,
            PathFormat = pathFormat

         }, false, false, sourcePath, destinationPath, null);
      }
      

      /// <summary>[AlphaFS] 既存のファイルを新しいファイルにコピーします。同名ファイルの上書きが許可されます。 <see cref="CopyOptions"/> can be specified.</summary>
      /// <remarks>
      ///   <para>非常に大きなファイル転送には<see cref="CopyOptions.NoBuffering"/>オプションが推奨されます。</para>
      ///   <para>元のファイルの属性はコピーされたファイルに保持されます。</para>
      ///   <para>可能な限り、このメソッドでは短いファイル名(<c>XXXXXX~1.XXX</c>など)の使用を避けてください。</para>
      ///   <para>2つのファイルの短いファイル名が同等の場合、このメソッドは失敗して例外を発生させるか、望ましくない動作になる可能性があります。</para>
      /// </remarks>
      /// <returns>コピー操作の詳細を含む<see cref="CopyMoveResult"/>クラスを返します。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="FileNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="sourcePath">コピーするファイル。</param>
      /// <param name="destinationPath">コピー先ファイルの名前。ディレクトリは指定できません。</param>
      /// <param name="copyOptions">ファイルのコピー方法を指定する<see cref="CopyOptions"/>。このパラメータは<c>null</c>にできます。</param>
      /// <param name="retry">コピー失敗時のリトライ回数。</param>
      /// <param name="retryTimeout">リトライ間の待機時間(秒)。</param>
      [SecurityCritical]
      public static CopyMoveResult Copy(string sourcePath, string destinationPath, CopyOptions copyOptions, int retry, int retryTimeout)
      {
         return CopyMoveCore(true, new CopyMoveArguments
         {
            Retry = retry,
            RetryTimeout = retryTimeout,
            CopyOptions = copyOptions

         }, false, false, sourcePath, destinationPath, null);
      }


      /// <summary>[AlphaFS] 既存のファイルを新しいファイルにコピーします。同名ファイルの上書きが許可されます。 <see cref="CopyOptions"/> can be specified.</summary>
      /// <remarks>
      ///   <para>非常に大きなファイル転送には<see cref="CopyOptions.NoBuffering"/>オプションが推奨されます。</para>
      ///   <para>元のファイルの属性はコピーされたファイルに保持されます。</para>
      ///   <para>可能な限り、このメソッドでは短いファイル名(<c>XXXXXX~1.XXX</c>など)の使用を避けてください。</para>
      ///   <para>2つのファイルの短いファイル名が同等の場合、このメソッドは失敗して例外を発生させるか、望ましくない動作になる可能性があります。</para>
      /// </remarks>
      /// <returns>コピー操作の詳細を含む<see cref="CopyMoveResult"/>クラスを返します。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="FileNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="sourcePath">コピーするファイル。</param>
      /// <param name="destinationPath">コピー先ファイルの名前。ディレクトリは指定できません。</param>
      /// <param name="copyOptions">ファイルのコピー方法を指定する<see cref="CopyOptions"/>。このパラメータは<c>null</c>にできます。</param>
      /// <param name="retry">コピー失敗時のリトライ回数。</param>
      /// <param name="retryTimeout">リトライ間の待機時間(秒)。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static CopyMoveResult Copy(string sourcePath, string destinationPath, CopyOptions copyOptions, int retry, int retryTimeout, PathFormat pathFormat)
      {
         return CopyMoveCore(true, new CopyMoveArguments
         {
            Retry = retry,
            RetryTimeout = retryTimeout,
            CopyOptions = copyOptions,
            PathFormat = pathFormat

         }, false, false, sourcePath, destinationPath, null);
      }
      

      /// <summary>[AlphaFS] 既存のファイルを新しいファイルにコピーします。同名ファイルの上書きが許可されます。  <see cref="CopyOptions"/> can be specified,
      /// およびコールバック関数を通じてアプリケーションに進行状況を通知する可能性があります。
      /// </summary>
      /// <remarks>
      ///   <para>非常に大きなファイル転送には<see cref="CopyOptions.NoBuffering"/>オプションが推奨されます。</para>
      ///   <para>元のファイルの属性はコピーされたファイルに保持されます。</para>
      ///   <para>可能な限り、このメソッドでは短いファイル名(<c>XXXXXX~1.XXX</c>など)の使用を避けてください。</para>
      ///   <para>2つのファイルの短いファイル名が同等の場合、このメソッドは失敗して例外を発生させるか、望ましくない動作になる可能性があります。</para>
      /// </remarks>
      /// <returns>コピー操作の詳細を含む<see cref="CopyMoveResult"/>クラスを返します。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="FileNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="sourcePath">コピーするファイル。</param>
      /// <param name="destinationPath">コピー先ファイルの名前。ディレクトリは指定できません。</param>
      /// <param name="copyOptions">ファイルのコピー方法を指定する<see cref="CopyOptions"/>。このパラメータは<c>null</c>にできます。</param>
      /// <param name="progressHandler">ファイルの別の部分がコピーされるたびに呼び出されるコールバック関数。このパラメータは<c>null</c>にできます。</param>
      /// <param name="userProgressData">コールバック関数に渡される引数。このパラメータは<c>null</c>にできます。</param>
      [SecurityCritical]
      public static CopyMoveResult Copy(string sourcePath, string destinationPath, CopyOptions copyOptions, CopyMoveProgressRoutine progressHandler, object userProgressData)
      {
         return CopyMoveCore(false, new CopyMoveArguments
         {
            CopyOptions = copyOptions,
            ProgressHandler = progressHandler,
            UserProgressData = userProgressData

         }, false, false, sourcePath, destinationPath, null);
      }


      /// <summary>[AlphaFS] 既存のファイルを新しいファイルにコピーします。同名ファイルの上書きが許可されます。  <see cref="CopyOptions"/> can be specified,
      /// およびコールバック関数を通じてアプリケーションに進行状況を通知する可能性があります。
      /// </summary>
      /// <remarks>
      ///   <para>非常に大きなファイル転送には<see cref="CopyOptions.NoBuffering"/>オプションが推奨されます。</para>
      ///   <para>元のファイルの属性はコピーされたファイルに保持されます。</para>
      ///   <para>可能な限り、このメソッドでは短いファイル名(<c>XXXXXX~1.XXX</c>など)の使用を避けてください。</para>
      ///   <para>2つのファイルの短いファイル名が同等の場合、このメソッドは失敗して例外を発生させるか、望ましくない動作になる可能性があります。</para>
      /// </remarks>
      /// <returns>コピー操作の詳細を含む<see cref="CopyMoveResult"/>クラスを返します。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="FileNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="sourcePath">コピーするファイル。</param>
      /// <param name="destinationPath">コピー先ファイルの名前。ディレクトリは指定できません。</param>
      /// <param name="copyOptions">ファイルのコピー方法を指定する<see cref="CopyOptions"/>。このパラメータは<c>null</c>にできます。</param>
      /// <param name="progressHandler">ファイルの別の部分がコピーされるたびに呼び出されるコールバック関数。このパラメータは<c>null</c>にできます。</param>
      /// <param name="userProgressData">コールバック関数に渡される引数。このパラメータは<c>null</c>にできます。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static CopyMoveResult Copy(string sourcePath, string destinationPath, CopyOptions copyOptions, CopyMoveProgressRoutine progressHandler, object userProgressData, PathFormat pathFormat)
      {
         return CopyMoveCore(false, new CopyMoveArguments
         {
            CopyOptions = copyOptions,
            ProgressHandler = progressHandler,
            UserProgressData = userProgressData,
            PathFormat = pathFormat

         }, false, false, sourcePath, destinationPath, null);
      }
      

      /// <summary>[AlphaFS] 既存のファイルを新しいファイルにコピーします。同名ファイルの上書きが許可されます。  <see cref="CopyOptions"/> can be specified,
      /// およびコールバック関数を通じてアプリケーションに進行状況を通知する可能性があります。
      /// </summary>
      /// <remarks>
      ///   <para>非常に大きなファイル転送には<see cref="CopyOptions.NoBuffering"/>オプションが推奨されます。</para>
      ///   <para>元のファイルの属性はコピーされたファイルに保持されます。</para>
      ///   <para>可能な限り、このメソッドでは短いファイル名(<c>XXXXXX~1.XXX</c>など)の使用を避けてください。</para>
      ///   <para>2つのファイルの短いファイル名が同等の場合、このメソッドは失敗して例外を発生させるか、望ましくない動作になる可能性があります。</para>
      /// </remarks>
      /// <returns>コピー操作の詳細を含む<see cref="CopyMoveResult"/>クラスを返します。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="FileNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="sourcePath">コピーするファイル。</param>
      /// <param name="destinationPath">コピー先ファイルの名前。ディレクトリは指定できません。</param>
      /// <param name="copyOptions">ファイルのコピー方法を指定する<see cref="CopyOptions"/>。このパラメータは<c>null</c>にできます。</param>
      /// <param name="retry">コピー失敗時のリトライ回数。</param>
      /// <param name="retryTimeout">リトライ間の待機時間(秒)。</param>
      /// <param name="progressHandler">ファイルの別の部分がコピーされるたびに呼び出されるコールバック関数。このパラメータは<c>null</c>にできます。</param>
      /// <param name="userProgressData">コールバック関数に渡される引数。このパラメータは<c>null</c>にできます。</param>
      [SecurityCritical]
      public static CopyMoveResult Copy(string sourcePath, string destinationPath, CopyOptions copyOptions, int retry, int retryTimeout, CopyMoveProgressRoutine progressHandler, object userProgressData)
      {
         return CopyMoveCore(true, new CopyMoveArguments
         {
            Retry = retry,
            RetryTimeout = retryTimeout,
            CopyOptions = copyOptions,
            ProgressHandler = progressHandler,
            UserProgressData = userProgressData

         }, false, false, sourcePath, destinationPath, null);
      }


      /// <summary>[AlphaFS] 既存のファイルを新しいファイルにコピーします。同名ファイルの上書きが許可されます。  <see cref="CopyOptions"/> can be specified,
      /// およびコールバック関数を通じてアプリケーションに進行状況を通知する可能性があります。
      /// </summary>
      /// <remarks>
      ///   <para>非常に大きなファイル転送には<see cref="CopyOptions.NoBuffering"/>オプションが推奨されます。</para>
      ///   <para>元のファイルの属性はコピーされたファイルに保持されます。</para>
      ///   <para>可能な限り、このメソッドでは短いファイル名(<c>XXXXXX~1.XXX</c>など)の使用を避けてください。</para>
      ///   <para>2つのファイルの短いファイル名が同等の場合、このメソッドは失敗して例外を発生させるか、望ましくない動作になる可能性があります。</para>
      /// </remarks>
      /// <returns>コピー操作の詳細を含む<see cref="CopyMoveResult"/>クラスを返します。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="FileNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="sourcePath">コピーするファイル。</param>
      /// <param name="destinationPath">コピー先ファイルの名前。ディレクトリは指定できません。</param>
      /// <param name="copyOptions">ファイルのコピー方法を指定する<see cref="CopyOptions"/>。このパラメータは<c>null</c>にできます。</param>
      /// <param name="retry">コピー失敗時のリトライ回数。</param>
      /// <param name="retryTimeout">リトライ間の待機時間(秒)。</param>
      /// <param name="progressHandler">ファイルの別の部分がコピーされるたびに呼び出されるコールバック関数。このパラメータは<c>null</c>にできます。</param>
      /// <param name="userProgressData">コールバック関数に渡される引数。このパラメータは<c>null</c>にできます。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static CopyMoveResult Copy(string sourcePath, string destinationPath, CopyOptions copyOptions, int retry, int retryTimeout, CopyMoveProgressRoutine progressHandler, object userProgressData, PathFormat pathFormat)
      {
         return CopyMoveCore(true, new CopyMoveArguments
         {
            Retry = retry,
            RetryTimeout = retryTimeout,
            CopyOptions = copyOptions,
            ProgressHandler = progressHandler,
            UserProgressData = userProgressData,
            PathFormat = pathFormat

         }, false, false, sourcePath, destinationPath, null);
      }
   }
}
