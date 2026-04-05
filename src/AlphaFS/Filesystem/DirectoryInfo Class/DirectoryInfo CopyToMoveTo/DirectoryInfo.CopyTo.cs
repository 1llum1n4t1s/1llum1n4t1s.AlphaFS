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
      // .NET: Directory クラスには Copy() メソッドが含まれていないため、.NET File.Copy() メソッドを模倣します。


      #region Obsolete

      /// <summary>[AlphaFS] <see cref="DirectoryInfo"/> インスタンスとその内容を新しいパスにコピーします。</summary>
      /// <returns>新しい <see cref="DirectoryInfo"/> イ���スタンスを返します。</returns>
      /// <remarks>
      ///   <para>非常に大きなファイル転送には、<see cref="CopyOptions.NoBuffering"/> オプションが推奨されます。</para>
      ///   <para>このメソッドを使用して、既存のディレクトリの上書きを許可または防止します。</para>
      ///   <para>可能な限り、このメソッドでは短いファイル名（<c>XXXXXX~1.XXX</c> など）の使用を避けてください。</para>
      ///   <para>2 つのディレクトリに同等の短いファイル名がある場合、このメソッドが失敗して例外がスローされるか、望ましくない動作が発生する可能性があります。</para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="destinationPath">コピー先のディレクトリパス。</param>
      /// <param name="preserveDates">元のタイムスタンプを保持する場合は <c>true</c>、それ以外の場合は <c>false</c>。</param>
      [Obsolete("Use other overload and add CopyOptions.CopyTimestamp enum flag.")]
      [SecurityCritical]
      public DirectoryInfo CopyTo(string destinationPath, bool preserveDates)
      {

         CopyToMoveToCore(destinationPath, preserveDates, CopyOptions.FailIfExists, null, null, null, null, out var destinationPathLp, PathFormat.RelativePath);

         UpdateSourcePath(destinationPath, destinationPathLp);

         return new DirectoryInfo(Transaction, destinationPathLp, PathFormat.LongFullPath);
      }


      /// <summary>[AlphaFS] Copies an existing directory to a new directory, allowing the overwriting of an existing directory, <see cref="CopyOptions"/> can be specified.</summary>
      /// <returns>Returns a new <see cref="DirectoryInfo"/> instance.</returns>
      /// <remarks>
      ///   <para>Option <see cref="CopyOptions.NoBuffering"/> is recommended for very large file transfers.</para>
      ///   <para>Use this method to allow or prevent overwriting of an existing directory.</para>
      ///   <para>Whenever possible, avoid using short file names (such as <c>XXXXXX~1.XXX</c>) with this method.</para>
      ///   <para>If two directories have equivalent short file names then this method may fail and raise an exception and/or result in undesirable behavior.</para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="destinationPath">コピー先のディレクトリパス。</param>
      /// <param name="preserveDates">元のタイムスタンプを保持する場合は <c>true</c>、それ以外の場合は <c>false</c>。</param>
      /// <param name="pathFormat">パスパラメーターの形式を示します。</param>
      [Obsolete("Use other overload and add CopyOptions.CopyTimestamp enum flag.")]
      [SecurityCritical]
      public DirectoryInfo CopyTo(string destinationPath, bool preserveDates, PathFormat pathFormat)
      {

         CopyToMoveToCore(destinationPath, preserveDates, CopyOptions.FailIfExists, null, null, null, null, out var destinationPathLp, pathFormat);

         UpdateSourcePath(destinationPath, destinationPathLp);

         return new DirectoryInfo(Transaction, destinationPathLp, PathFormat.LongFullPath);
      }


      /// <summary>[AlphaFS] 既存のディレクトリを新しいディレクトリにコピーします。既存のディレクトリの上書きを許可し、<see cref="CopyOptions"/> を指定できます。</summary>
      /// <returns>
      ///   <para><paramref name="copyOptions"/> が <see cref="CopyOptions.FailIfExists"/> でない場合、新しいディレクトリまたは既存のディレクトリの上書きを返します。</para>
      ///   <para>ディレクトリが存在し、<paramref name="copyOptions"/> に <see cref="CopyOptions.FailIfExists"/> が含まれている場合、<see cref="IOException"/> がスローされます。</para>
      /// </returns>
      /// <remarks>
      ///   <para>非常に大きなファイル転送には、<see cref="CopyOptions.NoBuffering"/> オプションが推奨されます。</para>
      ///   <para>このメソッドを使用して、既存のディレクトリの上書きを許可または防止します。</para>
      ///   <para>可能な限り、このメソッドでは短いファイル名（<c>XXXXXX~1.XXX</c> など）の使用を避けてください。</para>
      ///   <para>2 つのディレクトリに同等の短いファイル名がある場合、このメソッドが失敗して例外がスローされるか、望ましくない動作が発生する可能性があります。</para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="destinationPath">コピー先のディレクトリパス。</param>
      /// <param name="copyOptions">ディレクトリのコピー方法を指定する <see cref="CopyOptions"/>。このパラメーターは <c>null</c> にできます。</param>
      /// <param name="preserveDates">元のタイムスタンプを保持する場合は <c>true</c>、それ以外の場合は <c>false</c>。</param>
      [Obsolete("Use other overload and add CopyOptions.CopyTimestamp enum flag.")]
      [SecurityCritical]
      public DirectoryInfo CopyTo(string destinationPath, CopyOptions copyOptions, bool preserveDates)
      {

         CopyToMoveToCore(destinationPath, preserveDates, copyOptions, null, null, null, null, out var destinationPathLp, PathFormat.RelativePath);

         UpdateSourcePath(destinationPath, destinationPathLp);

         return new DirectoryInfo(Transaction, destinationPathLp, PathFormat.LongFullPath);
      }


      /// <summary>[AlphaFS] 既存のディレクトリを新しいディレクトリにコピーします。既存のディレクトリの上書きを許可し、<see cref="CopyOptions"/> を指定できます。</summary>
      /// <returns>
      ///   <para><paramref name="copyOptions"/> が <see cref="CopyOptions.FailIfExists"/> でない場合、新しいディレクトリまたは既存のディレクトリの上書きを返します。</para>
      ///   <para>ディレクトリが存在し、<paramref name="copyOptions"/> に <see cref="CopyOptions.FailIfExists"/> が含まれている場合、<see cref="IOException"/> がスローされます。</para>
      /// </returns>
      /// <remarks>
      ///   <para>非常に大きなファイル転送には、<see cref="CopyOptions.NoBuffering"/> オプションが推奨されます。</para>
      ///   <para>このメソッドを使用して、既存のディレクトリの上書きを許可または防止します。</para>
      ///   <para>可能な限り、このメソッドでは短いファイル名（<c>XXXXXX~1.XXX</c> など）の使用を避けてください。</para>
      ///   <para>2 つのディレクトリに同等の短いファイル名がある場合、このメソッドが失敗して例外がスローされるか、望ましくない動作が発生する可能性があります。</para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="destinationPath">コピー先のディレクトリパス。</param>
      /// <param name="copyOptions">ディレクトリのコピー方法を指定する <see cref="CopyOptions"/>。このパラメーターは <c>null</c> にできます。</param>
      /// <param name="preserveDates">元のタイムスタンプを保持する場合は <c>true</c>、それ以外の場合は <c>false</c>。</param>
      /// <param name="pathFormat">パスパラメーターの形式を示します。</param>
      [Obsolete("Use other overload and add CopyOptions.CopyTimestamp enum flag.")]
      [SecurityCritical]
      public DirectoryInfo CopyTo(string destinationPath, CopyOptions copyOptions, bool preserveDates, PathFormat pathFormat)
      {

         CopyToMoveToCore(destinationPath, preserveDates, copyOptions, null, null, null, null, out var destinationPathLp, pathFormat);

         UpdateSourcePath(destinationPath, destinationPathLp);

         return new DirectoryInfo(Transaction, destinationPathLp, PathFormat.LongFullPath);
      }


      /// <summary>[AlphaFS] 既存のディレクトリを新しいディレクトリにコピーします。既存のディレクトリの上書きを許可し、<see cref="CopyOptions"/> を指定でき、
      /// コールバック関数を通じてアプリケーションに進行状況を通知できます。</summary>
      /// <returns>コピー操作の詳細を含む <see cref="CopyMoveResult"/> クラス。</returns>
      /// <remarks>
      ///   <para>非常に大きなファイル転送には、<see cref="CopyOptions.NoBuffering"/> オプションが推奨されます。</para>
      ///   <para>このメソッドを使用して、既存のディレクトリの上書きを許可または防止します。</para>
      ///   <para>可能な限り、このメソッドでは短いファイル名（<c>XXXXXX~1.XXX</c> など）の使用を避けてください。</para>
      ///   <para>2 つのディレクトリに同等の短いファイル名がある場合、このメソッドが失敗して例外がスローされるか、望ましくない動作が発生する可能性があります。</para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="destinationPath">コピー先のディレクトリパス。</param>
      /// <param name="copyOptions">ディレクトリのコピー方法を指定する <see cref="CopyOptions"/>。このパラメーターは <c>null</c> にできます。</param>
      /// <param name="preserveDates">元のタイムスタンプを保持する場合は <c>true</c>、それ以外の場合は <c>false</c>。</param>
      /// <param name="progressHandler">ディレクトリの別の部分がコピーされるたびに呼び出されるコールバック関数。このパラメーターは <c>null</c> にできます。</param>
      /// <param name="userProgressData">コールバック関数に渡される引数。このパラメーターは <c>null</c> にできます。</param>
      [Obsolete("Use other overload and add CopyOptions.CopyTimestamp enum flag.")]
      [SecurityCritical]
      public CopyMoveResult CopyTo(string destinationPath, CopyOptions copyOptions, bool preserveDates, CopyMoveProgressRoutine progressHandler, object userProgressData)
      {

         var cmr = CopyToMoveToCore(destinationPath, preserveDates, copyOptions, null, null, progressHandler, userProgressData, out var destinationPathLp, PathFormat.RelativePath);

         UpdateSourcePath(destinationPath, destinationPathLp);

         return cmr;
      }


      /// <summary>[AlphaFS] 既存のディレクトリを新しいディレクトリにコピーします。既存のディレクトリの上書きを許可し、<see cref="CopyOptions"/> を指定でき、
      /// コールバック関数を通じてアプリケーションに進行状況を通知できます。</summary>
      /// <returns>コピー操作の詳細を含む <see cref="CopyMoveResult"/> クラス。</returns>
      /// <remarks>
      ///   <para>非常に大きなファイル転送には、<see cref="CopyOptions.NoBuffering"/> オプションが推奨されます。</para>
      ///   <para>このメソッドを使用して、既存のディレクトリの上書きを許可または防止します。</para>
      ///   <para>可能な限り、このメソッドでは短いファイル名（<c>XXXXXX~1.XXX</c> など）の使用を避けてください。</para>
      ///   <para>2 つのディレクトリに同等の短いファイル名がある場合、このメソッドが失敗して例外がスローされるか、望ましくない動作が発生する可能性があります。</para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="destinationPath">コピー先のディレクトリパス。</param>
      /// <param name="copyOptions">ディレクトリのコピー方法を指定する <see cref="CopyOptions"/>。このパラメーターは <c>null</c> にできます。</param>
      /// <param name="preserveDates">元のタイムスタンプを保持する場合は <c>true</c>、それ以外の場合は <c>false</c>。</param>
      /// <param name="progressHandler">ディレクトリの別の部分がコピーされるたびに呼び出されるコールバック関数。このパラメーターは <c>null</c> にできます。</param>
      /// <param name="userProgressData">コールバック関数に渡される引数。このパラメーターは <c>null</c> にできます。</param>
      /// <param name="pathFormat">パスパラメーターの形式を示します。</param>
      [Obsolete("Use other overload and add CopyOptions.CopyTimestamp enum flag.")]
      [SecurityCritical]
      public CopyMoveResult CopyTo(string destinationPath, CopyOptions copyOptions, bool preserveDates, CopyMoveProgressRoutine progressHandler, object userProgressData, PathFormat pathFormat)
      {

         var cmr = CopyToMoveToCore(destinationPath, preserveDates, copyOptions, null, null, progressHandler, userProgressData, out var destinationPathLp, pathFormat);

         UpdateSourcePath(destinationPath, destinationPathLp);

         return cmr;
      }
      
      #endregion // Obsolete


      /// <summary>[AlphaFS] <see cref="DirectoryInfo"/> インスタンスとその内容を新しいパスにコピーします。</summary>
      /// <returns>ディレクトリが完全にコピーされた場合の新しい <see cref="DirectoryInfo"/> インスタンス。</returns>
      /// <remarks>
      ///   <para>デフォルトで既存のディレクトリの上書きを防止するには、このメソッドを使用します。</para>
      ///   <para>可能な限り、このメソッドでは短いファイル名（<c>XXXXXX~1.XXX</c> など）の使用を避けてください。</para>
      ///   <para>2 つのディレクトリに同等の短いファイル名がある場合、このメソッドが失敗して例外がスローされるか、望ましくない動作が発生する可能性があります。</para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="destinationPath">コピー先のディレクトリパス。</param>
      [SecurityCritical]
      public DirectoryInfo CopyTo(string destinationPath)
      {

         CopyToMoveToCore(destinationPath, false, CopyOptions.FailIfExists, null, null, null, null, out var destinationPathLp, PathFormat.RelativePath);

         UpdateSourcePath(destinationPath, destinationPathLp);

         return new DirectoryInfo(Transaction, destinationPathLp, PathFormat.LongFullPath);
      }


      /// <summary>[AlphaFS] <see cref="DirectoryInfo"/> インスタンスとその内容を新しいパスにコピーします。</summary>
      /// <returns>ディレクトリが完全にコピーされた場合の新しい <see cref="DirectoryInfo"/> インスタンス。</returns>
      /// <remarks>
      ///   <para>デフォルトで既存のディレクトリの上書きを防止するには、このメソッドを使用します。</para>
      ///   <para>可能な限り、このメソッドでは短いファイル名（<c>XXXXXX~1.XXX</c> など）の使用を避けてください。</para>
      ///   <para>2 つのディレクトリに同等の短いファイル名がある場合、このメソッドが失敗して例外がスローされるか、望ましくない動作が発生する可能性があります。</para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="destinationPath">コピー先のディレクトリパス。</param>
      /// <param name="pathFormat">パスパラメーターの形式を示します。</param>
      [SecurityCritical]
      public DirectoryInfo CopyTo(string destinationPath, PathFormat pathFormat)
      {

         CopyToMoveToCore(destinationPath, false, CopyOptions.FailIfExists, null, null, null, null, out var destinationPathLp, pathFormat);

         UpdateSourcePath(destinationPath, destinationPathLp);

         return new DirectoryInfo(Transaction, destinationPathLp, PathFormat.LongFullPath);
      }
      

      /// <summary>[AlphaFS] 既存のディレクトリを新しいディレクトリにコピーします。既存のディレクトリの上書きを許可し、<see cref="CopyOptions"/> を指定できます。</summary>
      /// <returns>
      ///   <para><paramref name="copyOptions"/> が <see cref="CopyOptions.FailIfExists"/> でない場合、新しいディレクトリまたは既存のディレクトリの上書きを返します。</para>
      ///   <para>ディレクトリが存在し、<paramref name="copyOptions"/> に <see cref="CopyOptions.FailIfExists"/> が含まれている場合、<see cref="IOException"/> がスローされます。</para>
      /// </returns>
      /// <remarks>
      ///   <para>非常に大きなファイル転送には、<see cref="CopyOptions.NoBuffering"/> オプションが推奨されます。</para>
      ///   <para>このメソッドを使用して、既存のディレクトリの上書きを許可または防止します。</para>
      ///   <para>可能な限り、このメソッドでは短いファイル名（<c>XXXXXX~1.XXX</c> など）の使用を避けてください。</para>
      ///   <para>2 つのディレクトリに同等の短いファイル名がある場合、このメソッドが失敗して例外がスローされるか、望ましくない動作が発生する可能性があります。</para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="destinationPath">コピー先のディレクトリパス。</param>
      /// <param name="copyOptions">ディレクトリのコピー方法を指定する <see cref="CopyOptions"/>。このパラメーターは <c>null</c> にできます。</param>
      [SecurityCritical]
      public DirectoryInfo CopyTo(string destinationPath, CopyOptions copyOptions)
      {

         CopyToMoveToCore(destinationPath, false, copyOptions, null, null, null, null, out var destinationPathLp, PathFormat.RelativePath);

         UpdateSourcePath(destinationPath, destinationPathLp);

         return new DirectoryInfo(Transaction, destinationPathLp, PathFormat.LongFullPath);
      }


      /// <summary>[AlphaFS] 既存のディレクトリを新しいディレクトリにコピーします。既存のディレクトリの上書きを許可し、<see cref="CopyOptions"/> を指定できます。</summary>
      /// <returns>
      ///   <para><paramref name="copyOptions"/> が <see cref="CopyOptions.FailIfExists"/> でない場合、新しいディレクトリまたは既存のディレクトリの上書きを返します。</para>
      ///   <para>ディレクトリが存在し、<paramref name="copyOptions"/> に <see cref="CopyOptions.FailIfExists"/> が含まれている場合、<see cref="IOException"/> がスローされます。</para>
      /// </returns>
      /// <remarks>
      ///   <para>非常に大きなファイル転送には、<see cref="CopyOptions.NoBuffering"/> オプションが推奨されます。</para>
      ///   <para>このメソッドを使用して、既存のディレクトリの上書きを許可または防止します。</para>
      ///   <para>可能な限り、このメソッドでは短いファイル名（<c>XXXXXX~1.XXX</c> など）の使用を避けてください。</para>
      ///   <para>2 つのディレクトリに同等の短いファイル名がある場合、このメソッドが失敗して例外がスローされるか、望ましくない動作が発生する可能性があります。</para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="destinationPath">コピー先のディレクトリパス。</param>
      /// <param name="copyOptions">ディレクトリのコピー方法を指定する <see cref="CopyOptions"/>。このパラメーターは <c>null</c> にできます。</param>
      /// <param name="pathFormat">パスパラメーターの形式を示します。</param>
      [SecurityCritical]
      public DirectoryInfo CopyTo(string destinationPath, CopyOptions copyOptions, PathFormat pathFormat)
      {

         CopyToMoveToCore(destinationPath, false, copyOptions, null, null, null, null, out var destinationPathLp, pathFormat);

         UpdateSourcePath(destinationPath, destinationPathLp);

         return new DirectoryInfo(Transaction, destinationPathLp, PathFormat.LongFullPath);
      }
      

      /// <summary>[AlphaFS] 既存のディレクトリを新しいディレクトリにコピーします。既存のディレクトリの上書きを許可し、<see cref="CopyOptions"/> を指定でき、
      /// コールバック関数を通じてアプリケーションに進行状況を通知できます。
      /// </summary>
      /// <returns>コピー操作の詳細を含む <see cref="CopyMoveResult"/> クラス。</returns>
      /// <remarks>
      ///   <para>非常に大きなファイル転送には、<see cref="CopyOptions.NoBuffering"/> オプションが推奨されます。</para>
      ///   <para>このメソッドを使用して、既存のディレクトリの上書きを許可または防止します。</para>
      ///   <para>可能な限り、このメソッドでは短いファイル名（<c>XXXXXX~1.XXX</c> など）の使用を避けてください。</para>
      ///   <para>2 つのディレクトリに同等の短いファイル名がある場合、このメソッドが失敗して例外がスローされるか、望ましくない動作が発生する可能性があります。</para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="destinationPath">コピー先のディレクトリパス。</param>
      /// <param name="copyOptions">ディレクトリのコピー方法を指定する <see cref="CopyOptions"/>。このパラメーターは <c>null</c> にできます。</param>
      /// <param name="progressHandler">ディレクトリの別の部分がコピーされるたびに呼び出されるコールバック関数。このパラメーターは <c>null</c> にできます。</param>
      /// <param name="userProgressData">コールバック関数に渡される引数。このパラメーターは <c>null</c> にできます。</param>
      [SecurityCritical]
      public CopyMoveResult CopyTo(string destinationPath, CopyOptions copyOptions, CopyMoveProgressRoutine progressHandler, object userProgressData)
      {

         var cmr = CopyToMoveToCore(destinationPath, false, copyOptions, null, null, progressHandler, userProgressData, out var destinationPathLp, PathFormat.RelativePath);

         UpdateSourcePath(destinationPath, destinationPathLp);

         return cmr;
      }


      /// <summary>[AlphaFS] Copies an existing directory to a new directory, allowing the overwriting of an existing directory, <see cref="CopyOptions"/> can be specified
      /// and the possibility of notifying the application of its progress through a callback function.</summary>
      /// <returns>A <see cref="CopyMoveResult"/> class with details of the Copy action.</returns>
      /// <remarks>
      ///   <para>Option <see cref="CopyOptions.NoBuffering"/> is recommended for very large file transfers.</para>
      ///   <para>Use this method to allow or prevent overwriting of an existing directory.</para>
      ///   <para>Whenever possible, avoid using short file names (such as <c>XXXXXX~1.XXX</c>) with this method.</para>
      ///   <para>If two directories have equivalent short file names then this method may fail and raise an exception and/or result in undesirable behavior.</para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="destinationPath">コピー先のディレクトリパス。</param>
      /// <param name="copyOptions">ディレクトリのコピー方法を指定する <see cref="CopyOptions"/>。このパラメーターは <c>null</c> にできます。</param>
      /// <param name="progressHandler">ディレクトリの別の部分がコピーされるたびに呼び出されるコールバック関数。このパラメーターは <c>null</c> にできます。</param>
      /// <param name="userProgressData">コールバック関数に渡される引数。このパラメーターは <c>null</c> にできます。</param>
      /// <param name="pathFormat">パスパラメーターの形式を示します。</param>
      [SecurityCritical]
      public CopyMoveResult CopyTo(string destinationPath, CopyOptions copyOptions, CopyMoveProgressRoutine progressHandler, object userProgressData, PathFormat pathFormat)
      {

         var cmr = CopyToMoveToCore(destinationPath, false, copyOptions, null, null, progressHandler, userProgressData, out var destinationPathLp, pathFormat);

         UpdateSourcePath(destinationPath, destinationPathLp);

         return cmr;
      }


      /// <summary>[AlphaFS] 既存のディレクトリを新しいディレクトリにコピーします。既存のディレクトリの上書きを許可し、<see cref="CopyOptions"/> を指定できます。</summary>
      /// <returns>
      ///   <para><paramref name="copyOptions"/> が <see cref="CopyOptions.FailIfExists"/> でない場合、新しいディレクトリまたは既存のディレクトリの上書きを返します。</para>
      ///   <para>ディレクトリが存在し、<paramref name="copyOptions"/> に <see cref="CopyOptions.FailIfExists"/> が含まれている場合、<see cref="IOException"/> がスローされます。</para>
      /// </returns>
      /// <remarks>
      ///   <para>非常に大きなファイル転送には、<see cref="CopyOptions.NoBuffering"/> オプションが推奨されます。</para>
      ///   <para>このメソッドを使用して、既存のディレクトリの上書きを許可または防止します。</para>
      ///   <para>可能な限り、このメソッドでは短いファイル名（<c>XXXXXX~1.XXX</c> など）の使用を避けてください。</para>
      ///   <para>2 つのディレクトリに同等の短いファイル名がある場合、このメソッドが失敗して例外がスローされるか、望ましくない動作が発生する可能性があります。</para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="destinationPath">コピー先のディレクトリパス。</param>
      /// <param name="copyOptions">ディレクトリのコピー方法を指定する <see cref="CopyOptions"/>。このパラメーターは <c>null</c> にできます。</param>
      /// <param name="filters">処理で使用するカスタムフィルターの仕様。</param>
      [SecurityCritical]
      public DirectoryInfo CopyTo(string destinationPath, CopyOptions copyOptions, DirectoryEnumerationFilters filters)
      {

         CopyToMoveToCore(destinationPath, false, copyOptions, null, filters, null, null, out var destinationPathLp, PathFormat.RelativePath);

         UpdateSourcePath(destinationPath, destinationPathLp);

         return new DirectoryInfo(Transaction, destinationPathLp, PathFormat.LongFullPath);
      }


      /// <summary>[AlphaFS] 既存のディレクトリを新しいディレクトリにコピーします。既存のディレクトリの上書きを許可し、<see cref="CopyOptions"/> を指定できます。</summary>
      /// <returns>
      ///   <para><paramref name="copyOptions"/> が <see cref="CopyOptions.FailIfExists"/> でない場合、新しいディレクトリまたは既存のディレクトリの上書きを返します。</para>
      ///   <para>ディレクトリが存在し、<paramref name="copyOptions"/> に <see cref="CopyOptions.FailIfExists"/> が含まれている場合、<see cref="IOException"/> がスローされます。</para>
      /// </returns>
      /// <remarks>
      ///   <para>非常に大きなファイル転送には、<see cref="CopyOptions.NoBuffering"/> オプションが推奨されます。</para>
      ///   <para>このメソッドを使用して、既存のディレクトリの上書きを許可または防止します。</para>
      ///   <para>可能な限り、このメソッドでは短いファイル名（<c>XXXXXX~1.XXX</c> など）の使用を避けてください。</para>
      ///   <para>2 つのディレクトリに同等の短いファイル名がある場合、このメソッドが失敗して例外がスローされるか、望ましくない動作が発生する可能性があります。</para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="destinationPath">コピー先のディレクトリパス。</param>
      /// <param name="copyOptions">ディレクトリのコピー方法を指定する <see cref="CopyOptions"/>。このパラメーターは <c>null</c> にできます。</param>
      /// <param name="filters">処理で使用するカスタムフィルターの仕様。</param>
      /// <param name="pathFormat">パスパラメーターの形式を示します。</param>
      [SecurityCritical]
      public DirectoryInfo CopyTo(string destinationPath, CopyOptions copyOptions, DirectoryEnumerationFilters filters, PathFormat pathFormat)
      {

         CopyToMoveToCore(destinationPath, false, copyOptions, null, filters, null, null, out var destinationPathLp, pathFormat);

         UpdateSourcePath(destinationPath, destinationPathLp);

         return new DirectoryInfo(Transaction, destinationPathLp, PathFormat.LongFullPath);
      }


      /// <summary>[AlphaFS] Copies an existing directory to a new directory, allowing the overwriting of an existing directory, <see cref="CopyOptions"/> can be specified
      /// and the possibility of notifying the application of its progress through a callback function.</summary>
      /// <returns>
      ///   <para>Returns a new directory, or an overwrite of an existing directory if <paramref name="copyOptions"/> is not <see cref="CopyOptions.FailIfExists"/>.</para>
      ///   <para>If the directory exists and <paramref name="copyOptions"/> contains <see cref="CopyOptions.FailIfExists"/>, an <see cref="IOException"/> is thrown.</para>
      /// </returns>
      /// <remarks>
      ///   <para>Option <see cref="CopyOptions.NoBuffering"/> is recommended for very large file transfers.</para>
      ///   <para>Use this method to allow or prevent overwriting of an existing directory.</para>
      ///   <para>Whenever possible, avoid using short file names (such as <c>XXXXXX~1.XXX</c>) with this method.</para>
      ///   <para>If two directories have equivalent short file names then this method may fail and raise an exception and/or result in undesirable behavior.</para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="destinationPath">コピー先のディレクトリパス。</param>
      /// <param name="copyOptions">ディレクトリのコピー方法を指定する <see cref="CopyOptions"/>。このパラメーターは <c>null</c> にできます。</param>
      /// <param name="filters">処理で使用するカスタムフィルターの仕様。</param>
      /// <param name="progressHandler">ディレクトリの別の部分がコピーされるたびに呼び出されるコールバック関数。このパラメーターは <c>null</c> にできます。</param>
      /// <param name="userProgressData">コールバック関数に渡される引数。このパラメーターは <c>null</c> にできます。</param>
      [SecurityCritical]
      public DirectoryInfo CopyTo(string destinationPath, CopyOptions copyOptions, DirectoryEnumerationFilters filters, CopyMoveProgressRoutine progressHandler, object userProgressData)
      {

         CopyToMoveToCore(destinationPath, false, copyOptions, null, filters, progressHandler, userProgressData, out var destinationPathLp, PathFormat.RelativePath);

         UpdateSourcePath(destinationPath, destinationPathLp);

         return new DirectoryInfo(Transaction, destinationPathLp, PathFormat.LongFullPath);
      }


      /// <summary>[AlphaFS] Copies an existing directory to a new directory, allowing the overwriting of an existing directory, <see cref="CopyOptions"/> can be specified
      /// and the possibility of notifying the application of its progress through a callback function.</summary>
      /// <returns>
      ///   <para>Returns a new directory, or an overwrite of an existing directory if <paramref name="copyOptions"/> is not <see cref="CopyOptions.FailIfExists"/>.</para>
      ///   <para>If the directory exists and <paramref name="copyOptions"/> contains <see cref="CopyOptions.FailIfExists"/>, an <see cref="IOException"/> is thrown.</para>
      /// </returns>
      /// <remarks>
      ///   <para>Option <see cref="CopyOptions.NoBuffering"/> is recommended for very large file transfers.</para>
      ///   <para>Use this method to allow or prevent overwriting of an existing directory.</para>
      ///   <para>Whenever possible, avoid using short file names (such as <c>XXXXXX~1.XXX</c>) with this method.</para>
      ///   <para>If two directories have equivalent short file names then this method may fail and raise an exception and/or result in undesirable behavior.</para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="destinationPath">コピー先のディレクトリパス。</param>
      /// <param name="copyOptions">ディレクトリのコピー方法を指定する <see cref="CopyOptions"/>。このパラメーターは <c>null</c> にできます。</param>
      /// <param name="filters">処理で使用するカスタムフィルターの仕様。</param>
      /// <param name="progressHandler">ディレクトリの別の部分がコピーされるたびに呼び出されるコールバック関数。このパラメーターは <c>null</c> にできます。</param>
      /// <param name="userProgressData">コールバック関数に渡される引数。このパラメーターは <c>null</c> にできます。</param>
      /// <param name="pathFormat">パスパラメーターの形式を示します。</param>
      [SecurityCritical]
      public DirectoryInfo CopyTo(string destinationPath, CopyOptions copyOptions, DirectoryEnumerationFilters filters, CopyMoveProgressRoutine progressHandler, object userProgressData, PathFormat pathFormat)
      {

         CopyToMoveToCore(destinationPath, false, copyOptions, null, filters, progressHandler, userProgressData, out var destinationPathLp, pathFormat);

         UpdateSourcePath(destinationPath, destinationPathLp);

         return new DirectoryInfo(Transaction, destinationPathLp, PathFormat.LongFullPath);
      }
   }
}
