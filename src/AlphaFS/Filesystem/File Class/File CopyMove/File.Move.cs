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
      #region .NET

      /// <summary>指定されたファイルを新しい場所に移動します。新しいファイル名を指定するオプションを提供します。
      /// <remarks>
      ///   <para>このメソッドはディスクボリュームをまたいで動作します。</para>
      ///   <para>同名のファイルをそのディレクトリに移動してファイルを置換しようとすると、<see cref="IOException"/>が発生します。</para>
      ///   <para>Moveメソッドを使用して既存のファイルを上書きすることはできません。</para>
      ///   <para>可能な限り、このメソッドでは短いファイル名(<c>XXXXXX~1.XXX</c>など)の使用を避けてください。</para>
      ///   <para>2つのファイルの短いファイル名が同等の場合、このメソッドは失敗して例外を発生させるか、望ましくない動作になる可能性があります。</para>
      /// </remarks>
      /// </summary>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="FileNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="sourcePath">移動するファイルの名前。</param>
      /// <param name="destinationPath">ファイルの新しいパス。</param>
      [SecurityCritical]
      public static void Move(string sourcePath, string destinationPath)
      {
         CopyMoveCore(false, new CopyMoveArguments
         {
            MoveOptions = MoveOptions.CopyAllowed

         }, false, false, sourcePath, destinationPath, null);
      }

      #endregion // .NET


      /// <summary>[AlphaFS] 指定されたファイルを新しい場所に移動します。新しいファイル名を指定するオプションを提供します。</summary>
      /// <returns>移動操作のステータスを含む<see cref="CopyMoveResult"/>クラス。</returns>
      /// <remarks>
      ///   <para>このメソッドはディスクボリュームをまたいで動作します。</para>
      ///   <para>同名のファイルをそのディレクトリに移動してファイルを置換しようとすると、<see cref="IOException"/>が発生します。</para>
      ///   <para>Moveメソッドを使用して既存のファイルを上書きすることはできません。</para>
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
      /// <param name="sourcePath">移動するファイルの名前。</param>
      /// <param name="destinationPath">ファイルの新しいパス。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static CopyMoveResult Move(string sourcePath, string destinationPath, PathFormat pathFormat)
      {
         return CopyMoveCore(false, new CopyMoveArguments
         {
            MoveOptions = MoveOptions.CopyAllowed,
            PathFormat = pathFormat

         }, false, false, sourcePath, destinationPath, null);
      }


      /// <summary>[AlphaFS] 指定されたファイルを新しい場所に移動します。新しいファイル名を指定するオプションを提供します。
      /// <returns>移動操作のステータスを含む<see cref="CopyMoveResult"/>クラス。</returns>
      /// <remarks>
      ///   <para>このメソッドはディスクボリュームをまたいで動作します。</para>
      ///   <para>同名のファイルをそのディレクトリに移動してファイルを置換しようとすると、<see cref="IOException"/>が発生します。</para>
      ///   <para>Moveメソッドを使用して既存のファイルを上書きすることはできません。</para>
      ///   <para>可能な限り、このメソッドでは短いファイル名(<c>XXXXXX~1.XXX</c>など)の使用を避けてください。</para>
      ///   <para>2つのファイルの短いファイル名が同等の場合、このメソッドは失敗して例外を発生させるか、望ましくない動作になる可能性があります。</para>
      /// </remarks>
      /// </summary>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="FileNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="sourcePath">移動するファイルの名前。</param>
      /// <param name="destinationPath">ファイルの新しいパス。</param>
      /// <param name="retry">コピー失敗時のリトライ回数。</param>
      /// <param name="retryTimeout">リトライ間の待機時間(秒)。</param>
      [SecurityCritical]
      public static CopyMoveResult Move(string sourcePath, string destinationPath, int retry, int retryTimeout)
      {
         return CopyMoveCore(false, new CopyMoveArguments
         {
            Retry = retry,
            RetryTimeout = retryTimeout,
            MoveOptions = MoveOptions.CopyAllowed

         }, false, false, sourcePath, destinationPath, null);
      }


      /// <summary>[AlphaFS] 指定されたファイルを新しい場所に移動します。新しいファイル名を指定するオプションを提供します。</summary>
      /// <returns>移動操作のステータスを含む<see cref="CopyMoveResult"/>クラス。</returns>
      /// <remarks>
      ///   <para>このメソッドはディスクボリュームをまたいで動作します。</para>
      ///   <para>同名のファイルをそのディレクトリに移動してファイルを置換しようとすると、<see cref="IOException"/>が発生します。</para>
      ///   <para>Moveメソッドを使用して既存のファイルを上書きすることはできません。</para>
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
      /// <param name="sourcePath">移動するファイルの名前。</param>
      /// <param name="destinationPath">ファイルの新しいパス。</param>
      /// <param name="retry">コピー失敗時のリトライ回数。</param>
      /// <param name="retryTimeout">リトライ間の待機時間(秒)。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static CopyMoveResult Move(string sourcePath, string destinationPath, int retry, int retryTimeout, PathFormat pathFormat)
      {
         return CopyMoveCore(false, new CopyMoveArguments
         {
            Retry = retry,
            RetryTimeout = retryTimeout,
            MoveOptions = MoveOptions.CopyAllowed,
            PathFormat = pathFormat

         }, false, false, sourcePath, destinationPath, null);
      }


      /// <summary>[AlphaFS] 指定されたファイルを新しい場所に移動します。新しいファイル名を指定するオプションを提供します。</summary>
      /// <returns>移動操作のステータスを含む<see cref="CopyMoveResult"/>クラス。</returns>
      /// <remarks>
      ///   <para>このメソッドはディスクボリュームをまたいで動作します。</para>
      ///   <para>同名のファイルをそのディレクトリに移動してファイルを置換しようとすると、<see cref="IOException"/>が発生します。</para>
      ///   <para>Moveメソッドを使用して既存のファイルを上書きすることはできません。</para>
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
      /// <param name="sourcePath">移動するファイルの名前。</param>
      /// <param name="destinationPath">ファイルの新しいパス。</param>
      /// <param name="progressHandler">ファイルの別の部分が移動されるたびに呼び出されるコールバック関数。このパラメータは<c>null</c>にできます。</param>
      /// <param name="userProgressData">コールバック関数に渡される引数。このパラメータは<c>null</c>にできます。</param>
      [SecurityCritical]
      public static CopyMoveResult Move(string sourcePath, string destinationPath, CopyMoveProgressRoutine progressHandler, object userProgressData)
      {
         return CopyMoveCore(false, new CopyMoveArguments
         {
            MoveOptions = MoveOptions.CopyAllowed,
            ProgressHandler = progressHandler,
            UserProgressData = userProgressData

         }, false, false, sourcePath, destinationPath, null);
      }


      /// <summary>[AlphaFS] 指定されたファイルを新しい場所に移動します。新しいファイル名を指定するオプションを提供します。</summary>
      /// <returns>移動操作のステータスを含む<see cref="CopyMoveResult"/>クラス。</returns>
      /// <returns>移動操作のステータスを含む<see cref="CopyMoveResult"/>クラス。</returns>
      /// <remarks>
      ///   <para>このメソッドはディスクボリュームをまたいで動作します。</para>
      ///   <para>同名のファイルをそのディレクトリに移動してファイルを置換しようとすると、<see cref="IOException"/>が発生します。</para>
      ///   <para>Moveメソッドを使用して既存のファイルを上書きすることはできません。</para>
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
      /// <param name="sourcePath">移動するファイルの名前。</param>
      /// <param name="destinationPath">ファイルの新しいパス。</param>
      /// <param name="progressHandler">ファイルの別の部分が移動されるたびに呼び出されるコールバック関数。このパラメータは<c>null</c>にできます。</param>
      /// <param name="userProgressData">コールバック関数に渡される引数。このパラメータは<c>null</c>にできます。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static CopyMoveResult Move(string sourcePath, string destinationPath, CopyMoveProgressRoutine progressHandler, object userProgressData, PathFormat pathFormat)
      {
         return CopyMoveCore(false, new CopyMoveArguments
         {
            MoveOptions = MoveOptions.CopyAllowed,
            ProgressHandler = progressHandler,
            UserProgressData = userProgressData,
            PathFormat = pathFormat

         }, false, false, sourcePath, destinationPath, null);
      }


      /// <summary>[AlphaFS] 指定されたファイルを新しい場所に移動します。新しいファイル名を指定するオプションを提供します。</summary>
      /// <returns>移動操作のステータスを含む<see cref="CopyMoveResult"/>クラス。</returns>
      /// <remarks>
      ///   <para>このメソッドはディスクボリュームをまたいで動作します。</para>
      ///   <para>同名のファイルをそのディレクトリに移動してファイルを置換しようとすると、<see cref="IOException"/>が発生します。</para>
      ///   <para>Moveメソッドを使用して既存のファイルを上書きすることはできません。</para>
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
      /// <param name="sourcePath">移動するファイルの名前。</param>
      /// <param name="destinationPath">ファイルの新しいパス。</param>
      /// <param name="retry">コピー失敗時のリトライ回数。</param>
      /// <param name="retryTimeout">リトライ間の待機時間(秒)。</param>
      /// <param name="progressHandler">ファイルの別の部分が移動されるたびに呼び出されるコールバック関数。このパラメータは<c>null</c>にできます。</param>
      /// <param name="userProgressData">コールバック関数に渡される引数。このパラメータは<c>null</c>にできます。</param>
      [SecurityCritical]
      public static CopyMoveResult Move(string sourcePath, string destinationPath, int retry, int retryTimeout, CopyMoveProgressRoutine progressHandler, object userProgressData)
      {
         return CopyMoveCore(false, new CopyMoveArguments
         {
            Retry = retry,
            RetryTimeout = retryTimeout,
            MoveOptions = MoveOptions.CopyAllowed,
            ProgressHandler = progressHandler,
            UserProgressData = userProgressData

         }, false, false, sourcePath, destinationPath, null);
      }


      /// <summary>[AlphaFS] 指定されたファイルを新しい場所に移動します。新しいファイル名を指定するオプションを提供します。</summary>
      /// <returns>移動操作のステータスを含む<see cref="CopyMoveResult"/>クラス。</returns>
      /// <returns>移動操作のステータスを含む<see cref="CopyMoveResult"/>クラス。</returns>
      /// <remarks>
      ///   <para>このメソッドはディスクボリュームをまたいで動作します。</para>
      ///   <para>同名のファイルをそのディレクトリに移動してファイルを置換しようとすると、<see cref="IOException"/>が発生します。</para>
      ///   <para>Moveメソッドを使用して既存のファイルを上書きすることはできません。</para>
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
      /// <param name="sourcePath">移動するファイルの名前。</param>
      /// <param name="destinationPath">ファイルの新しいパス。</param>
      /// <param name="retry">コピー失敗時のリトライ回数。</param>
      /// <param name="retryTimeout">リトライ間の待機時間(秒)。</param>
      /// <param name="progressHandler">ファイルの別の部分が移動されるたびに呼び出されるコールバック関数。このパラメータは<c>null</c>にできます。</param>
      /// <param name="userProgressData">コールバック関数に渡される引数。このパラメータは<c>null</c>にできます。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static CopyMoveResult Move(string sourcePath, string destinationPath, int retry, int retryTimeout, CopyMoveProgressRoutine progressHandler, object userProgressData, PathFormat pathFormat)
      {
         return CopyMoveCore(false, new CopyMoveArguments
         {
            Retry = retry,
            RetryTimeout = retryTimeout,
            MoveOptions = MoveOptions.CopyAllowed,
            ProgressHandler = progressHandler,
            UserProgressData = userProgressData,
            PathFormat = pathFormat

         }, false, false, sourcePath, destinationPath, null);
      }




      /// <summary>[AlphaFS] 指定されたファイルを新しい場所に移動します。新しいファイル名を指定するオプションを提供します。</summary>
      /// <returns>移動操作のステータスを含む<see cref="CopyMoveResult"/>クラス。</returns>
      /// <remarks>
      ///   <para>このメソッドはディスクボリュームをまたいで動作します。</para>
      ///   <para>同名のファイルをそのディレクトリに移動してファイルを置換しようとすると、<see cref="IOException"/>が発生します。</para>
      ///   <para>Moveメソッドを使用して既存のファイルを上書きすることはできません。</para>
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
      /// <param name="sourcePath">移動するファイルの名前。</param>
      /// <param name="destinationPath">ファイルの新しいパス。</param>
      /// <param name="moveOptions">ファイルの移動方法を指定する<see cref="MoveOptions"/>。このパラメータは<c>null</c>にできます。</param>
      [SecurityCritical]
      public static CopyMoveResult Move(string sourcePath, string destinationPath, MoveOptions moveOptions)
      {
         return CopyMoveCore(false, new CopyMoveArguments
         {
            MoveOptions = moveOptions

         }, false, false, sourcePath, destinationPath, null);
      }


      /// <summary>[AlphaFS] 指定されたファイルを新しい場所に移動します。新しいファイル名を指定するオプションを提供します。</summary>
      /// <returns>移動操作のステータスを含む<see cref="CopyMoveResult"/>クラス。</returns>
      /// <remarks>
      ///   <para>このメソッドはディスクボリュームをまたいで動作します。</para>
      ///   <para>同名のファイルをそのディレクトリに移動してファイルを置換しようとすると、<see cref="IOException"/>が発生します。</para>
      ///   <para>Moveメソッドを使用して既存のファイルを上書きすることはできません。</para>
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
      /// <param name="sourcePath">移動するファイルの名前。</param>
      /// <param name="destinationPath">ファイルの新しいパス。</param>
      /// <param name="moveOptions">ファイルの移動方法を指定する<see cref="MoveOptions"/>。このパラメータは<c>null</c>にできます。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static CopyMoveResult Move(string sourcePath, string destinationPath, MoveOptions moveOptions, PathFormat pathFormat)
      {
         return CopyMoveCore(false, new CopyMoveArguments
         {
            MoveOptions = moveOptions,
            PathFormat = pathFormat

         }, false, false, sourcePath, destinationPath, null);
      }


      /// <summary>[AlphaFS] 指定されたファイルを新しい場所に移動します。新しいファイル名を指定するオプションを提供します。</summary>
      /// <returns>移動操作のステータスを含む<see cref="CopyMoveResult"/>クラス。</returns>
      /// <remarks>
      ///   <para>このメソッドはディスクボリュームをまたいで動作します。</para>
      ///   <para>同名のファイルをそのディレクトリに移動してファイルを置換しようとすると、<see cref="IOException"/>が発生します。</para>
      ///   <para>Moveメソッドを使用して既存のファイルを上書きすることはできません。</para>
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
      /// <param name="sourcePath">移動するファイルの名前。</param>
      /// <param name="destinationPath">ファイルの新しいパス。</param>
      /// <param name="moveOptions">ファイルの移動方法を指定する<see cref="MoveOptions"/>。このパラメータは<c>null</c>にできます。</param>
      /// <param name="retry">コピー失敗時のリトライ回数。</param>
      /// <param name="retryTimeout">リトライ間の待機時間(秒)。</param>
      [SecurityCritical]
      public static CopyMoveResult Move(string sourcePath, string destinationPath, MoveOptions moveOptions, int retry, int retryTimeout)
      {
         return CopyMoveCore(false, new CopyMoveArguments
         {
            Retry = retry,
            RetryTimeout = retryTimeout,
            MoveOptions = moveOptions

         }, false, false, sourcePath, destinationPath, null);
      }


      /// <summary>[AlphaFS] 指定されたファイルを新しい場所に移動します。新しいファイル名を指定するオプションを提供します。</summary>
      /// <returns>移動操作のステータスを含む<see cref="CopyMoveResult"/>クラス。</returns>
      /// <remarks>
      ///   <para>このメソッドはディスクボリュームをまたいで動作します。</para>
      ///   <para>同名のファイルをそのディレクトリに移動してファイルを置換しようとすると、<see cref="IOException"/>が発生します。</para>
      ///   <para>Moveメソッドを使用して既存のファイルを上書きすることはできません。</para>
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
      /// <param name="sourcePath">移動するファイルの名前。</param>
      /// <param name="destinationPath">ファイルの新しいパス。</param>
      /// <param name="moveOptions">ファイルの移動方法を指定する<see cref="MoveOptions"/>。このパラメータは<c>null</c>にできます。</param>
      /// <param name="retry">コピー失敗時のリトライ回数。</param>
      /// <param name="retryTimeout">リトライ間の待機時間(秒)。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static CopyMoveResult Move(string sourcePath, string destinationPath, MoveOptions moveOptions, int retry, int retryTimeout, PathFormat pathFormat)
      {
         return CopyMoveCore(false, new CopyMoveArguments
         {
            Retry = retry,
            RetryTimeout = retryTimeout,
            MoveOptions = moveOptions,
            PathFormat = pathFormat

         }, false, false, sourcePath, destinationPath, null);
      }


      /// <summary>[AlphaFS] 指定されたファイルを新しい場所に移動します。新しいファイル名を指定するオプションを提供します。</summary>
      /// <returns>移動操作のステータスを含む<see cref="CopyMoveResult"/>クラス。</returns>
      /// <remarks>
      ///   <para>このメソッドはディスクボリュームをまたいで動作します。</para>
      ///   <para>同名のファイルをそのディレクトリに移動してファイルを置換しようとすると、<see cref="IOException"/>が発生します。</para>
      ///   <para>Moveメソッドを使用して既存のファイルを上書きすることはできません。</para>
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
      /// <param name="sourcePath">移動するファイルの名前。</param>
      /// <param name="destinationPath">ファイルの新しいパス。</param>
      /// <param name="moveOptions">ファイルの移動方法を指定する<see cref="MoveOptions"/>。このパラメータは<c>null</c>にできます。</param>
      /// <param name="progressHandler">ファイルの別の部分が移動されるたびに呼び出されるコールバック関数。このパラメータは<c>null</c>にできます。</param>
      /// <param name="userProgressData">コールバック関数に渡される引数。このパラメータは<c>null</c>にできます。</param>
      [SecurityCritical]
      public static CopyMoveResult Move(string sourcePath, string destinationPath, MoveOptions moveOptions, CopyMoveProgressRoutine progressHandler, object userProgressData)
      {
         return CopyMoveCore(false, new CopyMoveArguments
         {
            MoveOptions = moveOptions,
            ProgressHandler = progressHandler,
            UserProgressData = userProgressData

         }, false, false, sourcePath, destinationPath, null);
      }


      /// <summary>[AlphaFS] 指定されたファイルを新しい場所に移動します。新しいファイル名を指定するオプションを提供します。</summary>
      /// <returns>移動操作のステータスを含む<see cref="CopyMoveResult"/>クラス。</returns>
      /// <returns>移動操作のステータスを含む<see cref="CopyMoveResult"/>クラス。</returns>
      /// <remarks>
      ///   <para>このメソッドはディスクボリュームをまたいで動作します。</para>
      ///   <para>同名のファイルをそのディレクトリに移動してファイルを置換しようとすると、<see cref="IOException"/>が発生します。</para>
      ///   <para>Moveメソッドを使用して既存のファイルを上書きすることはできません。</para>
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
      /// <param name="sourcePath">移動するファイルの名前。</param>
      /// <param name="destinationPath">ファイルの新しいパス。</param>
      /// <param name="moveOptions">ファイルの移動方法を指定する<see cref="MoveOptions"/>。このパラメータは<c>null</c>にできます。</param>
      /// <param name="progressHandler">ファイルの別の部分が移動されるたびに呼び出されるコールバック関数。このパラメータは<c>null</c>にできます。</param>
      /// <param name="userProgressData">コールバック関数に渡される引数。このパラメータは<c>null</c>にできます。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static CopyMoveResult Move(string sourcePath, string destinationPath, MoveOptions moveOptions, CopyMoveProgressRoutine progressHandler, object userProgressData, PathFormat pathFormat)
      {
         return CopyMoveCore(false, new CopyMoveArguments
         {
            MoveOptions = moveOptions,
            ProgressHandler = progressHandler,
            UserProgressData = userProgressData,
            PathFormat = pathFormat

         }, false, false, sourcePath, destinationPath, null);
      }


      /// <summary>[AlphaFS] 指定されたファイルを新しい場所に移動します。新しいファイル名を指定するオプションを提供します。</summary>
      /// <returns>移動操作のステータスを含む<see cref="CopyMoveResult"/>クラス。</returns>
      /// <remarks>
      ///   <para>このメソッドはディスクボリュームをまたいで動作します。</para>
      ///   <para>同名のファイルをそのディレクトリに移動してファイルを置換しようとすると、<see cref="IOException"/>が発生します。</para>
      ///   <para>Moveメソッドを使用して既存のファイルを上書きすることはできません。</para>
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
      /// <param name="sourcePath">移動するファイルの名前。</param>
      /// <param name="destinationPath">ファイルの新しいパス。</param>
      /// <param name="moveOptions">ファイルの移動方法を指定する<see cref="MoveOptions"/>。このパラメータは<c>null</c>にできます。</param>
      /// <param name="retry">コピー失敗時のリトライ回数。</param>
      /// <param name="retryTimeout">リトライ間の待機時間(秒)。</param>
      /// <param name="progressHandler">ファイルの別の部分が移動されるたびに呼び出されるコールバック関数。このパラメータは<c>null</c>にできます。</param>
      /// <param name="userProgressData">コールバック関数に渡される引数。このパラメータは<c>null</c>にできます。</param>
      [SecurityCritical]
      public static CopyMoveResult Move(string sourcePath, string destinationPath, MoveOptions moveOptions, int retry, int retryTimeout, CopyMoveProgressRoutine progressHandler, object userProgressData)
      {
         return CopyMoveCore(false, new CopyMoveArguments
         {
            Retry = retry,
            RetryTimeout = retryTimeout,
            MoveOptions = moveOptions,
            ProgressHandler = progressHandler,
            UserProgressData = userProgressData

         }, false, false, sourcePath, destinationPath, null);
      }


      /// <summary>[AlphaFS] 指定されたファイルを新しい場所に移動します。新しいファイル名を指定するオプションを提供します。</summary>
      /// <returns>移動操作のステータスを含む<see cref="CopyMoveResult"/>クラス。</returns>
      /// <returns>移動操作のステータスを含む<see cref="CopyMoveResult"/>クラス。</returns>
      /// <remarks>
      ///   <para>このメソッドはディスクボリュームをまたいで動作します。</para>
      ///   <para>同名のファイルをそのディレクトリに移動してファイルを置換しようとすると、<see cref="IOException"/>が発生します。</para>
      ///   <para>Moveメソッドを使用して既存のファイルを上書きすることはできません。</para>
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
      /// <param name="sourcePath">移動するファイルの名前。</param>
      /// <param name="destinationPath">ファイルの新しいパス。</param>
      /// <param name="moveOptions">ファイルの移動方法を指定する<see cref="MoveOptions"/>。このパラメータは<c>null</c>にできます。</param>
      /// <param name="progressHandler">ファイルの別の部分が移動されるたびに呼び出されるコールバック関数。このパラメータは<c>null</c>にできます。</param>
      /// <param name="userProgressData">コールバック関数に渡される引数。このパラメータは<c>null</c>にできます。</param>
      /// <param name="retry">コピー失敗時のリトライ回数。</param>
      /// <param name="retryTimeout">リトライ間の待機時間(秒)。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static CopyMoveResult Move(string sourcePath, string destinationPath, MoveOptions moveOptions, int retry, int retryTimeout, CopyMoveProgressRoutine progressHandler, object userProgressData, PathFormat pathFormat)
      {
         return CopyMoveCore(false, new CopyMoveArguments
         {
            Retry = retry,
            RetryTimeout = retryTimeout,
            MoveOptions = moveOptions,
            ProgressHandler = progressHandler,
            UserProgressData = userProgressData,
            PathFormat = pathFormat

         }, false, false, sourcePath, destinationPath, null);
      }
   }
}
