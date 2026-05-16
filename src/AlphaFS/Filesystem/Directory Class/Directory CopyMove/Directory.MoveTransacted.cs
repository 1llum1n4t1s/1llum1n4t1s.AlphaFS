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
   public static partial class Directory
   {
      /// <summary>[AlphaFS] ファイルまたはディレクトリとその内容を新しい場所に移動します。</summary>
      /// <remarks>
      ///   <para>このメソッドはディスクボリュームをまたいで動作しません。</para>
      ///   <para>可能な限り、このメソッドでは短いファイル名（<c>XXXXXX~1.XXX</c> など）の使用を避けてください。</para>
      ///   <para>2つのディレクトリが同等の短いファイル名を持つ場合、このメソッドは失敗して例外を発生させるか、望ましくない動作を引き起こす可能性があります。</para>
      /// </remarks>
      /// <returns>移動操作の詳細を含む <see cref="CopyMoveResult"/> クラス。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="sourcePath">移動元ディレクトリのパス。</param>
      /// <param name="destinationPath">移動先ディレクトリのパス。</param>
      [SecurityCritical]
      public static CopyMoveResult MoveTransacted(KernelTransaction transaction, string sourcePath, string destinationPath)
      {
         return CopyMoveCore(new CopyMoveArguments
         {
            Transaction = transaction,
            SourcePath = sourcePath,
            DestinationPath = destinationPath
         });
      }


      /// <summary>[AlphaFS] ファイルまたはディレクトリとその内容を新しい場所に移動します。</summary>
      /// <remarks>
      ///   <para>このメソッドはディスクボリュームをまたいで動作しません。</para>
      ///   <para>可能な限り、このメソッドでは短いファイル名（<c>XXXXXX~1.XXX</c> など）の使用を避けてください。</para>
      ///   <para>2つのディレクトリが同等の短いファイル名を持つ場合、このメソッドは失敗して例外を発生させるか、望ましくない動作を引き起こす可能性があります。</para>
      /// </remarks>
      /// <returns>移動操作の詳細を含む <see cref="CopyMoveResult"/> クラス。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="sourcePath">移動元ディレクトリのパス。</param>
      /// <param name="destinationPath">移動先ディレクトリのパス。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static CopyMoveResult MoveTransacted(KernelTransaction transaction, string sourcePath, string destinationPath, PathFormat pathFormat)
      {
         return CopyMoveCore(new CopyMoveArguments
         {
            Transaction = transaction,
            SourcePath = sourcePath,
            DestinationPath = destinationPath,
            PathFormat = pathFormat
         });
      }


      /// <summary>[AlphaFS] ファイルまたはディレクトリとその内容を新しい場所に移動します。<see cref="MoveOptions"/> を指定できます。</summary>
      /// <remarks>
      ///   <para><paramref name="moveOptions"/> に <see cref="MoveOptions.CopyAllowed"/> が含まれていない限り、このメソッドはディスクボリュームをまたいで動作しません。</para>
      ///   <para>可能な限り、このメソッドでは短いファイル名（<c>XXXXXX~1.XXX</c> など）の使用を避けてください。</para>
      ///   <para>2つのディレクトリが同等の短いファイル名を持つ場合、このメソッドは失敗して例外を発生させるか、望ましくない動作を引き起こす可能性があります。</para>
      /// </remarks>
      /// <returns>移動操作の詳細を含む <see cref="CopyMoveResult"/> クラス。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="sourcePath">移動元ディレクトリのパス。</param>
      /// <param name="destinationPath">移動先ディレクトリのパス。</param>
      /// <param name="moveOptions">ディレクトリの移動方法を指定する <see cref="MoveOptions"/>。このパラメータは <c>null</c> にできます。</param>
      [SecurityCritical]
      public static CopyMoveResult MoveTransacted(KernelTransaction transaction, string sourcePath, string destinationPath, MoveOptions moveOptions)
      {
         return CopyMoveCore(new CopyMoveArguments
         {
            Transaction = transaction,
            SourcePath = sourcePath,
            DestinationPath = destinationPath,
            MoveOptions = moveOptions
         });
      }

      /// <summary>[AlphaFS] ファイルまたはディレクトリとその内容を新しい場所に移動します。<see cref="MoveOptions"/> を指定できます。</summary>
      /// <remarks>
      ///   <para><paramref name="moveOptions"/> に <see cref="MoveOptions.CopyAllowed"/> が含まれていない限り、このメソッドはディスクボリュームをまたいで動作しません。</para>
      ///   <para>可能な限り、このメソッドでは短いファイル名（<c>XXXXXX~1.XXX</c> など）の使用を避けてください。</para>
      ///   <para>2つのディレクトリが同等の短いファイル名を持つ場合、このメソッドは失敗して例外を発生させるか、望ましくない動作を引き起こす可能性があります。</para>
      /// </remarks>
      /// <returns>移動操作の詳細を含む <see cref="CopyMoveResult"/> クラス。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="sourcePath">移動元ディレクトリのパス。</param>
      /// <param name="destinationPath">移動先ディレクトリのパス。</param>
      /// <param name="moveOptions">ディレクトリの移動方法を指定する <see cref="MoveOptions"/>。このパラメータは <c>null</c> にできます。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static CopyMoveResult MoveTransacted(KernelTransaction transaction, string sourcePath, string destinationPath, MoveOptions moveOptions, PathFormat pathFormat)
      {
         return CopyMoveCore(new CopyMoveArguments
         {
            Transaction = transaction,
            SourcePath = sourcePath,
            DestinationPath = destinationPath,
            MoveOptions = moveOptions,
            PathFormat = pathFormat
         });
      }


      /// <summary>[AlphaFS] ファイルまたはディレクトリとその内容を新しい場所に移動します。<see cref="MoveOptions"/> を指定でき、
      ///   コールバック関数を通じてアプリケーションに進行状況を通知できます。
      /// </summary>
      /// <remarks>
      ///   <para><paramref name="moveOptions"/> に <see cref="MoveOptions.CopyAllowed"/> が含まれていない限り、このメソッドはディスクボリュームをまたいで動作しません。</para>
      ///   <para>可能な限り、このメソッドでは短いファイル名（<c>XXXXXX~1.XXX</c> など）の使用を避けてください。</para>
      ///   <para>2つのディレクトリが同等の短いファイル名を持つ場合、このメソッドは失敗して例外を発生させるか、望ましくない動作を引き起こす可能性があります。</para>
      /// </remarks>
      /// <returns>移動操作の詳細を含む <see cref="CopyMoveResult"/> クラス。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="sourcePath">移動元ディレクトリのパス。</param>
      /// <param name="destinationPath">移動先ディレクトリのパス。</param>
      /// <param name="moveOptions">ディレクトリの移動方法を指定する <see cref="MoveOptions"/>。このパラメータは <c>null</c> にできます。</param>
      /// <param name="progressHandler">ディレクトリの一部が移動されるたびに呼び出されるコールバック関数。このパラメータは <c>null</c> にできます。</param>
      /// <param name="userProgressData">コールバック関数に渡される引数。このパラメータは <c>null</c> にできます。</param>
      [SecurityCritical]
      public static CopyMoveResult MoveTransacted(KernelTransaction transaction, string sourcePath, string destinationPath, MoveOptions moveOptions, CopyMoveProgressRoutine progressHandler, object userProgressData)
      {
         return CopyMoveCore(new CopyMoveArguments
         {
            Transaction = transaction,
            SourcePath = sourcePath,
            DestinationPath = destinationPath,
            MoveOptions = moveOptions,
            ProgressHandler = progressHandler,
            UserProgressData = userProgressData
         });
      }


      /// <summary>[AlphaFS] ファイルまたはディレクトリとその内容を新しい場所に移動します。<see cref="MoveOptions"/> を指定でき、
      ///   コールバック関数を通じてアプリケーションに進行状況を通知できます。
      /// </summary>
      /// <returns>移動操作のステータスを含む <see cref="CopyMoveResult"/> クラス。</returns>
      /// <remarks>
      ///   <para><paramref name="moveOptions"/> に <see cref="MoveOptions.CopyAllowed"/> が含まれていない限り、このメソッドはディスクボリュームをまたいで動作しません。</para>
      ///   <para>可能な限り、このメソッドでは短いファイル名（<c>XXXXXX~1.XXX</c> など）の使用を避けてください。</para>
      ///   <para>2つのディレクトリが同等の短いファイル名を持つ場合、このメソッドは失敗して例外を発生させるか、望ましくない動作を引き起こす可能性があります。</para>
      /// </remarks>
      /// <returns>移動操作の詳細を含む <see cref="CopyMoveResult"/> クラス。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="sourcePath">移動元ディレクトリのパス。</param>
      /// <param name="destinationPath">移動先ディレクトリのパス。</param>
      /// <param name="moveOptions">ディレクトリの移動方法を指定する <see cref="MoveOptions"/>。このパラメータは <c>null</c> にできます。</param>
      /// <param name="progressHandler">ディレクトリの一部が移動されるたびに呼び出されるコールバック関数。このパラメータは <c>null</c> にできます。</param>
      /// <param name="userProgressData">コールバック関数に渡される引数。このパラメータは <c>null</c> にできます。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static CopyMoveResult MoveTransacted(KernelTransaction transaction, string sourcePath, string destinationPath, MoveOptions moveOptions, CopyMoveProgressRoutine progressHandler, object userProgressData, PathFormat pathFormat)
      {
         return CopyMoveCore(new CopyMoveArguments
         {
            Transaction = transaction,
            SourcePath = sourcePath,
            DestinationPath = destinationPath,
            MoveOptions = moveOptions,
            ProgressHandler = progressHandler,
            UserProgressData = userProgressData,
            PathFormat = pathFormat
         });
      }
   }
}
