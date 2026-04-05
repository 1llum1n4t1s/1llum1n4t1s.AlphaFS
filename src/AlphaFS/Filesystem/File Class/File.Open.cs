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

using System.IO;
using System.Security;
using System.Security.AccessControl;
using FileStream = System.IO.FileStream;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class File
   {
      #region Using FileAccess

      #region .NET

      /// <summary>指定されたパスで読み取り/書き込みアクセスの<see cref="FileStream"/>を開きます。</summary>
      /// <param name="path">開くファイル。</param>
      /// <param name="mode">ファイルが存在しない場合に作成するかどうか、および既存ファイルの内容を保持するか上書きするかを指定する<see cref="FileMode"/>値。</param>
      /// <returns>指定されたモードとパスで、読み取り/書き込みアクセスで非共有の<see cref="FileStream"/>。</returns>
      [SecurityCritical]
      public static FileStream Open(string path, FileMode mode)
      {
         return OpenCore(null, path, mode, mode == FileMode.Append ? FileAccess.Write : FileAccess.ReadWrite, FileShare.None, ExtendedFileAttributes.Normal, null, null, PathFormat.RelativePath);
      }


      /// <summary>指定されたパスで、指定されたモードとアクセスの<see cref="FileStream"/>を開きます。</summary>
      /// <param name="path">開くファイル。</param>
      /// <param name="mode">ファイルが存在しない場合に作成するかどうか、および既存ファイルの内容を保持するか上書きするかを指定する<see cref="FileMode"/>値。</param>
      /// <param name="access">ファイルに対して実行できる操作を指定する<see cref="FileAccess"/>値。</param>
      /// <returns>指定されたモードとアクセスで指定されたファイルへのアクセスを提供する非共有<see cref="FileStream"/>。</returns>
      [SecurityCritical]
      public static FileStream Open(string path, FileMode mode, FileAccess access)
      {
         return OpenCore(null, path, mode, access, FileShare.None, ExtendedFileAttributes.Normal, null, null, PathFormat.RelativePath);
      }


      /// <summary>指定されたパスで、指定されたモード、読み取り/書き込みまたは読み書きアクセス、および指定された共有オプションの<see cref="FileStream"/>を開きます。</summary>
      /// <param name="path">開くファイル。</param>
      /// <param name="mode">ファイルが存在しない場合に作成するかどうか、および既存ファイルの内容を保持するか上書きするかを指定する<see cref="FileMode"/>値。</param>
      /// <param name="access">ファイルに対して実行できる操作を指定する<see cref="FileAccess"/>値。</param>
      /// <param name="share">他のスレッドがファイルに対して持つアクセスの種類を指定する<see cref="FileShare"/>値。</param>
      /// <returns>指定されたパスで、指定されたモード、読み取り/書き込みまたは読み書きアクセス、 および指定された共有オプションの<see cref="FileStream"/>。</returns>
      [SecurityCritical]
      public static FileStream Open(string path, FileMode mode, FileAccess access, FileShare share)
      {
         return OpenCore(null, path, mode, access, share, ExtendedFileAttributes.Normal, null, null, PathFormat.RelativePath);
      }

      #endregion // .NET


      /// <summary>[AlphaFS] 指定されたパスで読み取り/書き込みアクセスの<see cref="FileStream"/>を開きます。</summary>
      /// <param name="path">開くファイル。</param>
      /// <param name="mode">
      ///   ファイルが存在しない場合に作成するかどうか、および既存ファイルの内容を
      ///   保持するか上書きするかを指定する<see cref="FileMode"/>値。
      /// </param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      /// <returns>指定されたモードとパスで、読み取り/書き込みアクセスで非共有の<see cref="FileStream"/>。</returns>
      [SecurityCritical]
      public static FileStream Open(string path, FileMode mode, PathFormat pathFormat)
      {
         return OpenCore(null, path, mode, mode == FileMode.Append ? FileAccess.Write : FileAccess.ReadWrite, FileShare.None, ExtendedFileAttributes.Normal, null, null, pathFormat);
      }


      /// <summary>[AlphaFS] 指定されたパスで、指定されたモードとアクセスの<see cref="FileStream"/>を開きます。</summary>
      /// <param name="path">開くファイル。</param>
      /// <param name="mode">
      ///   ファイルが存在しない場合に作成するかどうか、および既存ファイルの内容を
      ///   保持するか上書きするかを指定する<see cref="FileMode"/>値。
      /// </param>
      /// <param name="access">ファイルに対して実行できる操作を指定する<see cref="FileAccess"/>値。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      /// <returns>
      ///   An unshared <see cref="FileStream"/> that provides access to the specified file, with the specified mode and access.
      /// </returns>
      [SecurityCritical]
      public static FileStream Open(string path, FileMode mode, FileAccess access, PathFormat pathFormat)
      {
         return OpenCore(null, path, mode, access, FileShare.None, ExtendedFileAttributes.Normal, null, null, pathFormat);
      }


      /// <summary>[AlphaFS] 指定されたパスで、指定されたモード、読み取り/書き込みまたは読み書きアクセス、および指定された共有オプションの<see cref="FileStream"/>を開きます。</summary>
      /// <param name="path">開くファイル。</param>
      /// <param name="mode">
      ///   ファイルが存在しない場合に作成するかどうか、および既存ファイルの内容を
      ///   保持するか上書きするかを指定する<see cref="FileMode"/>値。
      /// </param>
      /// <param name="access">ファイルに対して実行できる操作を指定する<see cref="FileAccess"/>値。</param>
      /// <param name="share">他のスレッドがファイルに対して持つアクセスの種類を指定する<see cref="FileShare"/>値。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      /// <returns>
      ///   指定されたパスで、指定されたモード、読み取り/書き込みまたは読み書きアクセス、 access and the
      ///   および指定された共有オプション。
      /// </returns>
      [SecurityCritical]
      public static FileStream Open(string path, FileMode mode, FileAccess access, FileShare share, PathFormat pathFormat)
      {
         return OpenCore(null, path, mode, access, share, ExtendedFileAttributes.Normal, null, null, pathFormat);
      }


      /// <summary>[AlphaFS] 指定されたパスで、指定されたモード、読み取り/書き込みまたは読み書きアクセス、および指定された共有オプションの<see cref="FileStream"/>を開きます。</summary>
      /// <param name="path">開くファイル。</param>
      /// <param name="mode">
      ///   ファイルが存在しない場合に作成するかどうか、および既存ファイルの内容を
      ///   保持するか上書きするかを指定する<see cref="FileMode"/>値。
      /// </param>
      /// <param name="access">ファイルに対して実行できる操作を指定する<see cref="FileAccess"/>値。</param>
      /// <param name="share">他のスレッドがファイルに対して持つアクセスの種類を指定する<see cref="FileShare"/>値。</param>
      /// <param name="extendedAttributes">The extended attributes.</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      /// <returns>
      ///   指定されたパスで、指定されたモード、読み取り/書き込みまたは読み書きアクセス、 access and the
      ///   および指定された共有オプション。
      /// </returns>
      [SecurityCritical]
      public static FileStream Open(string path, FileMode mode, FileAccess access, FileShare share, ExtendedFileAttributes extendedAttributes, PathFormat pathFormat)
      {
         return OpenCore(null, path, mode, access, share, extendedAttributes, null, null, pathFormat);
      }
      

      /// <summary>[AlphaFS] 指定されたパスで、指定された作成モード、読み取り/書き込みおよび���有権限、バッファサイズを使用して<see cref="FileStream"/>を開きます。</summary>
      /// <param name="path">開くファイル。</param>
      /// <param name="mode">ファイルの開き方または���成方法を決定する定数。</param>
      /// <param name="access">ファイルに対して実行できる操作を指定する<see cref="FileAccess"/>値。</param>
      
      /// <param name="share">プロセスによるファイルの共有方法を決定する定数。</param>
      /// <param name="bufferSize">0より大きい正の<see cref="System.Int32"/>値でバッファサイズを示します。 default buffer size is 4096. </param>
      /// <returns>
      ///   指定されたパスで、指定されたモード、読み取り/書き込みまたは読み書きアクセス、
      ///   および指定された共有オプションの<see cref="FileStream"/>。
      /// </returns>
      [SecurityCritical]
      public static FileStream Open(string path, FileMode mode, FileAccess access, FileShare share, int bufferSize)
      {
         return OpenCore(null, path, mode, access, share, ExtendedFileAttributes.Normal, bufferSize, null, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] 指定されたパスで、指定された作成モード、読み取り/書き込みおよび���有権限、バッファサイズを使用して<see cref="FileStream"/>を開きます。</summary>
      /// <param name="path">開くファイル。</param>
      /// <param name="mode">ファイルの開き方または���成方法を決定する定数。</param>
      /// <param name="access">ファイルに対して実行できる操作を指定する<see cref="FileAccess"/>値。</param>
      
      /// <param name="share">プロセスによるファイルの共有方法を決定する定数。</param>
      /// <param name="bufferSize">0より大きい正の<see cref="System.Int32"/>値でバッファサイズを示します。
      /// デフォルトのバッファサイズは4096です。</param>
      /// <param name="useAsync">非同期I/Oまたは同期I/Oを使用するかどうかを指定します。ただし、
      /// 基盤となるオペレーティングシステムが非同期I/Oをサポートしていない場合があるため、trueを指定しても
      /// プラットフォームによっては同期的に開かれる場合があります。非同期で開いた場合、BeginReadとBeginWriteメソッドは
      /// 大きな読み取りまたは書き込みでパフォーマンスが向上しますが、小さな読み取りまたは書き込みでは大幅に遅くなる可能性があります。
      /// アプリケーションが非同期I/Oを活用するように設計されている場合は、useAsyncパラメータをtrueに設定してください。
      /// 非同期I/Oを正しく使用すると、アプリケーションを最大10倍高速化できますが、
      /// 非同期I/O用にアプリケーションを再設計せずに使用すると、パフォーマンスが最大10分の1に低下する可能性があります。</param>
      /// <returns>
      ///   指定されたパスで、指定されたモード、読み取り/書き込みまたは読み書きアクセス、
      ///   および指定された共有オプションの<see cref="FileStream"/>。
      /// </returns>
      [SecurityCritical]
      public static FileStream Open(string path, FileMode mode, FileAccess access, FileShare share, int bufferSize, bool useAsync)
      {
         return OpenCore(null, path, mode, access, share, ExtendedFileAttributes.Normal | (useAsync ? ExtendedFileAttributes.Overlapped : ExtendedFileAttributes.Normal), bufferSize, null, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] 指定されたパスで、指定された作成モード、読み取り/書き込みおよび���有権限、バッファサイズを使用して<see cref="FileStream"/>を開きます。</summary>
      /// <param name="path">開くファイル。</param>
      /// <param name="mode">ファイルの開き方または���成方法を決定する定数。</param>
      /// <param name="access">ファイルに対して実行できる操作を指定する<see cref="FileAccess"/>値。</param>
      
      /// <param name="share">プロセスによるファイルの共有方法を決定する定数。</param>
      /// <param name="bufferSize">0より大きい正の<see cref="System.Int32"/>値でバッファサイズを示します。
      /// デフォルトのバッファサイズは4096です。</param>
      /// <param name="options">追加のファイルオプション���指定する値。</param>
      /// <returns>
      ///   指定されたパスで、指定されたモード、読み取り/書き込みまたは読み書きアクセス、
      ///   および指定された共有オプションの<see cref="FileStream"/>。
      /// </returns>
      [SecurityCritical]
      public static FileStream Open(string path, FileMode mode, FileAccess access, FileShare share, int bufferSize, FileOptions options)
      {
         return OpenCore(null, path, mode, access, share, (ExtendedFileAttributes) options, bufferSize, null, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] 指定されたパスで、指定された作成モード、読み取り/書き込みおよび���有権限、バッファサイズを使用して<see cref="FileStream"/>を開きます。</summary>
      /// <param name="path">開くファイル。</param>
      /// <param name="mode">ファイルの開き方または���成方法を決定する定数。</param>
      /// <param name="access">ファイルに対して実行できる操作を指定する<see cref="FileAccess"/>値。</param>
      
      /// <param name="share">プロセスによるファイルの共有方法を決定する定数。</param>
      /// <param name="bufferSize">0より大きい正の<see cref="System.Int32"/>値でバッファサイズを示します。
      /// デフォルトのバッファサイズは4096です。</param>
      /// <param name="extendedAttributes">追加オプションを指定する拡張属性。</param>
      /// <returns>
      ///   指定されたパスで、指定されたモード、読み取り/書き込みまたは読み書きアクセス、
      ///   および指定された共有オプションの<see cref="FileStream"/>。
      /// </returns>      
      [SecurityCritical]
      public static FileStream Open(string path, FileMode mode, FileAccess access, FileShare share, int bufferSize, ExtendedFileAttributes extendedAttributes)
      {
         return OpenCore(null, path, mode, access, share, extendedAttributes, bufferSize, null, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] 指定されたパスで、指定された作成モード、読み取り/書き込みおよび���有権限、バッファサイズを使用して<see cref="FileStream"/>を開きます。</summary>
      /// <param name="path">開くファイル。</param>
      /// <param name="mode">ファイルの開き方または���成方法を決定する定数。</param>
      /// <param name="access">ファイルに対して実行できる操作を指定する<see cref="FileAccess"/>値。</param>
      
      /// <param name="share">プロセスによるファイルの共有方法を決定する定数。</param>
      /// <param name="bufferSize">0より大きい正の<see cref="System.Int32"/>値でバッファサイズを示します。
      /// デフォルトのバッファサイズは4096です。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      /// <returns>
      ///   指定されたパスで、指定されたモード、読み取り/書き込みまたは読み書きアクセス、
      ///   および指定された共有オプションの<see cref="FileStream"/>。
      /// </returns>
      [SecurityCritical]
      public static FileStream Open(string path, FileMode mode, FileAccess access, FileShare share, int bufferSize, PathFormat pathFormat)
      {
         return OpenCore(null, path, mode, access, share, ExtendedFileAttributes.Normal, bufferSize, null, pathFormat);
      }


      /// <summary>[AlphaFS] 指定されたパスで、指定された作成モード、読み取り/書き込みおよび���有権限、バッファサイズを使用して<see cref="FileStream"/>を開きます。</summary>
      /// <param name="path">開くファイル。</param>
      /// <param name="mode">ファイルの開き方または���成方法を決定する定数。</param>
      /// <param name="access">ファイルに対して実行できる操作を指定する<see cref="FileAccess"/>値。</param>
      
      /// <param name="share">プロセスによるファイルの共有方法を決定する定数。</param>
      /// <param name="bufferSize">0より大きい正の<see cref="System.Int32"/>値でバッファサイズを示します。
      /// デフォルトのバッファサイズは4096です。</param>
      /// <param name="useAsync">非同期I/Oまたは同期I/Oを使用するかどうかを指定します。ただし、
      /// 基盤となるオペレーティングシステムが非同期I/Oをサポートしていない場合があるため、trueを指定しても
      /// プラットフォームによっては同期的に開かれる場合があります。非同期で開いた場合、BeginReadとBeginWriteメソッドは
      /// 大きな読み取りまたは書き込みでパフォーマンスが向上しますが、小さな読み取りまたは書き込みでは大幅に遅くなる可能性があります。
      /// アプリケーションが非同期I/Oを活用するように設計されている場合は、useAsyncパラメータをtrueに設定してください。
      /// 非同期I/Oを正しく使用すると、アプリケーションを最大10倍高速化できますが、
      /// 非同期I/O用にアプリケーションを再設計せずに使用すると、パフォーマンスが最大10分の1に低下する可能性があります。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      /// <returns>
      ///   指定されたパスで、指定されたモード、読み取り/書き込みまたは読み書きアクセス、
      ///   および指定された共有オプションの<see cref="FileStream"/>。
      /// </returns>
      [SecurityCritical]
      public static FileStream Open(string path, FileMode mode, FileAccess access, FileShare share, int bufferSize, bool useAsync, PathFormat pathFormat)
      {
         return OpenCore(null, path, mode, access, share, ExtendedFileAttributes.Normal | (useAsync ? ExtendedFileAttributes.Overlapped : ExtendedFileAttributes.Normal), bufferSize, null, pathFormat);
      }


      /// <summary>[AlphaFS] 指定されたパスで、指定された作成モード、読み取り/書き込みおよび���有権限、バッファサイズを使用して<see cref="FileStream"/>を開きます。</summary>
      /// <param name="path">開くファイル。</param>
      /// <param name="mode">ファイルの開き方または���成方法を決定する定数。</param>
      /// <param name="access">ファイルに対して実行できる操作を指定する<see cref="FileAccess"/>値。</param>
      
      /// <param name="share">プロセスによるファイルの共有方法を決定する定数。</param>
      /// <param name="bufferSize">0より大きい正の<see cref="System.Int32"/>値でバッファサイズを示します。
      /// デフォルトのバッファサイズは4096です。</param>
      /// <param name="options">追加のファイルオプション���指定する値。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      /// <returns>
      ///   指定されたパスで、指定されたモード、読み取り/書き込みまたは読み書きアクセス、
      ///   および指定された共有オプションの<see cref="FileStream"/>。
      /// </returns>
      [SecurityCritical]
      public static FileStream Open(string path, FileMode mode, FileAccess access, FileShare share, int bufferSize, FileOptions options, PathFormat pathFormat)
      {
         return OpenCore(null, path, mode, access, share, (ExtendedFileAttributes) options, bufferSize, null, pathFormat);
      }


      /// <summary>[AlphaFS] 指定されたパスで、指定された作成モード、読み取り/書き込みおよび���有権限、バッファサイズを使用して<see cref="FileStream"/>を開きます。</summary>
      /// <param name="path">開くファイル。</param>
      /// <param name="mode">ファイルの開き方または���成方法を決定する定数。</param>
      /// <param name="access">ファイルに対して実行できる操作を指定する<see cref="FileAccess"/>値。</param>
      
      /// <param name="share">プロセスによるファイルの共有方法を決定する定数。</param>
      /// <param name="bufferSize">0より大きい正の<see cref="System.Int32"/>値でバッファサイズを示します。
      /// デフォルトのバッファサイズは4096です。</param>
      /// <param name="extendedAttributes">追加オプションを指定する拡張属性。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      /// <returns>
      ///   指定されたパスで、指定されたモード、読み取り/書き込みまたは読み書きアクセス、
      ///   および指定された共有オプションの<see cref="FileStream"/>。
      /// </returns>
      [SecurityCritical]
      public static FileStream Open(string path, FileMode mode, FileAccess access, FileShare share, int bufferSize, ExtendedFileAttributes extendedAttributes, PathFormat pathFormat)
      {
         return OpenCore(null, path, mode, access, share, extendedAttributes, bufferSize, null, pathFormat);
      }
      
      #endregion // Using FileAccess


      #region Using FileSystemRights

      /// <summary>[AlphaFS] 指定されたパスで、指定された作成モード、読み取り/書き込みおよび���有権限、バッファサイズを使用して<see cref="FileStream"/>を開きます。</summary>
      /// <param name="path">開くファイル。</param>
      /// <param name="mode">ファイルの開き方または���成方法を決定する定数。</param>
      /// <param name="rights">ファイルに対して実行できる操作を指定する<see cref="FileAccess"/>値。</param>
      
      /// <param name="share">プロセスによるファイルの共有方法を決定する定数。</param>
      /// <param name="bufferSize">0より大きい正の<see cref="System.Int32"/>値でバッファサイズを示します。
      /// デフォルトのバッファサイズは4096です。</param>
      /// <param name="options">追加のファイルオプション���指定する値。</param>
      /// <returns>
      ///   指定されたパスで、指定されたモード、読み取り/書き込みまたは読み書きアクセス、
      ///   および指定された共有オプションの<see cref="FileStream"/>。
      /// </returns>
      [SecurityCritical]
      public static FileStream Open(string path, FileMode mode, FileSystemRights rights, FileShare share, int bufferSize, FileOptions options)
      {
         return OpenCore(null, path, mode, rights, share, (ExtendedFileAttributes)options, bufferSize, null, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] 指定されたパスで、指定された作成モード、読み取り/書き込みおよび���有権限、バッファサイズを使用して<see cref="FileStream"/>を開きます。</summary>
      /// <param name="path">開くファイル。</param>
      /// <param name="mode">ファイルの開き方または���成方法を決定する定数。</param>
      /// <param name="rights">ファイルに対して実行できる操作を指定する<see cref="FileAccess"/>値。</param>
      
      /// <param name="share">プロセスによるファイルの共有方法を決定する定数。</param>
      /// <param name="bufferSize">0より大きい正の<see cref="System.Int32"/>値でバッファサイズを示します。
      /// デフォルトのバッファサイズは4096です。</param>
      /// <param name="extendedAttributes">追加オプションを指定する拡張属性。</param>
      /// <returns>
      ///   指定されたパスで、指定されたモード、読み取り/書き込みまたは読み書きアクセス、
      ///   および指定された共有オプションの<see cref="FileStream"/>。
      /// </returns>
      [SecurityCritical]
      public static FileStream Open(string path, FileMode mode, FileSystemRights rights, FileShare share, int bufferSize, ExtendedFileAttributes extendedAttributes)
      {
         return OpenCore(null, path, mode, rights, share, extendedAttributes, bufferSize, null, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] 指定されたパスで、指定された作成モード、読み取り/書き込みおよび���有権限、バッファサイズを使用して<see cref="FileStream"/>を開きます。</summary>
      /// <param name="path">開くファイル。</param>
      /// <param name="mode">ファイルの開き方または���成方法を決定する定数。</param>
      /// <param name="rights">ファイルに対して実行できる操作を指定する<see cref="FileAccess"/>値。</param>
      
      /// <param name="share">プロセスによるファイルの共有方法を決定する定数。</param>
      /// <param name="bufferSize">0より大きい正の<see cref="System.Int32"/>値でバッファサイズを示します。
      /// デフォルトのバッファサイズは4096です。</param>
      /// <param name="options">追加のファイルオプション���指定する値。</param>
      /// <param name="security">ファイルのアクセス制御と監査セキュリティを決定する値。</param>
      /// <returns>
      ///   指定されたパスで、指定されたモード、読み取り/書き込みまたは読み書きアクセス、
      ///   および指定された共有オプションの<see cref="FileStream"/>。
      /// </returns>
      [SecurityCritical]
      public static FileStream Open(string path, FileMode mode, FileSystemRights rights, FileShare share, int bufferSize, FileOptions options, FileSecurity security)
      {
         return OpenCore(null, path, mode, rights, share, (ExtendedFileAttributes)options, bufferSize, security, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] 指定されたパスで、指定された作成モード、読み取り/書き込みおよび���有権限、バッファサイズを使用して<see cref="FileStream"/>を開きます。</summary>
      /// <param name="path">開くファイル。</param>
      /// <param name="mode">ファイルの開き方または���成方法を決定する定数。</param>
      /// <param name="rights">ファイルに対して実行できる操作を指定する<see cref="FileAccess"/>値。</param>
      
      /// <param name="share">プロセスによるファイルの共有方法を決定する定数。</param>
      /// <param name="bufferSize">0より大きい正の<see cref="System.Int32"/>値でバッファサイズを示します。
      /// デフォルトのバッファサイズは4096です。</param>
      /// <param name="extendedAttributes">追加オプションを指定する拡張属性。</param>
      /// <param name="security">ファイルのアクセス制御と監査セキュリティを決定する値。</param>
      /// <returns>
      ///   指定されたパスで、指定されたモード、読み取り/書き込みまたは読み書きアクセス、
      ///   および指定された共有オプションの<see cref="FileStream"/>。
      /// </returns>
      [SecurityCritical]
      public static FileStream Open(string path, FileMode mode, FileSystemRights rights, FileShare share, int bufferSize, ExtendedFileAttributes extendedAttributes, FileSecurity security)
      {
         return OpenCore(null, path, mode, rights, share, extendedAttributes, bufferSize, security, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] 指定されたパスで、指定された作成モード、読み取り/書き込みおよび���有権限、バッファサイズを使用して<see cref="FileStream"/>を開きます。</summary>
      /// <param name="path">開くファイル。</param>
      /// <param name="mode">ファイルの開き方または���成方法を決定する定数。</param>
      /// <param name="rights">ファイルに対して実行できる操作を指定する<see cref="FileAccess"/>値。</param>
      
      /// <param name="share">プロセスによるファイルの共有方法を決定する定数。</param>
      /// <param name="bufferSize">0より大きい正の<see cref="System.Int32"/>値でバッファサイズを示します。
      /// デフォルトのバッファサイズは4096です。</param>
      /// <param name="options">追加のファイルオプション���指定する値。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      /// <returns>
      ///   指定されたパスで、指定されたモード、読み取り/書き込みまたは読み書きアクセス、
      ///   および指定された共有オプションの<see cref="FileStream"/>。
      /// </returns>
      [SecurityCritical]
      public static FileStream Open(string path, FileMode mode, FileSystemRights rights, FileShare share, int bufferSize, FileOptions options, PathFormat pathFormat)
      {
         return OpenCore(null, path, mode, rights, share, (ExtendedFileAttributes) options, bufferSize, null, pathFormat);
      }


      /// <summary>[AlphaFS] 指定されたパスで、指定された作成モード、読み取り/書き込みおよび���有権限、バッファサイズを使用して<see cref="FileStream"/>を開きます。</summary>
      /// <param name="path">開くファイル。</param>
      /// <param name="mode">ファイルの開き方または���成方法を決定する定数。</param>
      /// <param name="rights">ファイルに対して実行できる操作を指定する<see cref="FileAccess"/>値。</param>
      
      /// <param name="share">プロセスによるファイルの共有方法を決定する定数。</param>
      /// <param name="bufferSize">0より大きい正の<see cref="System.Int32"/>値でバッファサイズを示します。
      /// デフォルトのバッファサイズは4096です。</param>
      /// <param name="extendedAttributes">追加オプションを指定する拡張属性。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      /// <returns>
      ///   指定されたパスで、指定されたモード、読み取り/書き込みまたは読み書きアクセス、
      ///   および指定された共有オプションの<see cref="FileStream"/>。
      /// </returns>
      [SecurityCritical]
      public static FileStream Open(string path, FileMode mode, FileSystemRights rights, FileShare share, int bufferSize, ExtendedFileAttributes extendedAttributes, PathFormat pathFormat)
      {
         return OpenCore(null, path, mode, rights, share, extendedAttributes, bufferSize, null, pathFormat);
      }


      /// <summary>[AlphaFS] Opens a <see cref="FileStream"/> on the specified path using the specified  creation mode, access rights and sharing permission, the buffer size, additional file options, access control and audit security.</summary>
      /// <param name="path">開くファイル。</param>
      /// <param name="mode">ファイルの開き方または���成方法を決定する定数。</param>
      /// <param name="rights">ファイルに対して実行できる操作を指定する<see cref="FileAccess"/>値。</param>
      
      /// <param name="share">プロセスによるファイルの共有方法を決定する定数。</param>
      /// <param name="bufferSize">0より大きい正の<see cref="System.Int32"/>値でバッファサイズを示します。
      /// デフォルトのバッファサイズは4096です。</param>
      /// <param name="options">追加のファイルオプション���指定する値。</param>
      /// <param name="security">ファイルのアクセス制御と監査セキュリティを決定する値。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      /// <returns>
      ///   指定されたパスで、指定されたモード、読み取り/書き込みまたは読み書きアクセス、
      ///   および指定された共有オプションの<see cref="FileStream"/>。
      /// </returns>
      [SecurityCritical]
      public static FileStream Open(string path, FileMode mode, FileSystemRights rights, FileShare share, int bufferSize, FileOptions options, FileSecurity security, PathFormat pathFormat)
      {
         return OpenCore(null, path, mode, rights, share, (ExtendedFileAttributes) options, bufferSize, security, pathFormat);
      }


      /// <summary>[AlphaFS] Opens a <see cref="FileStream"/> on the specified path using the specified  creation mode, access rights and sharing permission, the buffer size, additional file options, access control and audit security.</summary>
      /// <param name="path">開くファイル。</param>
      /// <param name="mode">ファイルの開き方または���成方法を決定する定数。</param>
      /// <param name="rights">ファイルに対して実行できる操作を指定する<see cref="FileAccess"/>値。</param>
      
      /// <param name="share">プロセスによるファイルの共有方法を決定する定数。</param>
      /// <param name="bufferSize">0より大きい正の<see cref="System.Int32"/>値でバッファサイズを示します。
      /// デフォルトのバッファサイズは4096です。</param>
      /// <param name="extendedAttributes">追加オプションを指定する拡張属性。</param>
      /// <param name="security">ファイルのアクセス制御と監査セキュリティを決定する値。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      /// <returns>
      ///   指定されたパスで、指定されたモード、読み取り/書き込みまたは読み書きアクセス、
      ///   および指定された共有オプションの<see cref="FileStream"/>。
      /// </returns>
      [SecurityCritical]
      public static FileStream Open(string path, FileMode mode, FileSystemRights rights, FileShare share, int bufferSize, ExtendedFileAttributes extendedAttributes, FileSecurity security, PathFormat pathFormat)
      {
         return OpenCore(null, path, mode, rights, share, extendedAttributes, bufferSize, security, pathFormat);
      }

      #endregion // Using FileSystemRights
   }
}
