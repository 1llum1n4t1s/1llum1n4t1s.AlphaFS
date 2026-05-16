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

using Alphaleonis.Win32.Security;
using Microsoft.Win32.SafeHandles;
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Runtime.InteropServices;
using System.Security;
using System.Security.AccessControl;
using System.Text;
using SecurityNativeMethods = Alphaleonis.Win32.Security.NativeMethods;

namespace Alphaleonis.Win32.Filesystem
{
   /// <summary><see cref="BackupFileStream"/> は、セキュリティ情報や代替データストリームを含む、特定のファイルまたはディレクトリに関連付けられたデータへのアクセスを提供し、バックアップおよび復元操作に使用します。</summary>
   /// <remarks>このクラスは Win32 API の <see href="http://msdn.microsoft.com/en-us/library/aa362509(VS.85).aspx">BackupRead</see>、
   /// <see href="http://msdn.microsoft.com/en-us/library/aa362510(VS.85).aspx">BackupSeek</see>、および
   /// <see href="http://msdn.microsoft.com/en-us/library/aa362511(VS.85).aspx">BackupWrite</see> 関数を使用してファイルまたはディレクトリへのアクセスを提供します。
   /// </remarks>
   public sealed class BackupFileStream : Stream
   {
      private readonly bool _canRead;
      private readonly bool _canWrite;
      private readonly bool _processSecurity;

      [SuppressMessage("Microsoft.Reliability", "CA2006:UseSafeHandleToEncapsulateNativeResources")]
      private IntPtr _context = IntPtr.Zero;




      /// <summary>指定されたパスと作成モードで <see cref="BackupFileStream"/> クラスの新しいインスタンスを初期化します。</summary>
      /// <param name="path">現在の <see cref="BackupFileStream"/> オブジェクトがカプセル化するファイルの相対パスまたは絶対パス。</param>
      /// <param name="mode">ファイルを開くまたは作成する方法を決定する <see cref="FileMode"/> 定数。</param>
      /// <remarks>ファイルは読み取りと書き込みの両方で排他アクセスとして開かれます。</remarks>
      [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope")]
      [SecurityCritical]
      public BackupFileStream(string path, FileMode mode) : this(File.CreateFileCore(null, false, path, ExtendedFileAttributes.Normal, null, mode, FileSystemRights.Read | FileSystemRights.Write, FileShare.None, true, false, PathFormat.RelativePath), FileSystemRights.Read | FileSystemRights.Write)
      {
      }


      /// <summary>指定されたパス、作成モード、およびアクセス権で <see cref="BackupFileStream"/> クラスの新しいインスタンスを初期化します。</summary>
      /// <param name="path">現在の <see cref="BackupFileStream"/> オブジェクトがカプセル化するファイルの相対パスまたは絶対パス。</param>
      /// <param name="mode">ファイルを開くまたは作成する方法を決定する <see cref="FileMode"/> 定数。</param>
      /// <param name="access">ファイルのアクセスルールおよび監査ルールを作成する際に使用するアクセス権を決定する <see cref="FileSystemRights"/> 定数。</param>
      /// <remarks>ファイルは排他アクセスとして開かれます。</remarks>
      [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope")]
      [SecurityCritical]
      public BackupFileStream(string path, FileMode mode, FileSystemRights access) : this(File.CreateFileCore(null, false, path, ExtendedFileAttributes.Normal, null, mode, access, FileShare.None, true, false, PathFormat.RelativePath), access)
      {
      }


      /// <summary>指定されたパス、作成モード、アクセス権、および共有許可で <see cref="BackupFileStream"/> クラスの新しいインスタンスを初期化します。</summary>
      /// <param name="path">現在の <see cref="BackupFileStream"/> オブジェクトがカプセル化するファイルの相対パスまたは絶対パス。</param>
      /// <param name="mode">ファイルを開くまたは作成する方法を決定する <see cref="FileMode"/> 定数。</param>
      /// <param name="access">ファイルのアクセスルールおよび監査ルールを作成する際に使用するアクセス権を決定する <see cref="FileSystemRights"/> 定数。</param>
      /// <param name="share">プロセス間でファイルを共有する方法を決定する <see cref="FileShare"/> 定数。</param>
      [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope")]
      [SecurityCritical]
      public BackupFileStream(string path, FileMode mode, FileSystemRights access, FileShare share) : this(File.CreateFileCore(null, false, path, ExtendedFileAttributes.Normal, null, mode, access, share, true, false, PathFormat.RelativePath), access)
      {
      }


      /// <summary>指定されたパス、作成モード、アクセス権、共有許可、および追加のファイル属性で <see cref="BackupFileStream"/> クラスの新しいインスタンスを初期化します。</summary>
      /// <param name="path">現在の <see cref="BackupFileStream"/> オブジェクトがカプセル化するファイルの相対パスまたは絶対パス。</param>
      /// <param name="mode">ファイルを開くまたは作成する方法を決定する <see cref="FileMode"/> 定数。</param>
      /// <param name="access">ファイルのアクセスルールおよび監査ルールを作成する際に使用するアクセス権を決定する <see cref="FileSystemRights"/> 定数。</param>
      /// <param name="share">プロセス間でファイルを共有する方法を決定する <see cref="FileShare"/> 定数。</param>
      /// <param name="attributes">追加のファイル属性を指定する <see cref="ExtendedFileAttributes"/> 定数。</param>
      [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope")]
      [SecurityCritical]
      public BackupFileStream(string path, FileMode mode, FileSystemRights access, FileShare share, ExtendedFileAttributes attributes) : this(File.CreateFileCore(null, false, path, attributes, null, mode, access, share, true, false, PathFormat.RelativePath), access)
      {
      }


      /// <summary>指定されたパス、作成モード、アクセス権、共有許可、追加のファイル属性、アクセス制御および監査セキュリティで <see cref="BackupFileStream"/> クラスの新しいインスタンスを初期化します。</summary>
      /// <param name="path">現在の <see cref="BackupFileStream"/> オブジェクトがカプセル化するファイルの相対パスまたは絶対パス。</param>
      /// <param name="mode">ファイルを開くまたは作成する方法を決定する <see cref="FileMode"/> 定数。</param>
      /// <param name="access">ファイルのアクセスルールおよび監査ルールを作成する際に使用するアクセス権を決定する <see cref="FileSystemRights"/> 定数。</param>
      /// <param name="share">プロセス間でファイルを共有する方法を決定する <see cref="FileShare"/> 定数。</param>
      /// <param name="attributes">追加のファイル属性を指定する <see cref="ExtendedFileAttributes"/> 定数。</param>
      /// <param name="security">ファイルのアクセス制御および監査セキュリティを決定する <see cref="FileSecurity"/> 定数。このパラメーターは <c>null</c> にすることができます。</param>
      [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope")]
      [SecurityCritical]
      public BackupFileStream(string path, FileMode mode, FileSystemRights access, FileShare share, ExtendedFileAttributes attributes, FileSecurity security) : this(File.CreateFileCore(null, false, path, attributes, security, mode, access, share, true, false, PathFormat.RelativePath), access)
      {
      }


      /// <summary>指定されたパスと作成モードで <see cref="BackupFileStream"/> クラスの新しいインスタンスを初期化します。</summary>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">現在の <see cref="BackupFileStream"/> オブジェクトがカプセル化するファイルの相対パスまたは絶対パス。</param>
      /// <param name="mode">ファイルを開くまたは作成する方法を決定する <see cref="FileMode"/> 定数。</param>
      /// <remarks>ファイルは読み取りと書き込みの両方で排他アクセスとして開かれます。</remarks>
      [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope")]
      [SecurityCritical]
      public BackupFileStream(KernelTransaction transaction, string path, FileMode mode) : this(File.CreateFileCore(transaction, false, path, ExtendedFileAttributes.Normal, null, mode, FileSystemRights.Read | FileSystemRights.Write, FileShare.None, true, false, PathFormat.RelativePath), FileSystemRights.Read | FileSystemRights.Write)
      {
      }


      /// <summary>指定されたパス、作成モード、およびアクセス権で <see cref="BackupFileStream"/> クラスの新しいインスタンスを初期化します。</summary>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">現在の <see cref="BackupFileStream"/> オブジェクトがカプセル化するファイルの相対パスまたは絶対パス。</param>
      /// <param name="mode">ファイルを開くまたは作成する方法を決定する <see cref="FileMode"/> 定数。</param>
      /// <param name="access">ファイルのアクセスルールおよび監査ルールを作成する際に使用するアクセス権を決定する <see cref="FileSystemRights"/> 定数。</param>
      /// <remarks>ファイルは排他アクセスとして開かれます。</remarks>
      [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope")]
      [SecurityCritical]
      public BackupFileStream(KernelTransaction transaction, string path, FileMode mode, FileSystemRights access) : this(File.CreateFileCore(transaction, false, path, ExtendedFileAttributes.Normal, null, mode, access, FileShare.None, true, false, PathFormat.RelativePath), access)
      {
      }


      /// <summary>指定されたパス、作成モード、アクセス権、および共有許可で <see cref="BackupFileStream"/> クラスの新しいインスタンスを初期化します。</summary>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">現在の <see cref="BackupFileStream"/> オブジェクトがカプセル化するファイルの相対パスまたは絶対パス。</param>
      /// <param name="mode">ファイルを開くまたは作成する方法を決定する <see cref="FileMode"/> 定数。</param>
      /// <param name="access">ファイルのアクセスルールおよび監査ルールを作成する際に使用するアクセス権を決定する <see cref="FileSystemRights"/> 定数。</param>
      /// <param name="share">プロセス間でファイルを共有する方法を決定する <see cref="FileShare"/> 定数。</param>
      [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope")]
      [SecurityCritical]
      public BackupFileStream(KernelTransaction transaction, string path, FileMode mode, FileSystemRights access, FileShare share) : this(File.CreateFileCore(transaction, false, path, ExtendedFileAttributes.Normal, null, mode, access, share, true, false, PathFormat.RelativePath), access)
      {
      }


      /// <summary>指定されたパス、作成モード、アクセス権、共有許可、および追加のファイル属性で <see cref="BackupFileStream"/> クラスの新しいインスタンスを初期化します。</summary>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">現在の <see cref="BackupFileStream"/> オブジェクトがカプセル化するファイルの相対パスまたは絶対パス。</param>
      /// <param name="mode">ファイルを開くまたは作成する方法を決定する <see cref="FileMode"/> 定数。</param>
      /// <param name="access">ファイルのアクセスルールおよび監査ルールを作成する際に使用するアクセス権を決定する <see cref="FileSystemRights"/> 定数。</param>
      /// <param name="share">プロセス間でファイルを共有する方法を決定する <see cref="FileShare"/> 定数。</param>
      /// <param name="attributes">追加のファイル属性を指定する <see cref="ExtendedFileAttributes"/> 定数。</param>
      [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope")]
      [SecurityCritical]
      public BackupFileStream(KernelTransaction transaction, string path, FileMode mode, FileSystemRights access, FileShare share, ExtendedFileAttributes attributes) : this(File.CreateFileCore(transaction, false, path, attributes, null, mode, access, share, true, false, PathFormat.RelativePath), access)
      {
      }


      /// <summary>指定されたパス、作成モード、アクセス権、共有許可、追加のファイル属性、アクセス制御および監査セキュリティで <see cref="BackupFileStream"/> クラスの新しいインスタンスを初期化します。</summary>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">現在の <see cref="BackupFileStream"/> オブジェクトがカプセル化するファイルの相対パスまたは絶対パス。</param>
      /// <param name="mode">ファイルを開くまたは作成する方法を決定する <see cref="FileMode"/> 定数。</param>
      /// <param name="access">ファイルのアクセスルールおよび監査ルールを作成する際に使用するアクセス権を決定する <see cref="FileSystemRights"/> 定数。</param>
      /// <param name="share">プロセス間でファイルを共有する方法を決定する <see cref="FileShare"/> 定数。</param>
      /// <param name="attributes">追加のファイル属性を指定する <see cref="ExtendedFileAttributes"/> 定数。</param>
      /// <param name="security">ファイルのアクセス制御および監査セキュリティを決定する <see cref="FileSecurity"/> 定数。このパラメーターは <c>null</c> にすることができます。</param>
      [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope")]
      [SecurityCritical]
      public BackupFileStream(KernelTransaction transaction, string path, FileMode mode, FileSystemRights access, FileShare share, ExtendedFileAttributes attributes, FileSecurity security) : this(File.CreateFileCore(transaction, false, path, attributes, security, mode, access, share, true, false, PathFormat.RelativePath), access)
      {
      }


      /// <summary>指定されたファイルハンドルと読み取り/書き込み許可で <see cref="BackupFileStream"/> クラスの新しいインスタンスを初期化します。</summary>
      /// <param name="handle">この <see cref="BackupFileStream"/> オブジェクトがカプセル化するファイルのファイルハンドル。</param>
      /// <param name="access"><see cref="BackupFileStream"/> オブジェクトの <see cref="CanRead"/> および <see cref="CanWrite"/> プロパティを取得する <see cref="FileSystemRights"/> 定数。</param>
      [SecurityCritical]
      public BackupFileStream(SafeFileHandle handle, FileSystemRights access)
      {
         NativeMethods.IsValidHandle(handle);

         SafeFileHandle = handle;

         _canRead = (access & FileSystemRights.ReadData) != 0;
         _canWrite = (access & FileSystemRights.WriteData) != 0;
         _processSecurity = true;
      }


      /// <summary>派生クラスでオーバーライドされた場合、ストリームの長さをバイト単位で取得します。</summary>
      /// <value>このメソッドは常に例外をスローします。</value>
      /// <exception cref="NotSupportedException"/>
      public override long Length
      {
         get { throw new NotSupportedException(Resources.No_Stream_Seeking_Support); }
      }


      /// <summary>派生クラスでオーバーライドされた場合、現在のストリーム内の位置を取得または設定します。</summary>
      /// <value>このメソッドは常に例外をスローします。</value>
      /// <exception cref="NotSupportedException"/>
      public override long Position
      {
         get { throw new NotSupportedException(Resources.No_Stream_Seeking_Support); }
         set { throw new NotSupportedException(Resources.No_Stream_Seeking_Support); }
      }


      /// <summary>派生クラスでオーバーライドされた場合、現在のストリーム内の位置を設定します。</summary>
      /// <param name="offset"><paramref name="origin"/> パラメーターに対する相対的なバイトオフセット。</param>
      /// <param name="origin">新しい位置を取得するために使用される参照ポイントを示す <see cref="System.IO.SeekOrigin"/> 型の値。</param>
      /// <returns>現在のストリーム内の新しい位置。</returns>
      /// <remarks><para><note><para>このストリームはこのメソッドによるシークをサポートしていません。このメソッドを呼び出すと常に <see cref="NotSupportedException"/> がスローされます。前方へのシークの代替方法については <see cref="Skip"/> を参照してください。</para></note></para></remarks>
      /// <exception cref="NotSupportedException"/>
      public override long Seek(long offset, SeekOrigin origin)
      {
         throw new NotSupportedException(Resources.No_Stream_Seeking_Support);
      }


      /// <summary>派生クラスでオーバーライドされた場合、現在のストリームの長さを設定します。</summary>
      /// <param name="value">現在のストリームの希望する長さ（バイト単位）。</param>
      /// <remarks>このメソッドは <see cref="BackupFileStream"/> クラスではサポートされておらず、呼び出すと常に <see cref="NotSupportedException"/> が生成されます。</remarks>
      /// <exception cref="NotSupportedException"/>
      public override void SetLength(long value)
      {
         throw new NotSupportedException(Resources.No_Stream_Seeking_Support);
      }
      
      


      /// <summary>現在のストリームが読み取りをサポートするかどうかを示す値を取得します。</summary>
      /// <returns>ストリームが読み取りをサポートする場合は <c>true</c>、それ以外の場合は <c>false</c>。</returns>
      public override bool CanRead
      {
         get { return _canRead; }
      }

      /// <summary>現在のストリームがシークをサポートするかどうかを示す値を取得します。</summary>
      /// <returns>このメソッドは常に <c>false</c> を返します。</returns>
      public override bool CanSeek
      {
         get { return false; }
      }

      /// <summary>現在のストリームが書き込みをサポートするかどうかを示す値を取得します。</summary>
      /// <returns>ストリームが書き込みをサポートする場合は <c>true</c>、それ以外の場合は <c>false</c>。</returns>
      public override bool CanWrite
      {
         get { return _canWrite; }
      }

      /// <summary>現在の <see cref="BackupFileStream"/> オブジェクトがカプセル化するファイルのオペレーティングシステムファイルハンドルを表す <see cref="SafeFileHandle"/> オブジェクトを取得します。</summary>
      /// <value>現在の <see cref="BackupFileStream"/> オブジェクトがカプセル化するファイルのオペレーティングシステムファイルハンドルを表す
      /// <see cref="SafeFileHandle"/> オブジェクト。</value>
      private SafeFileHandle SafeFileHandle { get; set; }




      /// <summary>現在のストリームからバイトシーケンスを読み取り、読み取ったバイト数だけストリーム内の位置を進めます。</summary>
      /// <remarks>このメソッドはファイルまたはディレクトリのアクセス制御リスト (ACL) データをバックアップしません。</remarks>
      /// <param name="buffer">
      ///   バイト配列。このメソッドが戻ると、バッファーには指定されたバイト配列の
      ///   <paramref name="offset"/> から (<paramref name="offset"/> + <paramref name="count"/> - 1) までの値が、
      ///   現在のソースから読み取ったバイトで置き換えられた状態で格納されます。
      /// </param>
      /// <param name="offset">
      ///   現在のストリームから読み取ったデータの格納を開始する、<paramref name="buffer"/> 内のゼロベースのバイトオフセット。
      /// </param>
      /// <param name="count">現在のストリームから読み取る最大バイト数。</param>
      /// <returns>
      ///   バッファーに読み取られたバイトの合計数。要求されたバイト数が現在利用できない場合は要求数より少ない値、
      ///   ストリームの末尾に達した場合はゼロ (0) になります。
      /// </returns>
      ///
      /// <exception cref="System.ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="System.ArgumentOutOfRangeException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="ObjectDisposedException"/>
      public override int Read(byte[] buffer, int offset, int count)
      {
         return Read(buffer, offset, count, false);
      }


      /// <summary>派生クラスでオーバーライドされた場合、現在のストリームからバイトシーケンスを読み取り、読み取ったバイト数だけストリーム内の位置を進めます。</summary>
      /// <param name="buffer">バイト配列。このメソッドが戻ると、バッファーには指定されたバイト配列の
      /// <paramref name="offset"/> から (<paramref name="offset"/> + <paramref name="count"/> - 1) までの値が、現在のソースから読み取ったバイトで置き換えられた状態で格納されます。</param>
      /// <param name="offset">現在のストリームから読み取ったデータの格納を開始する、<paramref name="buffer"/> 内のゼロベースのバイトオフセット。</param>
      /// <param name="count">現在のストリームから読み取る最大バイト数。</param>
      /// <param name="processSecurity">関数がファイルまたはディレクトリのアクセス制御リスト (ACL) データをバックアップするかどうかを示します。</param>
      /// <returns>
      /// バッファーに読み取られたバイトの合計数。要求されたバイト数が現在利用できない場合は要求数より少ない値、
      /// ストリームの末尾に達した場合はゼロ (0) になります。
      /// </returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="ArgumentOutOfRangeException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="ObjectDisposedException"/>
      [SuppressMessage("Microsoft.Security", "CA2122:DoNotIndirectlyExposeMethodsWithLinkDemands")]
      [SecurityCritical]
      public int Read(byte[] buffer, int offset, int count, bool processSecurity)
      {
         if (buffer == null)
         {
            throw new ArgumentNullException("buffer");
         }

         if (!CanRead)
         {
            throw new NotSupportedException("Stream does not support reading");
         }

         if (offset + count > buffer.Length)
         {
            throw new ArgumentException("The sum of offset and count is larger than the size of the buffer.", "offset");
         }

         if (offset < 0)
         {
            throw new ArgumentOutOfRangeException("offset", offset, Resources.Negative_Offset);
         }

         if (count < 0)
         {
            throw new ArgumentOutOfRangeException("count", count, Resources.Negative_Count);
         }


         using var safeBuffer = new SafeGlobalMemoryBufferHandle(count);

         var success = NativeMethods.BackupRead(SafeFileHandle, safeBuffer, (uint) safeBuffer.Capacity, out var numberOfBytesRead, false, processSecurity, ref _context);
            
         var lastError = Marshal.GetLastWin32Error();
         if (!success)
         {
            NativeError.ThrowException(lastError);
         }


         // File.GetAccessControlCore() 参照: .CopyTo() はそこでは動作しない？
         // 2017-06-13: .CopyTo() はここで何か有用なことをしているのか？
         safeBuffer.CopyTo(buffer, offset, count);

         return (int) numberOfBytesRead;
      }


      /// <summary>現在のストリームにバイトシーケンスを書き込み、書き込んだバイト数だけこのストリーム内の現在位置を進めます。</summary>
      /// <overloads>
      /// 現在のストリームにバイトシーケンスを書き込み、書き込んだバイト数だけこのストリーム内の現在位置を進めます。
      /// </overloads>
      /// <param name="buffer">バイト配列。このメソッドは <paramref name="buffer"/> から <paramref name="count"/> バイトを現在のストリームにコピーします。</param>
      /// <param name="offset">現在のストリームへのバイトのコピーを開始する、<paramref name="buffer"/> 内のゼロベースのバイトオフセット。</param>
      /// <param name="count">現在のストリームに書き込むバイト数。</param>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="System.ArgumentNullException"/>
      /// <exception cref="System.ArgumentOutOfRangeException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="ObjectDisposedException"/>
      /// <remarks>このメソッドはファイルまたはディレクトリのアクセス制御リスト (ACL) データを処理しません。</remarks>
      public override void Write(byte[] buffer, int offset, int count)
      {
         Write(buffer, offset, count, false);
      }


      /// <summary>派生クラスでオーバーライドされた場合、現在のストリームにバイトシーケンスを書き込み、書き込んだバイト数だけこのストリーム内の現在位置を進めます。</summary>
      /// <param name="buffer">バイト配列。このメソッドは <paramref name="buffer"/> から <paramref name="count"/> バイトを現在のストリームにコピーします。</param>
      /// <param name="offset">現在のストリームへのバイトのコピーを開始する、<paramref name="buffer"/> 内のゼロベースのバイトオフセット。</param>
      /// <param name="count">現在のストリームに書き込むバイト数。</param>
      /// <param name="processSecurity">関数がファイルまたはディレクトリのアクセス制御リスト (ACL) データを復元するかどうかを指定します。
      /// これが <c>true</c> の場合、ファイルまたはディレクトリハンドルを開く際に <see cref="FileSystemRights.TakeOwnership"/> および <see cref="FileSystemRights.ChangePermissions"/> アクセスを指定する必要があります。
      /// ハンドルにこれらのアクセス権がない場合、オペレーティングシステムは ACL データへのアクセスを拒否し、ACL データの復元は行われません。</param>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="System.ArgumentOutOfRangeException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="ObjectDisposedException"/>
      [SuppressMessage("Microsoft.Security", "CA2122:DoNotIndirectlyExposeMethodsWithLinkDemands")]
      [SecurityCritical]
      public void Write(byte[] buffer, int offset, int count, bool processSecurity)
      {
         if (buffer == null)
         {
            throw new ArgumentNullException("buffer");
         }

         if (offset < 0)
         {
            throw new ArgumentOutOfRangeException("offset", offset, Resources.Negative_Offset);
         }

         if (count < 0)
         {
            throw new ArgumentOutOfRangeException("count", count, Resources.Negative_Count);
         }

         if (offset + count > buffer.Length)
         {
            throw new ArgumentException(Resources.Buffer_Not_Large_Enough, "offset");
         }


         using var safeBuffer = new SafeGlobalMemoryBufferHandle(count);
         safeBuffer.CopyFrom(buffer, offset, count);

         uint bytesWritten;

         var success = NativeMethods.BackupWrite(SafeFileHandle, safeBuffer, (uint)safeBuffer.Capacity, out bytesWritten, false, processSecurity, ref _context);
            
         var lastError = Marshal.GetLastWin32Error();
         if (!success)
         {
            NativeError.ThrowException(lastError);
         }
      }


      /// <summary>このストリームのすべてのバッファーをクリアし、バッファーされたデータを基になるデバイスに書き込みます。</summary>
      public override void Flush()
      {
         var success = NativeMethods.FlushFileBuffers(SafeFileHandle);

         var lastError = Marshal.GetLastWin32Error();
         if (!success)
         {
            NativeError.ThrowException(lastError);
         }
      }


      /// <summary>現在のストリームから指定されたバイト数だけ前方にスキップします。</summary>
      /// <remarks><para>このメソッドは Win32 API の <see href="http://msdn.microsoft.com/en-us/library/aa362509(VS.85).aspx">BackupSeek</see> の実装です。</para>
      /// <para>
      /// アプリケーションは <see cref="Skip"/> メソッドを使用して、エラーの原因となるデータストリームの部分をスキップします。この関数はストリームヘッダーを
      /// 越えてシークしません。例えば、この関数はストリーム名をスキップするために使用できません。アプリケーションがサブストリームの末尾を越えて
      /// シークしようとすると関数は失敗し、戻り値は関数が実際にシークしたバイト数を示し、ファイル位置は次のストリームヘッダーの先頭に配置されます。
      /// </para>
      /// </remarks>
      /// <param name="bytes">スキップするバイト数。</param>
      /// <returns>実際にスキップされたバイト数。</returns>
      [SecurityCritical]
      public long Skip(long bytes)
      {

         var success = NativeMethods.BackupSeek(SafeFileHandle, NativeMethods.GetLowOrderDword(bytes), NativeMethods.GetHighOrderDword(bytes), out var lowSought, out var highSought, ref _context);

         var lastError = Marshal.GetLastWin32Error();
         if (!success && lastError != Win32Errors.ERROR_SEEK)
         {
            // エラーコード25はシークエラーを示すため、ここではスキップする。
            NativeError.ThrowException(lastError);
         }

         return NativeMethods.ToLong(highSought, lowSought);
      }


      /// <summary>現在の <see cref="BackupFileStream"/> オブジェクトが記述するファイルのアクセス制御リスト (ACL) エントリをカプセル化する <see cref="FileSecurity"/> オブジェクトを取得します。</summary>
      /// <exception cref="IOException"/>
      /// <returns>
      ///   現在の <see cref="BackupFileStream"/> オブジェクトが記述するファイルのアクセス制御リスト (ACL) エントリをカプセル化する
      ///   <see cref="FileSecurity"/> オブジェクト。
      /// </returns>
      [SuppressMessage("Microsoft.Design", "CA1024:UsePropertiesWhereAppropriate")]
      [SecurityCritical]
      public FileSecurity GetAccessControl()
      {
         IntPtr pSidOwner, pSidGroup, pDacl, pSacl;

         var lastError = (int) SecurityNativeMethods.GetSecurityInfo(SafeFileHandle, SE_OBJECT_TYPE.SE_FILE_OBJECT, SECURITY_INFORMATION.GROUP_SECURITY_INFORMATION | SECURITY_INFORMATION.OWNER_SECURITY_INFORMATION | SECURITY_INFORMATION.LABEL_SECURITY_INFORMATION | SECURITY_INFORMATION.DACL_SECURITY_INFORMATION | SECURITY_INFORMATION.SACL_SECURITY_INFORMATION, out pSidOwner, out pSidGroup, out pDacl, out pSacl, out var pSecurityDescriptor);

         try
         {
            if (lastError != Win32Errors.ERROR_SUCCESS)
            {
               NativeError.ThrowException(lastError);
            }

            if (null != pSecurityDescriptor && pSecurityDescriptor.IsInvalid)
            {
               pSecurityDescriptor.Close();
               pSecurityDescriptor = null;

               throw new IOException(new Win32Exception((int) Win32Errors.ERROR_INVALID_SECURITY_DESCR).Message);
            }


            var length = SecurityNativeMethods.GetSecurityDescriptorLength(pSecurityDescriptor);
            var managedBuffer = new byte[length];

            
            // .CopyTo() はそこでは動作しない？
            if (null != pSecurityDescriptor)
            {
               pSecurityDescriptor.CopyTo(managedBuffer, 0, (int) length);
            }


            var fs = new FileSecurity();
            fs.SetSecurityDescriptorBinaryForm(managedBuffer);

            return fs;
         }
         finally
         {
            if (null != pSecurityDescriptor)
            {
               pSecurityDescriptor.Close();
            }
         }
      }


      /// <summary><see cref="FileSecurity"/> オブジェクトで記述されたアクセス制御リスト (ACL) エントリを、現在の <see cref="BackupFileStream"/> オブジェクトが記述するファイルに適用します。</summary>
      /// <param name="fileSecurity">現在のファイルに適用する ACL エントリを記述する <see cref="FileSecurity"/> オブジェクト。</param>
      [SecurityCritical]
      public void SetAccessControl(ObjectSecurity fileSecurity)
      {
         File.SetAccessControlCore(null, SafeFileHandle, fileSecurity, AccessControlSections.All, PathFormat.LongFullPath);
      }


      /// <summary>読み取りアクセスを許可しながら、他のプロセスが <see cref="BackupFileStream"/> を変更することを防止します。</summary>
      /// <param name="position">ロックする範囲の先頭。このパラメーターの値はゼロ (0) 以上でなければなりません。</param>
      /// <param name="length">ロックする範囲。</param>
      /// <exception cref="ArgumentOutOfRangeException"/>
      /// <exception cref="ObjectDisposedException"/>
      [SecurityCritical]
      public void Lock(long position, long length)
      {
         if (position < 0)
         {
            throw new ArgumentOutOfRangeException("position", position, new Win32Exception((int) Win32Errors.ERROR_NEGATIVE_SEEK).Message);
         }

         if (length < 0)
         {
            throw new ArgumentOutOfRangeException("length", length, Resources.Negative_Lock_Length);
         }


         var success = NativeMethods.LockFile(SafeFileHandle, NativeMethods.GetLowOrderDword(position), NativeMethods.GetHighOrderDword(position), NativeMethods.GetLowOrderDword(length), NativeMethods.GetHighOrderDword(length));

         var lastError = Marshal.GetLastWin32Error();
         if (!success)
         {
            NativeError.ThrowException(lastError);
         }
      }


      /// <summary>以前にロックされたファイルの全体または一部への他のプロセスによるアクセスを許可します。</summary>
      /// <param name="position">ロック解除する範囲の先頭。</param>
      /// <param name="length">ロック解除する範囲。</param>
      /// <exception cref="ArgumentOutOfRangeException"/>
      /// <exception cref="ArgumentOutOfRangeException"/>
      /// <exception cref="ObjectDisposedException"/>
      [SecurityCritical]
      public void Unlock(long position, long length)
      {
         if (position < 0)
         {
            throw new ArgumentOutOfRangeException("position", position, new Win32Exception((int) Win32Errors.ERROR_NEGATIVE_SEEK).Message);
         }

         if (length < 0)
         {
            throw new ArgumentOutOfRangeException("length", length, Resources.Negative_Lock_Length);
         }


         var success = NativeMethods.UnlockFile(SafeFileHandle, NativeMethods.GetLowOrderDword(position), NativeMethods.GetHighOrderDword(position), NativeMethods.GetLowOrderDword(length), NativeMethods.GetHighOrderDword(length));

         var lastError = Marshal.GetLastWin32Error();
         if (!success)
         {
            NativeError.ThrowException(lastError);
         }
      }


      /// <summary>現在の <see cref="BackupFileStream"/> からストリームヘッダーを読み取ります。</summary>
      /// <returns>現在の <see cref="BackupFileStream"/> から読み取ったストリームヘッダー。ヘッダーの必要バイト数を読み取る前に
      /// ファイル終端に達した場合は <c>null</c>。</returns>
      /// <exception cref="IOException"/>
      /// <remarks>返されるオブジェクトが有効な情報を表すためには、ストリームが実際のヘッダーが開始する位置に
      /// 配置されている必要があります。</remarks>
      [SecurityCritical]
      public BackupStreamInfo ReadStreamInfo()
      {
         var sizeOf = Marshal.SizeOf<NativeMethods.WIN32_STREAM_ID>();

         using var hBuf = new SafeGlobalMemoryBufferHandle(sizeOf);

         var success = NativeMethods.BackupRead(SafeFileHandle, hBuf, (uint) sizeOf, out var numberOfBytesRead, false, _processSecurity, ref _context);

         var lastError = Marshal.GetLastWin32Error();
         if (!success)
         {
            NativeError.ThrowException(lastError);
         }


         if (numberOfBytesRead == 0)
         {
            return null;
         }


         if (numberOfBytesRead < sizeOf)
         {
            throw new IOException(Resources.Read_Incomplete_Header);
         }


         var streamID = hBuf.PtrToStructure<NativeMethods.WIN32_STREAM_ID>(0);
         var nameLength = (uint) Math.Min(streamID.dwStreamNameSize, hBuf.Capacity);


         success = NativeMethods.BackupRead(SafeFileHandle, hBuf, nameLength, out numberOfBytesRead, false, _processSecurity, ref _context);

         lastError = Marshal.GetLastWin32Error();
         if (!success)
         {
            NativeError.ThrowException(lastError);
         }


         var name = hBuf.PtrToStringUni(0, (int) nameLength / UnicodeEncoding.CharSize);

         return new BackupStreamInfo(streamID, name);
      }




      #region Disposable Members

      /// <summary><see cref="BackupFileStream"/> がガベージコレクションによって回収される前に、アンマネージリソースを解放し、その他のクリーンアップ操作を実行します。</summary>
      ~BackupFileStream()
      {
         Dispose(false);
      }


      /// <summary><see cref="System.IO.Stream"/> で使用されるアンマネージリソースを解放し、オプションでマネージリソースも解放します。</summary>
      /// <param name="disposing">マネージリソースとアンマネージリソースの両方を解放する場合は <c>true</c>、アンマネージリソースのみを解放する場合は <c>false</c>。</param>
      [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope")]
      protected override void Dispose(bool disposing)
      {
         if (disposing)
         {
            // コンストラクタのいずれかが以前に例外をスローした場合、
            // オブジェクトは正しく初期化されておらず、ファイナライザーからの呼び出しは失敗する。

            if (null != SafeFileHandle && !SafeFileHandle.IsInvalid)
            {
               if (_context != IntPtr.Zero)
               {
                  try
                  {
                     uint temp;

                     // MSDN: データ構造が使用するメモリを解放するには、バックアップ操作完了時に bAbort パラメーターを TRUE に設定して BackupRead を呼び出す。
                     var success = NativeMethods.BackupRead(SafeFileHandle, new SafeGlobalMemoryBufferHandle(), 0, out temp, true, false, ref _context);

                     var lastError = Marshal.GetLastWin32Error();
                     if (!success)
                     {
                        NativeError.ThrowException(lastError);
                     }
                  }
                  finally
                  {
                     _context = IntPtr.Zero;

                     NativeMethods.CloseSafeHandle(SafeFileHandle);
                  }
               }
            }
         }


         base.Dispose(disposing);
      }
      
      #endregion // Disposable Members
   }
}
