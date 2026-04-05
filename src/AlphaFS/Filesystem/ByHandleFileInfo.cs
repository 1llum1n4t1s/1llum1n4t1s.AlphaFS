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
   /// <summary>GetFileInformationByHandle関数が取得する情報を格納します。</summary>
   [Serializable]
   [SecurityCritical]
   public sealed class ByHandleFileInfo
   {
      internal ByHandleFileInfo(NativeMethods.BY_HANDLE_FILE_INFORMATION fibh)
      {
         CreationTimeUtc = DateTime.FromFileTimeUtc(fibh.ftCreationTime);
         LastAccessTimeUtc = DateTime.FromFileTimeUtc(fibh.ftLastAccessTime);
         LastWriteTimeUtc = DateTime.FromFileTimeUtc(fibh.ftLastWriteTime);

         Attributes = fibh.dwFileAttributes;
         FileIndex = NativeMethods.ToLong(fibh.nFileIndexHigh, fibh.nFileIndexLow);
         FileSize = NativeMethods.ToLong(fibh.nFileSizeHigh, fibh.nFileSizeLow);
         NumberOfLinks = (int) fibh.nNumberOfLinks;
         VolumeSerialNumber = fibh.dwVolumeSerialNumber;
      }


      /// <summary>ファイル属性を取得します。</summary>
      /// <value>ファイル属性。</value>
      public FileAttributes Attributes { get; private set; }


      /// <summary>このエントリが作成された時刻を取得します。</summary>
      /// <value>このエントリが作成された時刻。</value>
      public DateTime CreationTime
      {
         get { return CreationTimeUtc.ToLocalTime(); }
      }


      /// <summary>このエントリが作成された協定世界時（UTC）での時刻を取得します。</summary>
      /// <value>このエントリが作成されたUTC時刻。</value>
      public DateTime CreationTimeUtc { get; private set; }


      /// <summary>このエントリに最後にアクセスした時刻を取得します。
      /// ファイルの場合、ファイルが最後に読み取りまたは書き込みされた時刻を示します。
      /// ディレクトリの場合、ディレクトリが作成された時刻を示します。
      /// ファイルとディレクトリの両方で、日付は正しいですが、時刻は常に午前0時に設定されます。
      /// 基盤のファイルシステムが最終アクセス時刻をサポートしない場合、このメンバーはゼロ（0）です。
      /// </summary>
      /// <value>このエントリに最後にアクセスした時刻。</value>
      public DateTime LastAccessTime
      {
         get { return LastAccessTimeUtc.ToLocalTime(); }
      }


      /// <summary>このエントリに最後にアクセスした協定世界時（UTC）での時刻を取得します。
      /// ファイルの場合、ファイルが最後に読み取りまたは書き込みされた時刻を示します。
      /// ディレクトリの場合、ディレクトリが作成された時刻を示します。
      /// ファイルとディレクトリの両方で、日付は正しいですが、時刻は常に午前0時に設定されます。
      /// 基盤のファイルシステムが最終アクセス時刻をサポートしない場合、このメンバーはゼロ（0）です。
      /// </summary>
      /// <value>このエントリに最後にアクセスしたUTC時刻。</value>
      public DateTime LastAccessTimeUtc { get; private set; }


      /// <summary>このエントリが最後に変更された時刻を取得します。
      /// ファイルの場合、ファイルが最後に書き込まれた時刻を示します。
      /// ディレクトリの場合、ディレクトリが作成された時刻を示します。
      /// 基盤のファイルシステムが最終アクセス時刻をサポートしない場合、このメンバーはゼロ（0）です。
      /// </summary>
      /// <value>このエントリが最後に変更された時刻。</value>
      public DateTime LastWriteTime
      {
         get { return LastWriteTimeUtc.ToLocalTime(); }
      }


      /// <summary>このエントリが最後に変更された協定世界時（UTC）での時刻を取得します。
      /// ファイルの場合、ファイルが最後に書き込まれた時刻を示します。
      /// ディレクトリの場合、ディレクトリが作成された時刻を示します。
      /// 基盤のファイルシステムが最終アクセス時刻をサポートしない場合、このメンバーはゼロ（0）です。
      /// </summary>
      /// <value>このエントリが最後に変更されたUTC時刻。</value>
      public DateTime LastWriteTimeUtc { get; private set; }


      /// <summary>ファイルが含まれるボリュームのシリアル番号を取得します。</summary>
      /// <value>ファイルが含まれるボリュームのシリアル番号。</value>
      public long VolumeSerialNumber { get; private set; }


      /// <summary>ファイルサイズを取得します。</summary>
      /// <value>ファイルサイズ。</value>
      public long FileSize { get; private set; }


      /// <summary>このファイルへのリンク数を取得します。FATファイルシステムではこのメンバーは常に1です。NTFSファイルシステムでは1より大きい場合があります。</summary>
      /// <value>このファイルへのリンク数。</value>
      public int NumberOfLinks { get; private set; }


      /// <summary>
      /// ファイルに関連付けられた一意の識別子を取得します。この識別子とボリュームシリアル番号の組み合わせにより、
      /// 単一のコンピューター上でファイルを一意に識別できます。2つのオープンハンドルが同じファイルを表すかどうかを
      /// 判断するには、各ファイルの識別子とボリュームシリアル番号を組み合わせて比較します。
      /// </summary>
      /// <value>ファイルの一意の識別子。</value>
      public long FileIndex { get; private set; }
   }
}
