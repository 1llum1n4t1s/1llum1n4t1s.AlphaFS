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

using Microsoft.Win32.SafeHandles;
using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;

namespace Alphaleonis.Win32.Filesystem
{
   /// <summary>ファイルシステムボリュームに関する情報を含みます。</summary>
   [Serializable]
   [SecurityCritical]
   public sealed class VolumeInfo
   {
      [NonSerialized] private readonly bool _continueOnAccessError;
      [NonSerialized] private readonly SafeFileHandle _volumeHandle;
      [NonSerialized] private NativeMethods.VOLUME_INFO_FLAGS _volumeInfoAttributes;


      /// <summary>VolumeInfo インスタンスを初期化します。</summary>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="ArgumentException"/>
      /// <param name="volumeName">有効なドライブパスまたはドライブ文字。大文字または小文字の 'a' ～ 'z'、または \\server\share 形式のネットワーク共有を指定できます。</param>
      [SuppressMessage("Microsoft.Design", "CA1062:Validate arguments of public methods", MessageId = "0", Justification = "Utils.IsNullOrWhiteSpace validates arguments.")]
      [SecurityCritical]
      public VolumeInfo(string volumeName)
      {
         if (Utils.IsNullOrWhiteSpace(volumeName))
         {
            throw new ArgumentNullException("volumeName");
         }


         if (!volumeName.StartsWith(Path.LongPathPrefix, StringComparison.Ordinal))
         {
            volumeName = Path.IsUncPathCore(volumeName, false, false) ? Path.GetLongPathCore(volumeName, GetFullPathOptions.None) : Path.LongPathPrefix + volumeName;
         }

         else
         {
            volumeName = volumeName.Length == 1 ? volumeName + Path.VolumeSeparatorChar : Path.GetPathRoot(volumeName, false);

            if (!volumeName.StartsWith(Path.GlobalRootPrefix, StringComparison.OrdinalIgnoreCase))
            {
               volumeName = Path.GetPathRoot(volumeName, false);
            }
         }


         if (Utils.IsNullOrWhiteSpace(volumeName))
         {
            throw new ArgumentException(Resources.InvalidDriveLetterArgument, "volumeName");
         }


         Name = Path.AddTrailingDirectorySeparator(volumeName, false);

         _volumeHandle = null;
      }


      /// <summary>VolumeInfo インスタンスを初期化します。</summary>
      /// <param name="driveName">有効なドライブパスまたはドライブ文字。大文字または小文字の 'a' ～ 'z'、または "\\server\share" 形式のネットワーク共有を指定できます。</param>
      /// <param name="refresh">オブジェクトの状態を更新します。</param>
      /// <param name="continueOnException"><c>true</c> はリソース不足などの失敗から発生する可能性のある例外を抑制します。</param>
      [SecurityCritical]
      public VolumeInfo(string driveName, bool refresh, bool continueOnException) : this(driveName)
      {
         _continueOnAccessError = continueOnException;

         if (refresh)
         {
            Refresh();
         }
      }


      /// <summary>VolumeInfo インスタンスを初期化します。</summary>
      /// <param name="volumeHandle"><see cref="SafeFileHandle"/> ハンドルのインスタンス。</param>
      [SecurityCritical]
      public VolumeInfo(SafeFileHandle volumeHandle)
      {
         _volumeHandle = volumeHandle;
      }


      /// <summary>VolumeInfo インスタンスを初期化します。</summary>
      /// <param name="volumeHandle"><see cref="SafeFileHandle"/> ハンドルのインスタンス。</param>
      /// <param name="refresh">オブジェクトの状態を更新します。</param>
      /// <param name="continueOnException"><c>true</c> はリソース不足などの失敗から発生する可能性のある例外を抑制します。</param>
      [SecurityCritical]
      public VolumeInfo(SafeFileHandle volumeHandle, bool refresh, bool continueOnException) : this(volumeHandle)
      {
         _continueOnAccessError = continueOnException;

         if (refresh)
         {
            Refresh();
         }
      }


      

      /// <summary>オブジェクトの状態を更新します。</summary>
      public void Refresh()
      {
         var volumeNameBuffer = new StringBuilder(NativeMethods.MaxPath + 1);
         var fileSystemNameBuffer = new StringBuilder(NativeMethods.MaxPath + 1);
         int maximumComponentLength;
         uint serialNumber;

         using (new NativeMethods.ChangeErrorMode(NativeMethods.ErrorMode.FailCriticalErrors))
         {
            // GetVolumeInformationXxx()
            // 2013-07-18: MSDN は LongPath の使用を確認していませんが、この関数��� Unicode バージョンが存在します。

            uint lastError;

            do
            {
               var success = null != _volumeHandle && NativeMethods.IsAtLeastWindowsVista

                  // GetVolumeInformationByHandle() / GetVolumeInformation()
                  // 2013-07-18: MSDN は LongPath の使用を確認していませんが、この関数の Unicode バージョンが存在します。

                  ? NativeMethods.GetVolumeInformationByHandle(_volumeHandle, volumeNameBuffer, (uint) volumeNameBuffer.Capacity, out serialNumber, out maximumComponentLength, out _volumeInfoAttributes, fileSystemNameBuffer, (uint) fileSystemNameBuffer.Capacity)

                  // 末尾のバックスラッシュが必要です。
                  : NativeMethods.GetVolumeInformation(Path.AddTrailingDirectorySeparator(Name, false), volumeNameBuffer, (uint) volumeNameBuffer.Capacity, out serialNumber, out maximumComponentLength, out _volumeInfoAttributes, fileSystemNameBuffer, (uint) fileSystemNameBuffer.Capacity);


               lastError = (uint) Marshal.GetLastWin32Error();
               if (!success)
               {
                  switch (lastError)
                  {
                     case Win32Errors.ERROR_NOT_READY:
                        if (!_continueOnAccessError)
                        {
                           throw new DeviceNotReadyException(Name, true);
                        }
                        break;

                     case Win32Errors.ERROR_MORE_DATA:
                        // 十分な大きさのバッファがあれば、このコードは実行されません。
                        volumeNameBuffer.Capacity = volumeNameBuffer.Capacity*2;
                        fileSystemNameBuffer.Capacity = fileSystemNameBuffer.Capacity*2;
                        break;

                     default:
                        if (!_continueOnAccessError)
                        {
                           NativeError.ThrowException(lastError, Name);
                        }
                        break;
                  }
               }

               else
               {
                  break;
               }

            } while (lastError == Win32Errors.ERROR_MORE_DATA);
         }

         FullPath = Path.GetRegularPathCore(Name, GetFullPathOptions.None, false);
         Name = volumeNameBuffer.ToString();

         FileSystemName = fileSystemNameBuffer.ToString();
         FileSystemName = !Utils.IsNullOrWhiteSpace(FileSystemName) ? FileSystemName : null;

         MaximumComponentLength = maximumComponentLength;
         SerialNumber = serialNumber;
      }


      /// <summary>ボリュームのフルパスを返します。</summary>
      /// <returns>このインスタンスを表す文字列。</returns>
      public override string ToString()
      {
         return Guid;
      }




      /// <summary>指定されたボリュームは、ディスクに名前を配置する際にファイル名の大文字小文字を保持します。</summary>
      public bool CasePreservedNames
      {
         get { return (_volumeInfoAttributes & NativeMethods.VOLUME_INFO_FLAGS.FILE_CASE_PRESERVED_NAMES) != 0; }
      }


      /// <summary>指定されたボリュームは大文字小文字を区別するファイル名をサポートします。</summary>
      public bool CaseSensitiveSearch
      {
         get { return (_volumeInfoAttributes & NativeMethods.VOLUME_INFO_FLAGS.FILE_CASE_SENSITIVE_SEARCH) != 0; }
      }


      /// <summary>指定されたボリュームはファイルベースの圧縮をサポートします。</summary>
      public bool Compression
      {
         get { return (_volumeInfoAttributes & NativeMethods.VOLUME_INFO_FLAGS.FILE_FILE_COMPRESSION) != 0; }
      }


      /// <summary>指定されたボリュームはダイレクトアクセス (DAX) ボリュームです。</summary>
      public bool DirectAccess
      {
         get { return (_volumeInfoAttributes & NativeMethods.VOLUME_INFO_FLAGS.FILE_DAX_VOLUME) != 0; }
      }


      /// <summary>ファイルシステムの名前を取得します。たとえば、FAT ファイルシステムや NTFS ファイルシステムなど。</summary>
      /// <value>ファイルシステムの名前。</value>
      public string FileSystemName { get; private set; }


      /// <summary>ボリュームへのフルパス。</summary>
      public string FullPath { get; private set; }


      private string _guid;
      /// <summary>ボリューム GUID。</summary>
      public string Guid
      {
         get
         {
            if (Utils.IsNullOrWhiteSpace(_guid))
            {
               _guid = !Utils.IsNullOrWhiteSpace(FullPath) ? Volume.GetUniqueVolumeNameForPath(FullPath) : null;
            }

            return _guid;
         }
      }


      /// <summary>ファイルシステムがサポートするファイル名コンポーネントの最大長を取得します。</summary>
      /// <value>ファイルシステムがサポートするファイル名コンポーネントの最大長。</value>      
      public int MaximumComponentLength { get; set; }


      /// <summary>ボリュームのラベルを取得します。</summary>
      /// <returns>ボリュームのラベル。</returns>
      /// <remarks>このプロパティは、"MyDrive" などのボリュームに割り当てられたラベルです。</remarks>
      public string Name { get; private set; }


      /// <summary>指定されたボリュームは名前付きストリームをサポートします。</summary>
      public bool NamedStreams
      {
         get { return (_volumeInfoAttributes & NativeMethods.VOLUME_INFO_FLAGS.FILE_NAMED_STREAMS) != 0; }
      }


      /// <summary>指定されたボリュームはアクセス制御リスト (ACL) を保持および適用します。</summary>
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Acls")]
      public bool PersistentAcls
      {
         get { return (_volumeInfoAttributes & NativeMethods.VOLUME_INFO_FLAGS.FILE_PERSISTENT_ACLS) != 0; }
      }

      
      /// <summary>指定されたボリュームは読み取り専用です。</summary>
      public bool ReadOnlyVolume
      {
         get { return (_volumeInfoAttributes & NativeMethods.VOLUME_INFO_FLAGS.FILE_READ_ONLY_VOLUME) != 0; }
      }

      
      /// <summary>指定されたボリュームは単一の順次書き込みをサポートします。</summary>
      public bool SequentialWriteOnce
      {
         get { return (_volumeInfoAttributes & NativeMethods.VOLUME_INFO_FLAGS.FILE_SEQUENTIAL_WRITE_ONCE) != 0; }
      }


      /// <summary>ハードディスクのフォーマット時にオペレーティングシステムが割り当てるボリュームシリアル番号を取得します。</summary>
      /// <value>ハードディスクのフォーマット時にオペレーティングシステムが割り当てるボリュームシリアル番号。</value>
      public long SerialNumber { get; private set; }


      /// <summary>指定されたボリュームは暗号化ファイルシステム (EFS) をサポートします。</summary>
      public bool SupportsEncryption
      {
         get { return (_volumeInfoAttributes & NativeMethods.VOLUME_INFO_FLAGS.FILE_SUPPORTS_ENCRYPTION) != 0; }
      }


      /// <summary>指定されたボリュームは拡張属性をサポートします。</summary>
      public bool SupportsExtendedAttributes
      {
         get { return (_volumeInfoAttributes & NativeMethods.VOLUME_INFO_FLAGS.FILE_SUPPORTS_EXTENDED_ATTRIBUTES) != 0; }
      }


      /// <summary>指定されたボリュームはハードリンクをサポートします。</summary>
      public bool SupportsHardLinks
      {
         get { return (_volumeInfoAttributes & NativeMethods.VOLUME_INFO_FLAGS.FILE_SUPPORTS_HARD_LINKS) != 0; }
      }


      /// <summary>指定されたボリュームはオブジェクト識別子をサポートします。</summary>
      public bool SupportsObjectIds
      {
         get { return (_volumeInfoAttributes & NativeMethods.VOLUME_INFO_FLAGS.FILE_SUPPORTS_OBJECT_IDS) != 0; }
      }


      /// <summary>ファイルシステムは FileID によるオープンをサポートします。</summary>
      public bool SupportsOpenByFileId
      {
         get { return (_volumeInfoAttributes & NativeMethods.VOLUME_INFO_FLAGS.FILE_SUPPORTS_OPEN_BY_FILE_ID) != 0; }
      }


      /// <summary>指定されたボリュームはリモートストレージをサポートします。（このプロパティは MSDN に記載されていません）</summary>
      public bool SupportsRemoteStorage
      {
         get { return (_volumeInfoAttributes & NativeMethods.VOLUME_INFO_FLAGS.FILE_SUPPORTS_REMOTE_STORAGE) != 0; }
      }


      /// <summary>指定されたボリュームはリパースポイントをサポートします。</summary>
      public bool SupportsReparsePoints
      {
         get { return (_volumeInfoAttributes & NativeMethods.VOLUME_INFO_FLAGS.FILE_SUPPORTS_REPARSE_POINTS) != 0; }
      }


      /// <summary>指定されたボリュームはスパースファイルをサポートします。</summary>
      public bool SupportsSparseFiles
      {
         get { return (_volumeInfoAttributes & NativeMethods.VOLUME_INFO_FLAGS.FILE_SUPPORTS_SPARSE_FILES) != 0; }
      }


      /// <summary>指定されたボリュームはトランザクションをサポートします。</summary>
      public bool SupportsTransactions
      {
         get { return (_volumeInfoAttributes & NativeMethods.VOLUME_INFO_FLAGS.FILE_SUPPORTS_TRANSACTIONS) != 0; }
      }


      /// <summary>指定されたボリュームは更新シーケンス番号 (USN) ジャーナルをサポートします。</summary>
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Usn")]
      public bool SupportsUsnJournal
      {
         get { return (_volumeInfoAttributes & NativeMethods.VOLUME_INFO_FLAGS.FILE_SUPPORTS_USN_JOURNAL) != 0; }
      }


      /// <summary>指定されたボリュームは、ディスク上のファイル名での Unicode をサポートします。</summary>
      public bool UnicodeOnDisk
      {
         get { return (_volumeInfoAttributes & NativeMethods.VOLUME_INFO_FLAGS.FILE_UNICODE_ON_DISK) != 0; }
      }


      /// <summary>指定されたボリュームは圧縮ボリュームです。たとえば、DoubleSpace ボリュームなど。</summary>
      public bool VolumeIsCompressed
      {
         get { return (_volumeInfoAttributes & NativeMethods.VOLUME_INFO_FLAGS.FILE_VOLUME_IS_COMPRESSED) != 0; }
      }


      /// <summary>指定されたボリュームはディスククォータをサポートします。</summary>
      public bool VolumeQuotas
      {
         get { return (_volumeInfoAttributes & NativeMethods.VOLUME_INFO_FLAGS.FILE_VOLUME_QUOTAS) != 0; }
      }
   }
}
