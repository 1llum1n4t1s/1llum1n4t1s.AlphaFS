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

namespace Alphaleonis.Win32.Filesystem
{
   internal static partial class NativeMethods
   {
      /// <summary>GetVolumeInfo() 関数で使用されるボリューム属性。</summary>
      [Flags]
      internal enum VOLUME_INFO_FLAGS
      {
         /// <summary>指定されたボリュームは大文字と小文字を区別するファイル名をサポートします。</summary>
         FILE_CASE_SENSITIVE_SEARCH = 1,


         /// <summary>指定されたボリュームは、ディスクに名前を配置する際にファイル名の大文字と小文字の保持をサポートします。</summary>
         FILE_CASE_PRESERVED_NAMES = 2,


         /// <summary>指定されたボリュームは、ディスク上に表示されるファイル名の Unicode をサポートします。</summary>
         FILE_UNICODE_ON_DISK = 4,


         /// <summary>指定されたボリュームはアクセス制御リスト (ACL) を保持および適用します。例えば、NTFS ファイルシステムは ACL を保持・適用しますが、FAT ファイルシステムはしません。</summary>
         FILE_PERSISTENT_ACLS = 8,


         /// <summary>指定されたボリュームはファイルベースの圧縮をサポートします。</summary>
         FILE_FILE_COMPRESSION = 16,


         /// <summary>指定されたボリュームはディスククォータをサポートします。</summary>
         FILE_VOLUME_QUOTAS = 32,


         /// <summary>指定されたボリュームはスパースファイルをサポートします。</summary>
         FILE_SUPPORTS_SPARSE_FILES = 64,


         /// <summary>指定されたボリュームはリパースポイントをサポートします。</summary>
         FILE_SUPPORTS_REPARSE_POINTS = 128,


         /// <summary>(MSDN に記載なし)</summary>
         FILE_SUPPORTS_REMOTE_STORAGE = 256,


         /// <summary>指定されたボリュームは圧縮ボリュームです（例: DoubleSpace ボリューム）。</summary>
         FILE_VOLUME_IS_COMPRESSED = 32768,


         /// <summary>指定されたボリュームはオブジェクト識別子をサポートします。</summary>
         FILE_SUPPORTS_OBJECT_IDS = 65536,


         /// <summary>指定されたボリュームは暗号化ファイルシステム (EFS) をサポートします。詳細については、ファイルの暗号化を参照してください。</summary>
         FILE_SUPPORTS_ENCRYPTION = 131072,


         /// <summary>指定されたボリュームは名前付きストリームをサポートします。</summary>
         FILE_NAMED_STREAMS = 262144,


         /// <summary>指定されたボリュームは読み取り専用です。</summary>
         FILE_READ_ONLY_VOLUME = 524288,


         /// <summary>指定されたボリュームは一度だけのシーケンシャル書き込みをサポートします。</summary>
         FILE_SEQUENTIAL_WRITE_ONCE = 1048576,


         /// <summary>指定されたボリュームはトランザクションをサポートします。詳細については、About KTM を参照してください。</summary>
         FILE_SUPPORTS_TRANSACTIONS = 2097152,


         /// <summary>指定されたボリュームはハードリンクをサポートします。詳細については、Hard Links and Junctions を参照してください。</summary>
         /// <remarks>Windows Server 2008、Windows Vista、Windows Server 2003、および Windows XP: この値は Windows Server 2008 R2 および Windows 7 までサポートされていません。</remarks>
         FILE_SUPPORTS_HARD_LINKS = 4194304,


         /// <summary>指定されたボリュームは拡張属性をサポートします。拡張属性はアプリケーション固有のメタデータで、アプリケーションがファイルに関連付けることができ、ファイルのデータの一部ではありません。</summary>
         /// <remarks>Windows Server 2008、Windows Vista、Windows Server 2003、および Windows XP: この値は Windows Server 2008 R2 および Windows 7 までサポートされていません。</remarks>
         FILE_SUPPORTS_EXTENDED_ATTRIBUTES = 8388608,


         /// <summary>ファイルシステムは FileID によるオープンをサポートします。詳細については、FILE_ID_BOTH_DIR_INFO を参照してください。</summary>
         /// <remarks>Windows Server 2008、Windows Vista、Windows Server 2003、および Windows XP: この値は Windows Server 2008 R2 および Windows 7 までサポートされていません。</remarks>
         FILE_SUPPORTS_OPEN_BY_FILE_ID = 16777216,


         /// <summary>指定されたボリュームは更新シーケンス番号 (USN) ジャーナルをサポートします。詳細については、Change Journal Records を参照してください。</summary>
         /// <remarks>Windows Server 2008、Windows Vista、Windows Server 2003、および Windows XP: この値は Windows Server 2008 R2 および Windows 7 までサポートされていません。</remarks>
         FILE_SUPPORTS_USN_JOURNAL = 33554432,


         /// <summary>指定されたボリュームはダイレクトアクセス (DAX) ボリュームです。</summary>
         /// <remarks>このフラグは Windows 10 バージョン 1607 で導入されました。</remarks>
         FILE_DAX_VOLUME = 536870912
      }
   }
}
