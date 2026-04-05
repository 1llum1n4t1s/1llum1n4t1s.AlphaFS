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

using Alphaleonis.Win32.Filesystem;
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

namespace Alphaleonis.Win32.Network
{
   /// <summary>分散ファイルシステム (DFS) ルートまたはリンクに関する情報を含みます。このクラスは継承できません。
   /// <para>この構造体には、ルートまたはリンクの名前、状態、GUID、タイムアウト、ターゲット数、および各ターゲットに関する情報が含まれます。</para>
   /// </summary>
   [Serializable]
   [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Dfs")]
   public sealed class DfsInfo
   {
      #region コンストラクター

      /// <summary>DFS ルートまたはリンクターゲットのラッパーとして機能する <see cref="DfsInfo"/> クラスの新しいインスタンスを初期化します。</summary>
      public DfsInfo()
      {
      }

      /// <summary>DFS ルートまたはリンクターゲットのラッパーとして機能する <see cref="DfsInfo"/> クラスの新しいインスタンスを初期化します。</summary>
      /// <param name="structure">初期化された <see cref="NativeMethods.DFS_INFO_9"/> インスタンス。</param>
      internal DfsInfo(NativeMethods.DFS_INFO_9 structure)
      {
         Comment = structure.Comment;
         EntryPath = structure.EntryPath;
         State = structure.State;
         Timeout = structure.Timeout;
         Guid = structure.Guid;
         MetadataSize = structure.MetadataSize;
         PropertyFlags = structure.PropertyFlags;
         SecurityDescriptor = structure.pSecurityDescriptor;

         if (structure.NumberOfStorages > 0)
         {
            var sizeOfStruct = Marshal.SizeOf<NativeMethods.DFS_STORAGE_INFO_1>();

            for (var i = 0; i < structure.NumberOfStorages; i++)
               _storageInfoCollection.Add(new DfsStorageInfo(Marshal.PtrToStructure<NativeMethods.DFS_STORAGE_INFO_1>(new IntPtr(structure.Storage.ToInt64() + i * sizeOfStruct))));
         }
      }

      #endregion // コンストラクター

      #region メソッド

      /// <summary>DFS ルートまたはリンクの汎用名前付け規則 (UNC) パスを返します。</summary>
      /// <returns>このインスタンスを表す文字列。</returns>
      public override string ToString()
      {
         return EntryPath;
      }

      #endregion // メソッド

      #region プロパティ

      private DirectoryInfo _directoryInfo;

      /// <summary>DFS ルートまたはリンクの <see cref="DirectoryInfo"/> インスタンス。</summary>
      public DirectoryInfo DirectoryInfo
      {
         get { return _directoryInfo ?? (_directoryInfo = new DirectoryInfo(null, EntryPath, PathFormat.FullPath)); }
      }

      /// <summary>DFS ルートまたはリンクのコメント。</summary>
      public string Comment { get; internal set; }

      /// <summary>DFS ルートまたはリンクの汎用名前付け規則 (UNC) パス。</summary>
      public string EntryPath { get; internal set; }

      /// <summary>DFS ルートまたはリンクの GUID を指定します。</summary>
      public Guid Guid { get; internal set; }


      private readonly List<DfsStorageInfo> _storageInfoCollection = new List<DfsStorageInfo>();

      /// <summary>DFS ルートまたはリンクの DFS ターゲットのコレクション。</summary>
      public IEnumerable<DfsStorageInfo> StorageInfoCollection
      {
         get { return _storageInfoCollection; }
      }

      /// <summary>DFS ルートまたはリンクを記述するビットフラグのセットを指定する <see cref="DfsVolumeStates"/> 列挙型。</summary>
      public DfsVolumeStates State { get; internal set; }

      //DfsVolumeStates flavorBits = (structure3.State & (DfsVolumeStates) DfsNamespaceFlavors.All);
      //If (flavorBits == DFS_VOLUME_FLAVOR_STANDALONE)     // Namespace is stand-alone DFS.
      //else if (flavorBits == DFS_VOLUME_FLAVOR_AD_BLOB)   // Namespace is AD Blob.
      //else StateBits = (Flavor & DFS_VOLUME_STATES)        // Unknown flavor.
      // StateBits can be one of the following: 
      //  (DFS_VOLUME_STATE_OK, DFS_VOLUME_STATE_INCONSISTENT, 
      //   DFS_VOLUME_STATE_OFFLINE or DFS_VOLUME_STATE_ONLINE)
      //State = flavorBits | structure3.State;

      /// <summary>DFS ルートまたはリンクのタイムアウト（秒単位）を指定します。</summary>
      public long Timeout { get; internal set; }

      /// <summary>DFS 名前空間、ルート、またはリンクの特定のプロパティを記述するフラグのセットを指定します。</summary>
      [SuppressMessage("Microsoft.Naming", "CA1726:UsePreferredTerms", MessageId = "Flags")]
      public DfsPropertyFlags PropertyFlags { get; internal set; }
      
      /// <summary>ドメインベースの DFS 名前空間の場合、対応する Active Directory データ BLOB のサイズ（バイト単位）を指定します。
      /// スタンドアロン DFS 名前空間の場合、レジストリに格納されているメタデータのサイズを指定します。
      /// これには、関連付けられた特定のデータ項目に加えて、キー名と値の名前が含まれます。このフィールドは DFS ルートでのみ有効です。
      /// </summary>
      public long MetadataSize { get; internal set; }


      /// <summary>DFS リンクのリパースポイントに関連付ける自己相対セキュリティ記述子を指定する SECURITY_DESCRIPTOR 構造体へのポインター。
      /// このフィールドは DFS リンクでのみ有効です。
      /// </summary>
      public IntPtr SecurityDescriptor { get; internal set; }
      
      #endregion // プロパティ
   }
}
