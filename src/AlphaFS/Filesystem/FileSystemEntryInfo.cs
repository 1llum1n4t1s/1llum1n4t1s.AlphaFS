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
   /// <summary>ファイルシステムエントリに関する情報を表します。
   /// <para>このクラスは継承できません。</para>
   /// </summary>
   [Serializable]
   [SecurityCritical]
   public sealed class FileSystemEntryInfo : IEquatable<FileSystemEntryInfo>
   {
      #region Fields

      private string _fullPath;
      private string _longFullPath;

      #endregion // Fields


      #region Constructor

      /// <summary><see cref="FileSystemEntryInfo"/> クラスの新しいインスタンスを初期化します。</summary>
      /// <param name="findData">NativeMethods.WIN32_FIND_DATA構造体。</param>
      internal FileSystemEntryInfo(NativeMethods.WIN32_FIND_DATA findData)
      {
         Win32FindData = findData;
      }

      #endregion // Constructor


      #region Properties

      /// <summary>ファイル名の8.3形式バージョン。</summary>
      public string AlternateFileName
      {
         // NativeMethods.FINDEX_INFO_LEVELS.Basicが使用されている場合、このプロパティは常に空です。

         get { return Win32FindData.cAlternateFileName; }
      }


      /// <summary>インスタンスの属性。</summary>
      public FileAttributes Attributes
      {
         get { return Win32FindData.dwFileAttributes; }
      }


      /// <summary>インスタンスの作成時刻。</summary>
      public DateTime CreationTime
      {
         get { return CreationTimeUtc.ToLocalTime(); }
      }


      /// <summary>インスタンスの協定世界時（UTC）での作成時刻。</summary>
      public DateTime CreationTimeUtc
      {
         get { return DateTime.FromFileTimeUtc(Win32FindData.ftCreationTime); }
      }


      /// <summary>インスタンスのファイル拡張子。</summary>
      public string Extension
      {
         get { return Path.GetExtension(Win32FindData.cFileName, false); }
      }


      /// <summary>インスタンスのファイル名。</summary>
      public string FileName
      {
         get { return Win32FindData.cFileName; }
      }


      /// <summary>インスタンスのファイルサイズ。</summary>
      public long FileSize
      {
         get { return NativeMethods.ToLong(Win32FindData.nFileSizeHigh, Win32FindData.nFileSizeLow); }
      }

      
      /// <summary>インスタンスのフルパス。</summary>
      public string FullPath
      {
         get { return _fullPath; }

         set
         {
            LongFullPath = value;
            _fullPath = Path.GetRegularPathCore(LongFullPath, GetFullPathOptions.None, false);
         }
      }


      /// <summary>インスタンスがバックアップまたは削除の候補であるかどうか。</summary>
      public bool IsArchive
      {
         get { return File.HasValidAttributes(Attributes) && (Attributes & FileAttributes.Archive) != 0; }
      }


      /// <summary>インスタンスが圧縮されているかどうか。</summary>
      public bool IsCompressed
      {
         get { return File.HasValidAttributes(Attributes) && (Attributes & FileAttributes.Compressed) != 0; }
      }


      /// <summary>将来の使用のために予約されています。</summary>
      public bool IsDevice
      {
         get { return File.HasValidAttributes(Attributes) && (Attributes & FileAttributes.Device) != 0; }
      }


      /// <summary>インスタンスがディレクトリであるかどうか。</summary>
      public bool IsDirectory
      {
         get { return File.IsDirectory(Attributes); }
      }


      /// <summary>インスタンスが暗号化されているかどうか。ファイルの場合、ファイル内の全データが暗号化されていることを意味します。ディレクトリの場合、新しく作成されるファイルとディレクトリのデフォルトが暗号化であることを意味します。</summary>
      public bool IsEncrypted
      {
         get { return File.HasValidAttributes(Attributes) && (Attributes & FileAttributes.Encrypted) != 0; }
      }


      /// <summary>インスタンスが隠しファイルであり、通常のディレクトリ一覧には含まれないかどうか。</summary>
      public bool IsHidden
      {
         get { return File.IsHidden(Attributes); }
      }


      /// <summary>インスタンスがマウントポイントであるかどうか。ローカルディレクトリおよびローカルボリュームに適用されます。</summary>
      public bool IsMountPoint
      {
         get { return ReparsePointTag == ReparsePointTag.MountPoint; }
      }


      /// <summary>インスタンスが特別な属性を持たない標準ファイルであるかどうか。この属性は単独で使用された場合にのみ有効です。</summary>
      public bool IsNormal
      {
         get { return File.HasValidAttributes(Attributes) && (Attributes & FileAttributes.Normal) != 0; }
      }


      /// <summary>インスタンスがオペレーティングシステムのコンテンツインデックスサービスによってインデックスされないかどうか。</summary>
      public bool IsNotContentIndexed
      {
         get { return File.HasValidAttributes(Attributes) && (Attributes & FileAttributes.NotContentIndexed) != 0; }
      }


      /// <summary>インスタンスがオフラインであるかどうか。ファイルのデータはすぐには利用できません。</summary>
      public bool IsOffline
      {
         get { return File.HasValidAttributes(Attributes) && (Attributes & FileAttributes.Offline) != 0; }
      }


      /// <summary>インスタンスが読み取り専用であるかどうか。</summary>
      public bool IsReadOnly
      {
         get { return File.IsReadOnly(Attributes); }
      }


      /// <summary>インスタンスがリパースポイントを含むかどうか。リパースポイントはファイルまたはディレクトリに関連付けられたユーザー定義データのブロックです。</summary>
      public bool IsReparsePoint
      {
         get { return File.HasValidAttributes(Attributes) && (Attributes & FileAttributes.ReparsePoint) != 0; }
      }


      /// <summary>インスタンスがスパースファイルであるかどうか。スパースファイルは通常、データの大部分がゼロで構成される大きなファイルです。</summary>
      public bool IsSparseFile
      {
         get { return File.HasValidAttributes(Attributes) && (Attributes & FileAttributes.SparseFile) != 0; }
      }


      /// <summary>インスタンスがシンボリックリンクであるかどうか。</summary>
      public bool IsSymbolicLink
      {
         get { return ReparsePointTag == ReparsePointTag.SymLink; }
      }


      /// <summary>インスタンスがシステムファイルであるかどうか。つまり、ファイルがオペレーティングシステムの一部であるか、オペレーティングシステムによって排他的に使用されます。</summary>
      public bool IsSystem
      {
         get { return File.HasValidAttributes(Attributes) && (Attributes & FileAttributes.System) != 0; }
      }


      /// <summary>インスタンスが一時ファイルであるかどうか。一時ファイルにはアプリケーション実行中に必要だが、終了後は不要になるデータが含まれます。
      /// ファイルシステムは、大容量ストレージにフラッシュするのではなく、すべてのデータをメモリに保持してアクセスを高速化しようとします。
      /// 一時ファイルは不要になった時点でアプリケーションが削除する必要があります。</summary>
      public bool IsTemporary
      {
         get { return File.HasValidAttributes(Attributes) && (Attributes & FileAttributes.Temporary) != 0; }
      }


      /// <summary>このエントリに最後にアクセスした時刻。</summary>
      public DateTime LastAccessTime
      {
         get { return LastAccessTimeUtc.ToLocalTime(); }
      }


      /// <summary>このエントリに最後にアクセスした協定世界時（UTC）での時刻。</summary>
      public DateTime LastAccessTimeUtc
      {
         get { return DateTime.FromFileTimeUtc(Win32FindData.ftLastAccessTime); }
      }


      /// <summary>このエントリが最後に変更された時刻。</summary>
      public DateTime LastWriteTime
      {
         get { return LastWriteTimeUtc.ToLocalTime(); }
      }


      /// <summary>このエントリが最後に変更された協定世界時（UTC）での時刻。</summary>
      public DateTime LastWriteTimeUtc
      {
         get { return DateTime.FromFileTimeUtc(Win32FindData.ftLastWriteTime); }
      }


      /// <summary>長いパス形式でのインスタンスのフルパス。</summary>
      public string LongFullPath
      {
         get { return _longFullPath; }

         private set { _longFullPath = Path.GetLongPathCore(value, GetFullPathOptions.None); }
      }


      /// <summary>インスタンスのリパースポイントタグ。</summary>
      public ReparsePointTag ReparsePointTag
      {
         get { return IsReparsePoint ? Win32FindData.dwReserved0 : ReparsePointTag.None; }
      }


      /// <summary>インスタンスの内部WIN32 FINDデータ。</summary>
      internal NativeMethods.WIN32_FIND_DATA Win32FindData { get; private set; }

      #endregion // Properties


      #region Methods

      /// <summary>FileSystemEntryInfoインスタンスの<see cref="FullPath"/>を返します。</summary>
      /// <returns>FileSystemEntryInfoインスタンスの<see cref="FullPath"/>。</returns>
      public override string ToString()
      {
         return FullPath;
      }


      /// <summary>特定の型のハッシュ関数として機能します。</summary>
      /// <returns>現在のオブジェクトのハッシュコード。</returns>
      public override int GetHashCode()
      {
         return Utils.CombineHashCodesOf(FullPath, LongFullPath);
      }


      /// <summary>指定されたオブジェクトが現在のオブジェクトと等しいかどうかを判断します。</summary>
      /// <param name="other">比較する別の<see cref="FileSystemInfo"/>インスタンス。</param>
      /// <returns>指定されたオブジェクトが現在のオブジェクトと等しい場合は<c>true</c>、それ以外の場合は<c>false</c>。</returns>
      public bool Equals(FileSystemEntryInfo other)
      {
         return null != other && GetType() == other.GetType() &&
                Equals(FileName, other.FileName) &&
                Equals(FullPath, other.FullPath) &&
                Equals(Attributes, other.Attributes) &&
                Equals(CreationTimeUtc, other.CreationTimeUtc) &&
                Equals(LastAccessTimeUtc, other.LastAccessTimeUtc);
      }


      /// <summary>指定されたオブジェクトが現在のオブジェクトと等しいかどうかを判断します。</summary>
      /// <param name="obj">比較する別のオブジェクト。</param>
      /// <returns>指定されたオブジェクトが現在のオブジェクトと等しい場合は<c>true</c>、それ以外の場合は<c>false</c>。</returns>
      public override bool Equals(object obj)
      {
         var other = obj as FileSystemEntryInfo;

         return null != other && Equals(other);
      }


      /// <summary>==演算子を実装します。</summary>
      /// <param name="left">左辺。</param>
      /// <param name="right">右辺。</param>
      /// <returns>演算子の結果。</returns>
      public static bool operator ==(FileSystemEntryInfo left, FileSystemEntryInfo right)
      {
         return ReferenceEquals(left, null) && ReferenceEquals(right, null) ||
                !ReferenceEquals(left, null) && !ReferenceEquals(right, null) && left.Equals(right);
      }


      /// <summary>!=演算子を実装します。</summary>
      /// <param name="left">左辺。</param>
      /// <param name="right">右辺。</param>
      /// <returns>演算子の結果。</returns>
      public static bool operator !=(FileSystemEntryInfo left, FileSystemEntryInfo right)
      {
         return !(left == right);
      }
      #endregion // Methods
   }
}
