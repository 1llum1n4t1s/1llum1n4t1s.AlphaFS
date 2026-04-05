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
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Runtime.InteropServices;
using System.Security;

namespace Alphaleonis.Win32.Filesystem
{
   /// <summary><see cref="FileInfo"/>と<see cref="DirectoryInfo"/>の両方のオブジェクトの基底クラスを提供します。</summary>
   [Serializable]
   [ComVisible(true)]
   public abstract class FileSystemInfo : MarshalByRefObject, IEquatable<FileSystemInfo>
   {
      #region Fields

      #region .NET

      /// <summary>ファイルまたはディレクトリの完全修飾パスを表します。</summary>
      /// <remarks>
      ///   <para><see cref="FileSystemInfo"/>から派生したクラスは、FullPathフィールドを使用して</para>
      ///   <para>操作対象オブジェクトのフルパスを取得できます。</para>
      /// </remarks>
      [SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")]
      protected string FullPath;

      /// <summary>ユーザーが最初に指定したパス（相対パスまたは絶対パス）。</summary>
      [SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")]
      protected string OriginalPath;

      #endregion // .NET


      // このフィールドはRefreshメソッドと組み合わせて使用します。成功時はゼロを格納し、
      // 失敗時はHResultを格納して汎用エラーを返せるようにします。
      [NonSerialized] internal int DataInitialised = -1;


      // 事前キャッシュされたFileSystemInfo情報。
      [NonSerialized] internal NativeMethods.WIN32_FILE_ATTRIBUTE_DATA Win32AttributeData;

      #endregion // Fields


      #region Properties

      #region .NET

      /// <summary>現在のファイルまたはディレクトリの属性を取得または設定します。</summary>
      /// <remarks>
      ///   <para>CreationTimeプロパティの値は事前キャッシュされています。</para>
      ///   <para>最新の値を取得するには、Refreshメソッドを呼び出してください。</para>
      /// </remarks>
      /// <value>現在の<see cref="FileSystemInfo"/>の<see cref="FileAttributes"/>。</value>
      ///
      /// <exception cref="FileNotFoundException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      public FileAttributes Attributes
      {
         [SecurityCritical]
         get
         {
            if (DataInitialised == -1)
            {
               Win32AttributeData = new NativeMethods.WIN32_FILE_ATTRIBUTE_DATA();
               Refresh();
            }

            // MSDN: .NET 3.5+: IOException: Refresh cannot initialize the data. 

            if (DataInitialised != 0)
            {
               NativeError.ThrowException(DataInitialised, FullPath);
            }

            return Win32AttributeData.dwFileAttributes;
         }


         [SecurityCritical]
         set
         {
            File.SetAttributesCore(Transaction, IsDirectory, LongFullName, value, PathFormat.LongFullPath);
            Reset();
         }
      }


      /// <summary>現在のファイルまたはディレクトリの作成時刻を取得または設定します。</summary>
      /// <remarks>
      ///   <para>CreationTimeプロパティの値は事前キャッシュされています。最新の値を取得するには、Refreshメソッドを呼び出してください。</para>
      ///   <para>このメソッドはネイティブ関数を使用するため、オペレーティングシステムによって継続的に更新されない値が返される可能性があり、不正確な値を返す場合があります。</para>
      ///   <para>FileSystemInfoオブジェクトで記述されたファイルが存在しない場合、このプロパティはローカル時間に調整された1601年1月1日午前0時（UTC）を返します。</para>
      ///   <para>NTFSフォーマットのドライブは、ファイル作成時刻などのファイルメタ情報を短期間キャッシュする場合があります。
      ///   このプロセスはファイルトンネリングと呼ばれます。その結果、既存のファイルを上書きまたは置換する場合は、明示的にファイルの作成時刻を設定する必要がある場合があります。</para>
      /// </remarks>
      /// <value>現在の<see cref="FileSystemInfo"/>オブジェクトの作成日時。</value>
      ///
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      public DateTime CreationTime
      {
         [SecurityCritical] get { return CreationTimeUtc.ToLocalTime(); }

         [SecurityCritical] set { CreationTimeUtc = value.ToUniversalTime(); }
      }


      /// <summary>現在のファイルまたはディレクトリの協定世界時（UTC）での作成時刻を取得または設定します。</summary>
      /// <remarks>
      ///   <para>CreationTimeUtcプロパティの値は事前キャッシュされています。最新の値を取得するには、Refreshメソッドを呼び出してください。</para>
      ///   <para>このメソッドはネイティブ関数を使用するため、オペレーティングシステムによって継続的に更新されない値が返される可能性があり、不正確な値を返す場合があります。</para>
      ///   <para>最新の値を取得するには、Refreshメソッドを呼び出してください。</para>
      ///   <para>FileSystemInfoオブジェクトで記述されたファイルが存在しない場合、このプロパティは1601年1月1日午前0時（UTC）を返します。</para>
      ///   <para>NTFSフォーマットのドライブは、ファイル作成時刻などのファイルメタ情報を短期間キャッシュする場合があります。
      ///   このプロセスはファイルトンネリングと呼ばれます。その結果、既存のファイルを上書きまたは置換する場合は、明示的にファイルの作成時刻を設定する必要がある場合があります。</para>
      /// </remarks>
      /// <value>現在の<see cref="FileSystemInfo"/>オブジェクトのUTC形式での作成日時。</value>
      ///
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      [ComVisible(false)]
      public DateTime CreationTimeUtc
      {
         [SecurityCritical]
         get
         {
            if (DataInitialised == -1)
            {
               Win32AttributeData = new NativeMethods.WIN32_FILE_ATTRIBUTE_DATA();
               Refresh();
            }

            // MSDN: .NET 3.5+: IOException: Refresh cannot initialize the data. 
            if (DataInitialised != 0)
            {
               NativeError.ThrowException(DataInitialised, LongFullName);
            }

            return DateTime.FromFileTimeUtc(Win32AttributeData.ftCreationTime);
         }


         [SecurityCritical]
         set
         {
            File.SetFsoDateTimeCore(Transaction, IsDirectory, LongFullName, value, null, null, false, PathFormat.LongFullPath);

            Reset();
         }
      }


      /// <summary>ファイルまたはディレクトリが存在するかどうかを示す値を取得します。</summary>
      /// <remarks>
      ///   <para>指定されたファイルまたはディレクトリの存在を確認しようとした際にエラーが発生した場合、<see cref="Exists"/>プロパティは<c>false</c>を返します。</para>
      ///   <para>これは、無効な文字または文字数が多すぎるディレクトリ名やファイル名を渡した場合など、例外が発生する状況で起こり得ます。</para>
      ///   <para>また、ディスクの障害や欠落、またはファイルやディレクトリの読み取り権限がない場合にも発生します。</para>
      /// </remarks>
      /// <value>ファイルまたはディレクトリが存在する場合は<c>true</c>、それ以外の場合は<c>false</c>。</value>
      public abstract bool Exists { get; }


      /// <summary>ファイルの拡張子部分を表す文字列を取得します。</summary>
      /// <remarks>
      ///   Extensionプロパティは、ピリオド（.）を含む<see cref="FileSystemInfo"/>の拡張子を返します。
      ///   例えば、c:\NewFile.txtというファイルの場合、このプロパティは「.txt」を返します。
      /// </remarks>
      /// <value><see cref="FileSystemInfo"/>の拡張子を含む文字列。</value>
      public string Extension
      {
         get { return Path.GetExtension(FullPath, false); }
      }


      /// <summary>ディレクトリまたはファイルのフルパスを取得します。</summary>
      /// <value>フルパスを含む文字列。</value>
      public virtual string FullName
      {
         [SecurityCritical] get { return FullPath; }
      }

      
      /// <summary>現在のファイルまたはディレクトリに最後にアクセスした時刻を取得または設定します。</summary>
      /// <remarks>
      ///   <para>LastAccessTimeプロパティの値は事前キャッシュされています。最新の値を取得するには、Refreshメソッドを呼び出してください。</para>
      ///   <para>このメソッドはネイティブ関数を使用するため、オペレーティングシステムによって継続的に更新されない値が返される可能性があり、不正確な値を返す場合があります。</para>
      ///   <para>FileSystemInfoオブジェクトで記述されたファイルが存在しない場合、このプロパティはローカル時間に調整された1601年1月1日午前0時（UTC）を返します。</para>
      /// </remarks>
      /// <value>現在のファイルまたはディレクトリに最後にアクセスした時刻。</value>
      ///
      /// <exception cref="IOException"/>
      public DateTime LastAccessTime
      {
         [SecurityCritical] get { return LastAccessTimeUtc.ToLocalTime(); }

         [SecurityCritical] set { LastAccessTimeUtc = value.ToUniversalTime(); }
      }


      /// <summary>現在のファイルまたはディレクトリに最後にアクセスした協定世界時（UTC）での時刻を取得または設定します。</summary>
      /// <remarks>
      ///   <para>LastAccessTimeUtcプロパティの値は事前キャッシュされています。最新の値を取得するには、Refreshメソッドを呼び出してください。</para>
      ///   <para>このメソッドはネイティブ関数を使用するため、オペレーティングシステムによって継続的に更新されない値が返される可能性があり、不正確な値を返す場合があります。</para>
      ///   <para>FileSystemInfoオブジェクトで記述されたファイルが存在しない場合、このプロパティはローカル時間に調整された1601年1月1日午前0時（UTC）を返します。</para>
      /// </remarks>
      /// <value>現在のファイルまたはディレクトリに最後にアクセスしたUTC時刻。</value>
      ///
      /// <exception cref="IOException"/>
      [ComVisible(false)]
      public DateTime LastAccessTimeUtc
      {
         [SecurityCritical]
         get
         {
            if (DataInitialised == -1)
            {
               Win32AttributeData = new NativeMethods.WIN32_FILE_ATTRIBUTE_DATA();
               Refresh();
            }

            // MSDN: .NET 3.5+: IOException: Refresh cannot initialize the data. 
            if (DataInitialised != 0)
            {
               NativeError.ThrowException(DataInitialised, LongFullName);
            }

            return DateTime.FromFileTimeUtc(Win32AttributeData.ftLastAccessTime);
         }


         [SecurityCritical]
         set
         {
            File.SetFsoDateTimeCore(Transaction, IsDirectory, LongFullName, null, value, null, false, PathFormat.LongFullPath);

            Reset();
         }
      }


      /// <summary>現在のファイルまたはディレクトリが最後に書き込まれた時刻を取得または設定します。</summary>
      /// <remarks>
      ///   <para>LastWriteTimeプロパティの値は事前キャッシュされています。最新の値を取得するには、Refreshメソッドを呼び出してください。</para>
      ///   <para>このメソッドはネイティブ関数を使用するため、オペレーティングシステムによって継続的に更新されない値が返される可能性があり、不正確な値を返す場合があります。</para>
      ///   <para>FileSystemInfoオブジェクトで記述されたファイルが存在しない場合、このプロパティはローカル時間に調整された1601年1月1日午前0時（UTC）を返します。</para>
      /// </remarks>
      /// <value>現在のファイルが最後に書き込まれた時刻。</value>
      ///
      /// <exception cref="IOException"/>
      public DateTime LastWriteTime
      {
         get { return LastWriteTimeUtc.ToLocalTime(); }

         set { LastWriteTimeUtc = value.ToUniversalTime(); }
      }


      /// <summary>現在のファイルまたはディレクトリが最後に書き込まれた協定世界時（UTC）での時刻を取得または設定します。</summary>
      /// <remarks>
      ///   <para>LastWriteTimeUtcプロパティの値は事前キャッシュされています。最新の値を取得するには、Refreshメソッドを呼び出してください。</para>
      ///   <para>このメソッドはネイティブ関数を使用するため、オペレーティングシステムによって継続的に更新されない値が返される可能性があり、不正確な値を返す場合があります。</para>
      ///   <para>FileSystemInfoオブジェクトで記述されたファイルが存在しない場合、このプロパティはローカル時間に調整された1601年1月1日午前0時（UTC）を返します。</para>
      /// </remarks>
      /// <value>現在のファイルが最後に書き込まれたUTC時刻。</value>
      [ComVisible(false)]
      public DateTime LastWriteTimeUtc
      {
         [SecurityCritical]
         get
         {
            if (DataInitialised == -1)
            {
               Win32AttributeData = new NativeMethods.WIN32_FILE_ATTRIBUTE_DATA();
               Refresh();
            }

            // MSDN: .NET 3.5+: IOException: Refresh cannot initialize the data. 
            if (DataInitialised != 0)
            {
               NativeError.ThrowException(DataInitialised, LongFullName);
            }

            return DateTime.FromFileTimeUtc(Win32AttributeData.ftLastWriteTime);
         }


         [SecurityCritical]
         set
         {
            File.SetFsoDateTimeCore(Transaction, IsDirectory, LongFullName, null, null, value, false, PathFormat.LongFullPath);

            Reset();
         }
      }


      /// <summary>
      ///   ファイルの場合はファイル名を取得します。ディレクトリの場合は、階層が存在する場合は階層内の最後のディレクトリ名を取得します。
      ///   <para>それ以外の場合、Nameプロパティはディレクトリ名を取得します。</para>
      /// </summary>
      /// <remarks>
      ///   <para>ディレクトリの場合、Nameは親ディレクトリの名前のみを返します（例: c:\Dirではなく Dir）。</para>
      ///   <para>サブディレクトリの場合、Nameはサブディレクトリの名前のみを返します（例: c:\Dir\Sub1ではなく Sub1）。</para>
      ///   <para>ファイルの場合、Nameはファイル名と拡張子のみを返します（例: c:\Dir\Myfile.txtではなく MyFile.txt）。</para>
      /// </remarks>
      /// <value>
      ///   <para>親ディレクトリの名前、階層内の最後のディレクトリの名前、</para>
      ///   <para>またはファイル拡張子を含むファイル名。</para>
      /// </value>
      public abstract string Name { get; }

      #endregion // .NET


      #region AlphaFS

      /// <summary>パスを文字列として返します。</summary>
      protected internal string DisplayPath { get; protected set; }


      private FileSystemEntryInfo _entryInfo;

      /// <summary>[AlphaFS] <see cref="FileSystemEntryInfo"/>クラスのインスタンスを取得します。</summary>
      public FileSystemEntryInfo EntryInfo
      {
         [SecurityCritical]
         get
         {
            if (null == _entryInfo)
            {
               Win32AttributeData = new NativeMethods.WIN32_FILE_ATTRIBUTE_DATA();
               RefreshEntryInfo();
            }

            // MSDN: .NET 3.5+: IOException: Refresh cannot initialize the data. 
            if (DataInitialised > 0)
            {
               NativeError.ThrowException(DataInitialised, LongFullName);
            }

            return _entryInfo;
         }


         internal set
         {
            _entryInfo = value;

            DataInitialised = value == null ? -1 : 0;

            if (DataInitialised == 0 && null != _entryInfo)
            {
               Win32AttributeData = new NativeMethods.WIN32_FILE_ATTRIBUTE_DATA(_entryInfo.Win32FindData);
            }
         }
      }


      /// <summary>[AlphaFS] コンストラクタに渡された初期の「IsDirectory」インジケーター。</summary>
      protected bool IsDirectory { get; set; }


      /// <summary>Unicode（LongPath）形式でのファイルシステムオブジェクトのフルパス。</summary>
      protected string LongFullName { get; set; }


      /// <summary>[AlphaFS] コンストラクタに渡されたKernelTransactionを表します。</summary>
      protected KernelTransaction Transaction { get; set; }

      #endregion // AlphaFS

      #endregion // Properties


      #region Methods

      #region .NET

      /// <summary>ファイルまたはディレクトリを削除します。</summary>
      [SecurityCritical]
      public abstract void Delete();


      /// <summary>オブジェクトの状態を更新します。</summary>
      /// <remarks>
      ///   <para>FileSystemInfo.Refresh()は現在のファイルシステムからファイルのスナップショットを取得します。</para>
      ///   <para>ファイルシステムが不正確または古い情報を返した場合でも、Refreshは基盤のファイルシステムを修正できません。</para>
      ///   <para>これはWindows 98などのプラットフォームで発生する可能性があります。</para>
      ///   <para>属性情報を取得する前にRefresh()を呼び出す必要があります。そうしないと情報が古くなります。</para>
      /// </remarks>
      [SecurityCritical]
      public void Refresh()
      {
         DataInitialised = File.FillAttributeInfoCore(Transaction, LongFullName, ref Win32AttributeData, false, false);

         IsDirectory = File.IsDirectory(Win32AttributeData.dwFileAttributes);
      }


      /// <summary>現在のオブジェクトを表す文字列を返します。</summary>
      /// <remarks>
      ///   ToStringは.NET Frameworkの主要なフォーマットメソッドです。オブジェクトを表示に適した文字列表現に変換します。
      /// </remarks>
      /// <returns>このインスタンスを表す文字列。</returns>
      public override string ToString()
      {
         // "Alphaleonis.Win32.Filesystem.FileSystemInfo"
         return GetType().ToString();
      }


      /// <summary>特定の型のハッシュ関数として機能します。</summary>
      /// <returns>現在のオブジェクトのハッシュコード。</returns>
      public override int GetHashCode()
      {
         return null != FullName ? FullName.GetHashCode() : 0;
      }


      /// <summary>指定されたオブジェクトが現在のオブジェクトと等しいかどうかを判断します。</summary>
      /// <param name="other">比較する別の<see cref="FileSystemInfo"/>インスタンス。</param>
      /// <returns>指定されたオブジェクトが現在のオブジェクトと等しい場合は<c>true</c>、それ以外の場合は<c>false</c>。</returns>
      public bool Equals(FileSystemInfo other)
      {
         return null != other && GetType() == other.GetType() &&
                Equals(Name, other.Name) &&
                Equals(FullName, other.FullName) &&
                Equals(Attributes, other.Attributes) &&
                Equals(CreationTimeUtc, other.CreationTimeUtc) &&
                Equals(LastAccessTimeUtc, other.LastAccessTimeUtc);
      }


      /// <summary>指定されたオブジェクトが現在のオブジェクトと等しいかどうかを判断します。</summary>
      /// <param name="obj">比較する別のオブジェクト。</param>
      /// <returns>指定されたオブジェクトが現在のオブジェクトと等しい場合は<c>true</c>、それ以外の場合は<c>false</c>。</returns>
      public override bool Equals(object obj)
      {
         var other = obj as FileSystemInfo;

         return null != other && Equals(other);
      }


      /// <summary>==演算子を実装します。</summary>
      /// <param name="left">左辺。</param>
      /// <param name="right">右辺。</param>
      /// <returns>演算子の結果。</returns>
      public static bool operator ==(FileSystemInfo left, FileSystemInfo right)
      {
         return ReferenceEquals(left, null) && ReferenceEquals(right, null) ||
                !ReferenceEquals(left, null) && !ReferenceEquals(right, null) && left.Equals(right);
      }


      /// <summary>!=演算子を実装します。</summary>
      /// <param name="left">左辺。</param>
      /// <param name="right">右辺。</param>
      /// <returns>演算子の結果。</returns>
      public static bool operator !=(FileSystemInfo left, FileSystemInfo right)
      {
         return !(left == right);
      }

      #endregion // .NET


      /// <summary>[AlphaFS] 新しい宛先パスで現在の<see cref="FileSystemInfo"/>インスタンス（<see cref="DirectoryInfo"/>または<see cref="FileInfo"/>）を更新します。</summary>
      internal void UpdateSourcePath(string destinationPath, string destinationPathLp)
      {
         LongFullName = destinationPathLp;

         FullPath = null != destinationPathLp ? Path.GetRegularPathCore(LongFullName, GetFullPathOptions.None, false) : null;

         OriginalPath = destinationPath;

         DisplayPath = null != OriginalPath ? Path.GetRegularPathCore(OriginalPath, GetFullPathOptions.None, false) : null;

         // FileSystemInfoインスタンスに関するキャッシュ情報をフラッシュします。
         Reset();
      }


      /// <summary>[AlphaFS] <see cref="FileSystemEntryInfo"/> EntryInfoプロパティの状態を更新します。</summary>
      /// <remarks>
      ///   <para>FileSystemInfo.RefreshEntryInfo()は現在のファイルシステムからファイルのスナップショットを取得します。</para>
      ///   <para>ファイルシステムが不正確または古い情報を返した場合でも、Refreshは基盤のファイルシステムを修正できません。</para>
      ///   <para>これはWindows 98などのプラットフォームで発生する可能性があります。</para>
      ///   <para>属性情報を取得する前にRefresh()を呼び出す必要があります。そうしないと情報が古くなります。</para>
      /// </remarks>
      [SecurityCritical]
      protected void RefreshEntryInfo()
      {
         _entryInfo = File.GetFileSystemEntryInfoCore(Transaction, IsDirectory, LongFullName, true, PathFormat.LongFullPath);

         if (null == _entryInfo)
         {
            DataInitialised = -1;
         }

         else
         {
            DataInitialised = 0;
            Win32AttributeData = new NativeMethods.WIN32_FILE_ATTRIBUTE_DATA(_entryInfo.Win32FindData);
         }
      }


      /// <summary>[AlphaFS] ファイルシステムオブジェクトの状態を未初期化にリセットします。</summary>
      private void Reset()
      {
         DataInitialised = -1;
      }


      /// <summary>指定されたファイル名を初期化します。</summary>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="NotSupportedException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="isFolder"><paramref name="path"/>がファイルかディレクトリかを指定します。</param>
      /// <param name="path">ファイルのフルパスと名前。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      internal void InitializeCore(KernelTransaction transaction, bool isFolder, string path, PathFormat pathFormat)
      {
         if (pathFormat == PathFormat.RelativePath)
         {
            Path.CheckSupportedPathFormat(path, true, true);
         }

         LongFullName = Path.GetExtendedLengthPathCore(transaction, path, pathFormat, GetFullPathOptions.TrimEnd | (isFolder ? GetFullPathOptions.RemoveTrailingDirectorySeparator : 0) | GetFullPathOptions.ContinueOnNonExist);
         
         // (MSDNに記載なし): .NET 4以降では、FileSystemInfoインスタンスの作成前にパスパラメータの末尾のスペースが除去されます。

         FullPath = Path.GetRegularPathCore(LongFullName, GetFullPathOptions.None, false);

         IsDirectory = isFolder;

         Transaction = transaction;

         OriginalPath = FullPath.Length == 2 && FullPath[1] == Path.VolumeSeparatorChar ? Path.CurrentDirectoryPrefix : path;

         DisplayPath = OriginalPath.Length != 2 || OriginalPath[1] != Path.VolumeSeparatorChar ? Path.GetRegularPathCore(OriginalPath, GetFullPathOptions.None, false) : Path.CurrentDirectoryPrefix;
      }


      internal static SafeFindFileHandle FindFirstFileNative(KernelTransaction transaction, string pathLp, NativeMethods.FINDEX_INFO_LEVELS infoLevel, NativeMethods.FINDEX_SEARCH_OPS searchOption, NativeMethods.FIND_FIRST_EX_FLAGS additionalFlags, out int lastError, out NativeMethods.WIN32_FIND_DATA win32FindData)
      {
         var safeHandle = null == transaction || !NativeMethods.IsAtLeastWindowsVista

            // FindFirstFileEx() / FindFirstFileTransacted()
            // 2013-01-13: MSDNがLongPathの使用を確認。

            // 末尾のバックスラッシュは許可されていません。
            ? NativeMethods.FindFirstFileEx(Path.RemoveTrailingDirectorySeparator(pathLp), infoLevel, out win32FindData, searchOption, IntPtr.Zero, additionalFlags)

            : NativeMethods.FindFirstFileTransacted(Path.RemoveTrailingDirectorySeparator(pathLp), infoLevel, out win32FindData, searchOption, IntPtr.Zero, additionalFlags, transaction.SafeHandle);

         lastError = Marshal.GetLastWin32Error();

         if (!NativeMethods.IsValidHandle(safeHandle, false))
         {
            safeHandle = null;
         }


         return safeHandle;
      }

      #endregion // Methods
   }
}
