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
using System.Globalization;
using System.IO;
using System.Security;

namespace Alphaleonis.Win32.Filesystem
{
   /// <summary>ファイルの作成、コピー、削除、移動、および開くためのプロパティとインスタンスメソッドを提供し、<see cref="FileStream"/> オブジェクトの作成を支援します。このクラスは継承できません。</summary>
   [Serializable]
   public sealed partial class FileInfo : FileSystemInfo
   {
      #region Fields

      [NonSerialized]
      private string _name;

      #endregion // Fields


      #region Constructors

      #region .NET

      /// <summary>ファイルパスのラッパーとして機能する <see cref="Alphaleonis.Win32.Filesystem.FileInfo"/> クラスの新しいインスタンスを初期化します。</summary>
      /// <param name="fileName">新しいファイルの完全修飾名、または相対ファイル名。パスの末尾にディレクトリ区切り文字を付けないでください。</param>
      /// <remarks>このコンストラクターはファイルの存在を確認しません。このコンストラクターは、後続の操作でファイルにアクセスするために使用される文字列のプレースホルダーです。</remarks>
      public FileInfo(string fileName) : this(null, fileName, PathFormat.RelativePath)
      {
      }

      #endregion // .NET


      /// <summary>[AlphaFS] ファイルパスのラッパーとして機能する <see cref="Alphaleonis.Win32.Filesystem.FileInfo"/> クラスの新しいインスタンスを初期化します。</summary>
      /// <param name="fileName">新しいファイルの完全修飾名、または相対ファイル名。パスの末尾にディレクトリ区切り文字を付けないでください。</param>
      /// <param name="pathFormat">パスパラメーターの形式を示します。</param>
      /// <remarks>このコンストラクターはファイルの存在を確認しません。このコンストラクターは、後続の操作でファイルにアクセスするために使用される文字列のプレースホルダーです。</remarks>
      public FileInfo(string fileName, PathFormat pathFormat) : this(null, fileName, pathFormat)
      {
      }


      /// <summary>[AlphaFS] ファイルパスのラッパーとして機能する <see cref="Alphaleonis.Win32.Filesystem.FileInfo"/> クラスの新しいインスタンスを初期化します。</summary>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="fileName">新しいファイルの完全修飾名、または相対ファイル名。パスの末尾にディレクトリ区切り文字を付けないでください。</param>
      /// <remarks>このコンストラクターはファイルの存在を確認しません。このコンストラクターは、後続の操作でファイルにアクセスするために使用される文字列のプレースホルダーです。</remarks>
      public FileInfo(KernelTransaction transaction, string fileName) : this(transaction, fileName, PathFormat.RelativePath)
      {
      }


      /// <summary>[AlphaFS] ファイルパスのラッパーとして機能する <see cref="Alphaleonis.Win32.Filesystem.FileInfo"/> クラスの新しいインスタンスを初期化します。</summary>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="fileName">新しいファイルの完全修飾名、または相対ファイル名。パスの末尾にディレクトリ区切り文字を付けないでください。</param>
      /// <param name="pathFormat">パスパラメーターの形式を示します。</param>
      /// <remarks>このコンストラクターはファイルの存在を確認しません。このコンストラクターは、後続の操作でファイルにアクセスするために使用される文字列のプレースホルダーです。</remarks>
      public FileInfo(KernelTransaction transaction, string fileName, PathFormat pathFormat)
      {
         InitializeCore(transaction, false, fileName, pathFormat);

         _name = Path.GetFileName(Path.RemoveTrailingDirectorySeparator(fileName), pathFormat != PathFormat.LongFullPath);
      }

      #endregion // Constructors


      #region Properties

      #region .NET

      /// <summary>親ディレクトリのインスタンスを取得します。</summary>
      /// <value>このファイルの親ディレクトリを表す <see cref="DirectoryInfo"/> オブジェクト。</value>
      /// <remarks>親ディレクトリを文字列として取得するには、DirectoryName プロパティを使用してください。</remarks>
      /// <exception cref="DirectoryNotFoundException"/>
      public DirectoryInfo Directory
      {
         get
         {
            var dirName = DirectoryName;
            return dirName == null ? null : new DirectoryInfo(Transaction, dirName, PathFormat.FullPath);
         }
      }


      /// <summary>ディレクトリの完全パスを表す文字列を取得します。</summary>
      /// <value>ディレクトリの完全パスを表す文字列。</value>
      /// <remarks>
      ///   <para>親ディレクトリを DirectoryInfo オブジェクトとして取得するには、Directory プロパティを使用してください。</para>
      ///   <para>最初に呼び出されたとき、FileInfo は Refresh を呼び出し、ファイルに関する情報をキャッシュします。</para>
      ///   <para>以降の呼び出しでは、最新の情報を取得するために Refresh を呼び出す必要があります。</para>
      /// </remarks>
      /// <exception cref="ArgumentNullException"/>
      public string DirectoryName
      {
         [SecurityCritical]
         get { return Path.GetDirectoryName(FullPath, false); }
      }


      /// <summary>ファイルが存在するかどうかを示す値を取得します。</summary>
      /// <value>ファイルが存在する場合は <c>true</c>、それ以外の場合は <c>false</c>。</value>
      /// <remarks>
      ///   <para>指定されたファイルの存在を確認しようとしたときにエラーが発生した場合、<see cref="Exists"/> プロパティは <c>false</c> を返します。</para>
      ///   <para>これは、無効な文字や文字数が多すぎるファイル名を渡した場合など、例外が発生する状況で起こる可能性があります。</para>
      ///   <para>また、ディスクの障害や欠落、またはファイルの読み取り権限がない場合にも発生します。</para>
      /// </remarks>
      [SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes")]
      public override bool Exists
      {
         [SecurityCritical]
         get
         {
            try
            {
               if (DataInitialised == -1)
               {
                  Refresh();
               }

               var attrs = Win32AttributeData.dwFileAttributes;

               return DataInitialised == 0 && File.HasValidAttributes(attrs) && !IsDirectory;
            }
            catch
            {
               return false;
            }
         }
      }


      /// <summary>現在のファイルが読み取り専用かどうかを判断する値を取得または設定します。</summary>
      /// <value>現在のファイルが読み取り専用の場合は <c>true</c>、それ以外の場合は <c>false</c>。</value>
      /// <remarks>
      ///   <para>IsReadOnly プロパティを使用して、現在のファイルが読み取り専用かどうかをすばやく判断または変更します。</para>
      ///   <para>最初に呼び出されたとき、FileInfo は Refresh を呼び出し、ファイルに関する情報をキャッシュします。</para>
      ///   <para>以降の呼び出しでは、最新の情報を取得するために Refresh を呼び出す必要があります。</para>
      /// </remarks>
      /// <exception cref="FileNotFoundException"/>
      /// <exception cref="IOException"/>
      public bool IsReadOnly
      {
         get { return EntryInfo == null || EntryInfo.IsReadOnly; }

         set
         {
            if (value)
            {
               Attributes |= FileAttributes.ReadOnly;
            }
            else
            {
               Attributes &= ~FileAttributes.ReadOnly;
            }
         }
      }


      /// <summary>現在のファイルのサイズをバイト単位で取得します。</summary>
      /// <value>現在のファイルのサイズ（バイト単位）。</value>
      /// <remarks>
      ///   <para>Length プロパティの値はプリキャッシュされています。</para>
      ///   <para>最新の値を取得するには、Refresh メソッドを呼び出してください。</para>
      /// </remarks>
      /// <exception cref="FileNotFoundException"/>
      /// <exception cref="IOException"/>
      [SuppressMessage("Microsoft.Design", "CA1065:DoNotRaiseExceptionsInUnexpectedLocations")]
      public long Length
      {
         [SecurityCritical]
         get
         {
            if (DataInitialised == -1)
            {
               Win32AttributeData = new NativeMethods.WIN32_FILE_ATTRIBUTE_DATA();
               Refresh();
            }

            // MSDN: .NET 3.5+: IOException: Refresh でデータを初期化できません。
            if (DataInitialised != 0)
            {
               NativeError.ThrowException(DataInitialised, FullName);
            }


            var attrs = Win32AttributeData.dwFileAttributes;

            // MSDN: .NET 3.5+: FileNotFoundException: ファイルが存在しないか、Length プロパティがディレクトリに対して呼び出されました。
            if (!File.HasValidAttributes(attrs))
            {
               NativeError.ThrowException(Win32Errors.ERROR_FILE_NOT_FOUND, FullName);
            }


            // MSDN: .NET 3.5+: FileNotFoundException: ファイルが存在しないか、Length プロパティがディレクトリに対して呼び出されました。
            if (File.IsDirectory(attrs))
            {
               NativeError.ThrowException(Win32Errors.ERROR_FILE_NOT_FOUND, string.Format(CultureInfo.InvariantCulture, Resources.Target_File_Is_A_Directory, FullName));
            }


            return Win32AttributeData.FileSize;
         }
      }


      /// <summary>ファイルの名前を取得します。</summary>
      /// <value>ファイルの名前。</value>
      /// <remarks>
      ///   <para>ファイルの名前にはファイル拡張子が含まれます。</para>
      ///   <para>最初に呼び出されたとき、<see cref="FileInfo"/> は Refresh を呼び出し、ファイルに関する情報をキャッシュします。</para>
      ///   <para>以降の呼び出しでは、最新の情報を取得するために Refresh を呼び出す必要があります。</para>
      ///   <para>ファイルの名前にはファイル拡張子が含まれます。</para>
      /// </remarks>
      public override string Name
      {
         get { return _name; }
      }

      #endregion // .NET

      #endregion // Properties


      #region Methods

      /// <summary>パスを文字列として返します。</summary>
      /// <returns>パス。</returns>
      public override string ToString()
      {
         return DisplayPath;
      }

      #endregion // Methods
   }
}
