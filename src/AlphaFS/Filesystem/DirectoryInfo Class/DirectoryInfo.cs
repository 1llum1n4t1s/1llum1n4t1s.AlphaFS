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
using System.Security;

namespace Alphaleonis.Win32.Filesystem
{
   /// <summary>ディレクトリおよびサブディレクトリの作成、移動、列挙を行うインスタンスメソッドを公開します。このクラスは継承できません。</summary>
   [Serializable]
   public sealed partial class DirectoryInfo : FileSystemInfo
   {
      #region Constructors

      #region .NET

      /// <summary><see cref="DirectoryInfo"/> クラスの新しいインスタンスを指定されたパスで初期化します。</summary>
      /// <param name="path"><see cref="DirectoryInfo"/> を作成するパス。</param>
      /// <remarks>
      /// このコンストラクターはディレクトリの存在を確認しません。このコンストラクターは、後続の操作でディスクにアクセスするために使用される文字列のプレースホルダーです。
      /// path パラメーターには、UNC (Universal Naming Convention) 共有上のファイルを含むファイル名を指定できます。
      /// </remarks>
      public DirectoryInfo(string path) : this(null, path, PathFormat.RelativePath)
      {
      }

      #endregion // .NET


      /// <summary>[AlphaFS] <see cref="DirectoryInfo"/> クラスの新しいインスタンスを指定されたパスで初期化します。</summary>
      /// <param name="path"><see cref="DirectoryInfo"/> を作成するパス。</param>
      /// <param name="pathFormat">パスパラメーターの形式を示します。</param>
      /// <remarks>このコンストラクターはディレクトリの存在を確認しません。このコンストラクターは、後続の操作でディスクにアクセスするために使用される文字列のプレースホルダーです。</remarks>
      public DirectoryInfo(string path, PathFormat pathFormat) : this(null, path, pathFormat)
      {
      }

      /// <summary>[AlphaFS] 特殊な内部実装です。</summary>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="fullPath"><see cref="DirectoryInfo"/> を作成する完全パス。</param>
      /// <param name="junk1">使用しません。</param>
      /// <param name="junk2">使用しません。</param>
      /// <remarks>このコンストラクターはディレクトリの存在を確認しません。このコンストラクターは、後続の操作でディスクにアクセスするために使用される文字列のプレースホルダーです。</remarks>
      [SuppressMessage("Microsoft.Usage", "CA1801:ReviewUnusedParameters", MessageId = "junk1")]
      [SuppressMessage("Microsoft.Usage", "CA1801:ReviewUnusedParameters", MessageId = "junk2")]
      private DirectoryInfo(KernelTransaction transaction, string fullPath, bool junk1, bool junk2)
      {
         IsDirectory = true;
         Transaction = transaction;

         LongFullName = Path.GetLongPathCore(fullPath, GetFullPathOptions.None);

         // .NET Framework は Parent / Root / 列挙結果の DirectoryInfo.ToString() に「名前だけ」を返していたが、
         // .NET Core 以降の System.IO はフルパスを返す。ドロップイン代替として現行挙動に合わせる。
         OriginalPath = fullPath;

         FullPath = fullPath;

         DisplayPath = OriginalPath.Length != 2 || OriginalPath[1] != Path.VolumeSeparatorChar ? OriginalPath : Path.CurrentDirectoryPrefix;
      }


      #region Transactional

      /// <summary>[AlphaFS] <see cref="DirectoryInfo"/> クラスの新しいインスタンスを指定されたパスで初期化します。</summary>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path"><see cref="DirectoryInfo"/> を作成するパス。</param>
      /// <remarks>このコンストラクターはディレクトリの存在を確認しません。このコンストラクターは、後続の操作でディスクにアクセスするために使用される文字列のプレースホルダーです。</remarks>
      public DirectoryInfo(KernelTransaction transaction, string path) : this(transaction, path, PathFormat.RelativePath)
      {
      }


      /// <summary>[AlphaFS] <see cref="DirectoryInfo"/> クラスの新しいインスタンスを指定されたパスで初期化します。</summary>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path"><see cref="DirectoryInfo"/> を作成するパス。</param>
      /// <param name="pathFormat">パスパラメーターの形式を示します。</param>
      /// <remarks>このコンストラクターはディレクトリの存在を確認しません。このコンストラクターは、後続の操作でディスクにアクセスするために使用される文字列のプレースホルダーです。</remarks>
      public DirectoryInfo(KernelTransaction transaction, string path, PathFormat pathFormat)
      {
         InitializeCore(transaction, true, path, pathFormat);
      }

      #endregion // Transactional

      #endregion // Constructors


      #region Properties

      #region .NET

      /// <summary>ディレクトリが存在するかどうかを示す値を取得します。</summary>
      /// <remarks>
      ///   <para>指定されたディレクトリの存在を確認しようとしたときにエラーが発生した場合、<see cref="Exists"/> プロパティは <c>false</c> を返します。</para>
      ///   <para>これは、無効な文字や文字数が多すぎるディレクトリ名を渡した場合など、例外が発生する状況で起こる可能性があります。</para>
      ///   <para>また、ディスクの障害や欠落、またはディレクトリの読み取り権限がない場合にも発生します。</para>
      /// </remarks>
      /// <value>ディレクトリが存在する場合は <c>true</c>、それ以外の場合は <c>false</c>。</value>
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

               return DataInitialised == 0 && IsDirectory;
            }
            catch
            {
               return false;
            }
         }
      }


      /// <summary>この <see cref="DirectoryInfo"/> インスタンスの名前を取得します。</summary>
      /// <value>ディレクトリ名。</value>
      /// <remarks>
      ///   <para>この Name プロパティは、"Bin" のようなディレクトリ名のみを返します。</para>
      ///   <para>"c:\public\Bin" のような完全パスを取得するには、FullName プロパティを使用してください。</para>
      /// </remarks>
      public override string Name
      {
         get { return FullPath.Length > 3 ? Path.GetFileName(Path.RemoveTrailingDirectorySeparator(FullPath), true) : FullPath; }
      }


      /// <summary>指定されたサブディレクトリの親ディレクトリを取得します。</summary>
      /// <value>親ディレクトリ。パスが null の場合、またはファイルパスがルート（"\"、"C:"、"\\server\share" など）を示す場合は null。</value>
      public DirectoryInfo Parent
      {
         [SecurityCritical]
         get
         {
            var path = FullPath;

            if (path.Length > 3)
            {
               path = Path.RemoveTrailingDirectorySeparator(FullPath);
            }

            var dirName = Path.GetDirectoryName(path, false);

            return null != dirName ? new DirectoryInfo(Transaction, dirName, true, true) : null;
         }
      }


      /// <summary>ディレクトリのルート部分を取得します。</summary>
      /// <value>ディレクトリのルートを表すオブジェクト。</value>
      public DirectoryInfo Root
      {
         [SecurityCritical]
         get { return new DirectoryInfo(Transaction, Path.GetPathRoot(FullPath, false), PathFormat.RelativePath); }
      }

      #endregion // .NET

      #endregion // Properties


      #region Methods

      /// <summary>ユーザーが渡した元のパスを返します。</summary>
      /// <returns>このオブジェクトを表す文字列。</returns>
      public override string ToString()
      {
         return DisplayPath;
      }

      #endregion // Methods
   }
}
