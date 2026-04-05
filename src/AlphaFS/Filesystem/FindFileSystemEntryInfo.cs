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
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Security;
using System.Text.RegularExpressions;

#if !NET35
using System.Threading;
#endif

namespace Alphaleonis.Win32.Filesystem
{
   /// <summary>Win32 API FindFirst()/FindNext()を使用してファイルシステムエントリ（ファイルやディレクトリ）を取得するクラス。</summary>
   [Serializable]
   internal sealed class FindFileSystemEntryInfo
   {
      #region Fields

      [NonSerialized] private static readonly Regex WildcardMatchAll = new Regex(@"^(\*)+(\.\*+)+$", RegexOptions.IgnoreCase); // *.* や *.** 等を認識する特殊ケース

      /// <summary>コンパイル済み正規表現キャッシュの最大エントリ数。上限に達するとキャッシュをクリアして無制限なメモリ増加を防止する。</summary>
      private const int RegexCacheMaxSize = 64;
      [NonSerialized] private static readonly ConcurrentDictionary<string, Regex> RegexCache = new ConcurrentDictionary<string, Regex>();
      [NonSerialized] private Regex _nameFilter;
      [NonSerialized] private string _searchPattern = Path.WildcardStarMatchAll;

      #endregion // Fields


      #region Private Static Methods

      /// <summary>検索パターンに対応するコンパイル済み正規表現をキャッシュから取得、または新規作成してキャッシュに追加する。
      /// キャッシュが<see cref="RegexCacheMaxSize"/>を超えた場合はクリアして無制限なメモリ増加を防止する。</summary>
      /// <param name="searchPattern">ワイルドカード検索パターン。</param>
      /// <returns>検索パターンに一致するコンパイル済み<see cref="Regex"/>インスタンス。</returns>
      private static Regex GetOrAddRegex(string searchPattern)
      {
         if (RegexCache.TryGetValue(searchPattern, out var cached))
            return cached;

         // キャッシュが上限に達した場合はクリアして、古いCompiledパターンがメモリに留まり続けるのを防ぐ。
         if (RegexCache.Count >= RegexCacheMaxSize)
            RegexCache.Clear();

         var regex = new Regex(
            string.Format(CultureInfo.InvariantCulture, "^{0}$", Regex.Escape(searchPattern).Replace(@"\*", ".*").Replace(@"\?", ".")),
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

         return RegexCache.GetOrAdd(searchPattern, regex);
      }

      #endregion // Private Static Methods


      #region Constructor

      /// <summary><see cref="FindFileSystemEntryInfo"/>クラスの新しいインスタンスを初期化します。</summary>
      /// <param name="transaction">使用する場合のNTFSカーネルトランザクション。</param>
      /// <param name="isFolder"><c>true</c>に設定するとパスはフォルダです。</param>
      /// <param name="path">パス。</param>
      /// <param name="searchPattern">ワイルドカード検索パターン。</param>
      /// <param name="options">列挙オプション。</param>
      /// <param name="customFilters">カスタムフィルター。</param>
      /// <param name="pathFormat">パスの形式。</param>
      /// <param name="typeOfT">取得するオブジェクトの型。</param>
      [SuppressMessage("Microsoft.Maintainability", "CA1502:AvoidExcessiveComplexity")]
      public FindFileSystemEntryInfo(KernelTransaction transaction, bool isFolder, string path, string searchPattern, DirectoryEnumerationOptions? options, DirectoryEnumerationFilters customFilters, PathFormat pathFormat, Type typeOfT)
      {
         if (null == options)
         {
            throw new ArgumentNullException("options");
         }


         Transaction = transaction;

         OriginalInputPath = path;

         InputPath = Path.GetExtendedLengthPathCore(transaction, path, pathFormat, GetFullPathOptions.RemoveTrailingDirectorySeparator | GetFullPathOptions.FullCheck);

         IsRelativePath = !Path.IsPathRooted(OriginalInputPath, false);

         RelativeAbsolutePrefix = IsRelativePath ? InputPath.Replace(OriginalInputPath, string.Empty) : null;
         
         SearchPattern = searchPattern.TrimEnd(Path.TrimEndChars); // .NET behaviour.

         FileSystemObjectType = null;

         ContinueOnException = (options & DirectoryEnumerationOptions.ContinueOnException) != 0;

         AsLongPath = (options & DirectoryEnumerationOptions.AsLongPath) != 0;

         AsString = typeOfT == typeof(string);
         AsFileSystemInfo = !AsString && (typeOfT == typeof(FileSystemInfo) || typeOfT.BaseType == typeof(FileSystemInfo));

         LargeCache = (options & DirectoryEnumerationOptions.LargeCache) != 0 ? NativeMethods.UseLargeCache : NativeMethods.FIND_FIRST_EX_FLAGS.NONE;

         // FileSystemEntryInfoのみが(8.3形式)AlternateFileNameを使用します。
         FindExInfoLevel = AsString || AsFileSystemInfo || (options & DirectoryEnumerationOptions.BasicSearch) != 0 ? NativeMethods.FindexInfoLevel : NativeMethods.FINDEX_INFO_LEVELS.Standard;


         if (null != customFilters)
         {
            InclusionFilter = customFilters.InclusionFilter;

            RecursionFilter = customFilters.RecursionFilter;

            ErrorHandler = customFilters.ErrorFilter;

#if !NET35
            CancellationToken = customFilters.CancellationToken;
#endif
         }


         if (isFolder)
         {
            IsDirectory = true;

            Recursive = (options & DirectoryEnumerationOptions.Recursive) != 0 || null != RecursionFilter;

            SkipReparsePoints = (options & DirectoryEnumerationOptions.SkipReparsePoints) != 0;


            // 列挙するためにフォルダまたはファイルが必要です。
            if ((options & DirectoryEnumerationOptions.FilesAndFolders) == 0)
            {
               options |= DirectoryEnumerationOptions.FilesAndFolders;
            }
         }

         else
         {
            options &= ~DirectoryEnumerationOptions.Folders; // Remove enumeration of folders.
            options |= DirectoryEnumerationOptions.Files; // Add enumeration of files.
         }


         FileSystemObjectType = (options & DirectoryEnumerationOptions.FilesAndFolders) == DirectoryEnumerationOptions.FilesAndFolders

            // フォルダとファイル(null)。
            ? (bool?) null

            // フォルダのみ(true)またはファイルのみ(false)。
            : (options & DirectoryEnumerationOptions.Folders) != 0;
      }

      #endregion // Constructor


      #region Properties

      /// <summary>オブジェクトを<see cref="FileSystemInfo"/>インスタンスとして返す機能を取得または設定します。</summary>
      /// <value><c>true</c>の場合、オブジェクトを<see cref="FileSystemInfo"/>インスタンスとして返します。</value>
      public bool AsFileSystemInfo { get; private set; }


      /// <summary>長いフルパス形式でフルパスを返す機能を取得または設定します。</summary>
      /// <value><c>true</c>の場合は長いフルパス形式で、<c>false</c>の場合は通常のパス形式でフルパスを返します。</value>
      public bool AsLongPath { get; private set; }


      /// <summary>オブジェクトインスタンスを<see cref="string"/>として返す機能を取得または設定します。</summary>
      /// <value><c>true</c>の場合、オブジェクトのフルパスを<see cref="string"/>として返します。</value>
      public bool AsString { get; private set; }


      /// <summary>アクセスエラー時にスキップする機能を取得または設定します。</summary>
      /// <value><c>true</c>の場合、ACL保護されたディレクトリやアクセスできないリパースポイントなどの障害による例外を抑制します。</value>
      public bool ContinueOnException { get; private set; }


      /// <summary>ファイルシステムオブジェクトの種類を取得します。</summary>
      /// <value>
      /// <c>null</c> = ファイルとディレクトリを返します。
      /// <c>true</c> = ディレクトリのみを返します。
      /// <c>false</c> = ファイルのみを返します。
      /// </value>
      public bool? FileSystemObjectType { get; private set; }


      /// <summary>パスが絶対パスか相対パスかを取得または設定します。</summary>
      /// <value>指定されたパス文字列が絶対パス情報または相対パス情報を含むかどうかを示す値を取得します。</value>
      public bool IsRelativePath { get; private set; }

      
      /// <summary>フォルダへの初期パスを取得または設定します。</summary>
      /// <value>長いパス形式でのファイルまたはフォルダへの初期パス。</value>
      public string OriginalInputPath { get; private set; }


      /// <summary>フォルダへのパスを取得または設定します。</summary>
      /// <value>長いパス形式でのファイルまたはフォルダへのパス。</value>
      public string InputPath { get; private set; }


      /// <summary>相対パスの絶対フルパスプレフィックスを取得または設定します。</summary>
      private string RelativeAbsolutePrefix { get; set; }

      
      /// <summary>使用する<see cref="NativeMethods.FINDEX_INFO_LEVELS"/>を示す値を取得または設定します。</summary>
      /// <value><c>true</c>はフォルダオブジェクト、<c>false</c>はファイルオブジェクトを示します。</value>
      public bool IsDirectory { get; private set; }


      /// <summary>ディレクトリクエリに大きなバッファを使用し、検索操作のパフォーマンスを向上させます。</summary>
      /// <remarks>この値はWindows Server 2008 R2およびWindows 7以降でサポートされています。</remarks>
      public NativeMethods.FIND_FIRST_EX_FLAGS LargeCache { get; private set; }


      /// <summary>FindFirstFileEx関数は短いファイル名をクエリしないため、全体的な列挙速度が向上します。</summary>
      /// <remarks>この値はWindows Server 2008 R2およびWindows 7以降でサポートされています。</remarks>
      public NativeMethods.FINDEX_INFO_LEVELS FindExInfoLevel { get; private set; }


      /// <summary>検索が現在のディレクトリのみを含むか、すべてのサブディレクトリを含むかを指定します。</summary>
      /// <value>すべてのサブディレクトリを含む場合は<c>true</c>。</value>
      public bool Recursive { get; private set; }


      /// <summary>パターンを使用してファイルシステムオブジェクト名を検索します。</summary>
      /// <value>ワイルドカード文字を含むパス。例えば、アスタリスク（<see cref="Path.WildcardStarMatchAll"/>）やクエスチョンマーク（<see cref="Path.WildcardQuestion"/>）。</value>
      [SuppressMessage("Microsoft.Usage", "CA2208:InstantiateArgumentExceptionsCorrectly")]
      public string SearchPattern
      {
         get { return _searchPattern; }

         internal set
         {
            if (null == value)
            {
               throw new ArgumentNullException("SearchPattern");
            }

            _searchPattern = value;

            _nameFilter = _searchPattern == Path.WildcardStarMatchAll || WildcardMatchAll.IsMatch(_searchPattern)
               ? null
               : GetOrAddRegex(_searchPattern);
         }
      }


      /// <summary><c>true</c>の場合リパースポイントをスキップし、<c>false</c>の場合リパースポイントをたどります。</summary>
      public bool SkipReparsePoints { get; private set; }


      /// <summary>KernelTransactionインスタンスを取得または設定します。</summary>
      /// <value>トランザクション。</value>
      public KernelTransaction Transaction { get; private set; }


      /// <summary>カスタム列挙の包含/除外フィルターを取得または設定します。</summary>
      /// <value>オブジェクトを出力に含めるか除外するかを決定するメソッド。</value>
      public Predicate<FileSystemEntryInfo> InclusionFilter { get; private set; }


      /// <summary>カスタム列挙の再帰フィルターを取得または設定します。</summary>
      /// <value>ディレクトリを再帰的にトラバースするかどうかを決定するメソッド。</value>
      public Predicate<FileSystemEntryInfo> RecursionFilter { get; private set; }


      /// <summary>発生する可能性のあるエラーのハンドラーを取得または設定します。</summary>
      /// <value>エラーハンドラーメソッド。</value>
      public ErrorHandler ErrorHandler { get; private set; }


#if !NET35
      /// <summary>列挙を中止するためのキャンセルトークンを取得または設定します。</summary>
      /// <value><see cref="CancellationToken"/>インスタンス。</value>
      private CancellationToken CancellationToken { get; set; }
#endif

      #endregion // Properties


      #region Methods

      private SafeFindFileHandle FindFirstFile(string pathLp, out NativeMethods.WIN32_FIND_DATA win32FindData, out int lastError, bool suppressException = false)
      {
         lastError = (int) Win32Errors.NO_ERROR;

         var searchOption = null != FileSystemObjectType && (bool) FileSystemObjectType ? NativeMethods.FINDEX_SEARCH_OPS.SearchLimitToDirectories : NativeMethods.FINDEX_SEARCH_OPS.SearchNameMatch;

         var handle = FileSystemInfo.FindFirstFileNative(Transaction, pathLp, FindExInfoLevel, searchOption, LargeCache, out lastError, out win32FindData);


         if (!suppressException && !ContinueOnException)
         {
            if (null == handle)
            {
               switch ((uint) lastError)
               {
                  case Win32Errors.ERROR_FILE_NOT_FOUND: // FileNotFoundException.
                  case Win32Errors.ERROR_PATH_NOT_FOUND: // DirectoryNotFoundException.
                  case Win32Errors.ERROR_NOT_READY:      // DeviceNotReadyException: Floppy device or network drive not ready.

                     Directory.ExistsDriveOrFolderOrFile(Transaction, pathLp, IsDirectory, lastError, true, true);
                     break;
               }


               ThrowPossibleException((uint) lastError, pathLp);
            }
         }


         return handle;
      }


      private FileSystemEntryInfo NewFilesystemEntry(string pathLp, string fileName, NativeMethods.WIN32_FIND_DATA win32FindData)
      {
         var fullPath = (IsRelativePath ? pathLp.Replace(RelativeAbsolutePrefix, string.Empty) : pathLp) + fileName;

         return new FileSystemEntryInfo(win32FindData) {FullPath = fullPath};
      }


      private T NewFileSystemEntryType<T>(bool isFolder, FileSystemEntryInfo fsei, string fileName, string pathLp, NativeMethods.WIN32_FIND_DATA win32FindData)
      {
         // yield（返却）を決定します。例えば、フォルダのみが要求された場合はファイルを返さず、その逆も同様です。

         if (null != FileSystemObjectType && (!(bool) FileSystemObjectType || !isFolder) && (!(bool) !FileSystemObjectType || isFolder))

         {
            return (T) (object) null;
         }


         // 名前フィルタリングからyield（返却）を決定します。

         if (null != fileName && !(null == _nameFilter || null != _nameFilter && _nameFilter.IsMatch(fileName)))

         {
            return (T) (object) null;
         }


         if (null == fsei)
         {
            fsei = NewFilesystemEntry(pathLp, fileName, win32FindData);
         }


         // オブジェクトインスタンスのFullPathプロパティを文字列として返します。オプションで長いパス形式で返します。

         return AsString ? null == InclusionFilter || InclusionFilter(fsei) ? (T) (object) (AsLongPath ? fsei.LongFullPath : fsei.FullPath) : (T) (object) null


            // 要求されたファイルシステムオブジェクトの種類が返されることを確認します。
            // null = ファイルとディレクトリを返します。
            // true = ディレクトリのみを返します。
            // false = ファイルのみを返します。

            : null != InclusionFilter && !InclusionFilter(fsei)
               ? (T) (object) null

               // FileSystemInfo型のオブジェクトインスタンスを返します。

               : AsFileSystemInfo
                  ? (T) (object) (fsei.IsDirectory

                     ? (FileSystemInfo) new DirectoryInfo(Transaction, fsei.LongFullPath, PathFormat.LongFullPath) {EntryInfo = fsei}

                     : new FileInfo(Transaction, fsei.LongFullPath, PathFormat.LongFullPath) {EntryInfo = fsei})

                  // FileSystemEntryInfo型のオブジェクトインスタンスを返します。

                  : (T) (object) fsei;
      }


      private void ThrowPossibleException(uint lastError, string pathLp)
      {
         switch (lastError)
         {
            case Win32Errors.ERROR_NO_MORE_FILES:
               lastError = Win32Errors.NO_ERROR;
               break;


            case Win32Errors.ERROR_FILE_NOT_FOUND: // On files.
            case Win32Errors.ERROR_PATH_NOT_FOUND: // On folders.
            case Win32Errors.ERROR_NOT_READY:      // DeviceNotReadyException: Floppy device or network drive not ready.
               // MSDN: .NET 3.5+: DirectoryNotFoundException: Path is invalid, such as referring to an unmapped drive.
               // Directory.Delete()

               lastError = IsDirectory ? (int) Win32Errors.ERROR_PATH_NOT_FOUND : Win32Errors.ERROR_FILE_NOT_FOUND;
               break;


            //case Win32Errors.ERROR_DIRECTORY:
            //   // MSDN: .NET 3.5+: IOException: path is a file name.
            //   // Directory.EnumerateDirectories()
            //   // Directory.EnumerateFiles()
            //   // Directory.EnumerateFileSystemEntries()
            //   // Directory.GetDirectories()
            //   // Directory.GetFiles()
            //   // Directory.GetFileSystemEntries()
            //   break;

            //case Win32Errors.ERROR_ACCESS_DENIED:
            //   // MSDN: .NET 3.5+: UnauthorizedAccessException: The caller does not have the required permission.
            //   break;
         }


         if (lastError != Win32Errors.NO_ERROR)
         {
            var regularPath = Path.GetCleanExceptionPath(pathLp);

            // ErrorHandlerが設定されている場合、制御を渡します。

            if (null == ErrorHandler || !ErrorHandler((int) lastError, new Win32Exception((int) lastError).Message, regularPath))
            {
               // ErrorHandlerがfalseを返した場合、例外をスローします。

               NativeError.ThrowException(lastError, regularPath);
            }
         }
      }


      private void VerifyInstanceType(NativeMethods.WIN32_FIND_DATA win32FindData)
      {
         var regularPath = Path.GetCleanExceptionPath(InputPath);

         var isFolder = File.IsDirectory(win32FindData.dwFileAttributes);

         if (IsDirectory)
         {
            if (!isFolder)
            {
               throw new DirectoryNotFoundException(string.Format(CultureInfo.InvariantCulture, "({0}) {1}", Win32Errors.ERROR_PATH_NOT_FOUND, string.Format(CultureInfo.InvariantCulture, Resources.Target_Directory_Is_A_File, regularPath)));
            }
         }

         else if (isFolder)
         {
            throw new FileNotFoundException(string.Format(CultureInfo.InvariantCulture, "({0}) {1}", Win32Errors.ERROR_FILE_NOT_FOUND, string.Format(CultureInfo.InvariantCulture, Resources.Target_File_Is_A_Directory, regularPath)));
         }
      }


      /// <summary>検索対象のディレクトリ内のワイルドカードとカスタム述語の両方に一致するすべてのファイルシステムオブジェクトを返す列挙子を取得します。</summary>
      /// <returns><see cref="IEnumerable{T}"/>インスタンス: FileSystemEntryInfo、DirectoryInfo、FileInfo、またはstring（フルパス）。</returns>
      [SecurityCritical]
      public IEnumerable<T> Enumerate<T>()
      {
         // MSDN: Queue
         // Represents a first-in, first-out collection of objects.
         // The capacity of a Queue is the number of elements the Queue can hold.
         // As elements are added to a Queue, the capacity is automatically increased as required through reallocation. The capacity can be decreased by calling TrimToSize.
         // The growth factor is the number by which the current capacity is multiplied when a greater capacity is required. The growth factor is determined when the Queue is constructed.
         // The capacity of the Queue will always increase by a minimum value, regardless of the growth factor; a growth factor of 1.0 will not prevent the Queue from increasing in size.
         // If the size of the collection can be estimated, specifying the initial capacity eliminates the need to perform a number of resizing operations while adding elements to the Queue.
         // This constructor is an O(n) operation, where n is capacity.

         var dirs = new Queue<string>(NativeMethods.DefaultDirectoryQueueCapacity);

         dirs.Enqueue(Path.AddTrailingDirectorySeparator(InputPath, false));


         using (new NativeMethods.ChangeErrorMode(NativeMethods.ErrorMode.FailCriticalErrors))

            while (dirs.Count > 0
#if !NET35
               && !CancellationToken.IsCancellationRequested
#endif
            )
            {


               // キューの先頭からオブジェクトを削除します。
               // このアルゴリズムの計算量はO(1)です。要素をループしません。

               var pathLp = dirs.Dequeue();


               using var handle = FindFirstFile(pathLp + Path.WildcardStarMatchAll, out var win32FindData, out var lastError);
               // ハンドルがnullでここにいるということは、ErrorHandlerがアクティブであることを意味します。
               // アクセスできないフォルダに遭遇したため、中断して次のフォルダに進みます。
               if (null == handle)
               {
                  continue;
               }

               do
               {
                  if (lastError == (int) Win32Errors.ERROR_NO_MORE_FILES)
                  {
                     lastError = (int) Win32Errors.NO_ERROR;
                     continue;
                  }


                  // リパースポイントをスキップして、通常のディレクトリとリンクをきれいに分離します。
                  if (SkipReparsePoints && (win32FindData.dwFileAttributes & FileAttributes.ReparsePoint) != 0)
                  {
                     continue;
                  }


                  var fileName = win32FindData.cFileName;

                  var isFolder = (win32FindData.dwFileAttributes & FileAttributes.Directory) != 0;

                  // ".."と"."エントリをスキップします。
                  if (isFolder && (fileName.Equals(Path.ParentDirectoryPrefix, StringComparison.Ordinal) || fileName.Equals(Path.CurrentDirectoryPrefix, StringComparison.Ordinal)))
                  {
                     continue;
                  }


                  // 事前フィルタリング: WIN32_FIND_DATA から直接判定してオブジェクト生成を回避
                  // フォルダは再帰キューイングが必要なためスキップしない
                  if (!isFolder)
                  {
                     // フォルダのみ要求なのにファイル → スキップ
                     if (null != FileSystemObjectType && (bool) FileSystemObjectType)
                     {
                        continue;
                     }

                     // _nameFilter にマッチしないファイル → オブジェクト生成をスキップ
                     if (null != _nameFilter && !_nameFilter.IsMatch(fileName))
                     {
                        continue;
                     }
                  }


                  var fsei = NewFilesystemEntry(pathLp, fileName, win32FindData);

                  var res = NewFileSystemEntryType<T>(isFolder, fsei, fileName, pathLp, win32FindData);


                  // 再帰が要求された場合、後でトラバースするためにキューに追加します。
                  if (isFolder && Recursive && (null == RecursionFilter || RecursionFilter(fsei)))

                  {
                     dirs.Enqueue(Path.AddTrailingDirectorySeparator(pathLp + fileName, false));
                  }


                  // Codacy: ジェネリック型パラメータを参照型に制限する制約が適用されていない場合、structなどの値型も渡される可能性があります。
                  // そのような場合、型パラメータとnullの比較は常にfalseになります。
                  // structは空にはなれますが、nullにはなれないためです。値型が本当に期待されるものであれば、比較にはdefault()を使用するべきです。
                  // そうでなければ、値型が渡せないように制約を追加するべきです。

                  if (Equals(res, default(T)))
                  {
                     continue;
                  }


                  yield return res;

               } while (
#if !NET35
                  !CancellationToken.IsCancellationRequested &&
#endif
                  NativeMethods.FindNextFile(handle, out win32FindData));


               lastError = Marshal.GetLastWin32Error();

               if (!ContinueOnException
#if !NET35
                   && !CancellationToken.IsCancellationRequested
#endif
                  )
               {
                  ThrowPossibleException((uint) lastError, pathLp);
               }
            }
      }


      /// <summary>特定のファイルシステムオブジェクトを取得します。</summary>
      /// <returns>
      /// <para>戻り値の型はC#の型推論に基づきます。可能な戻り値の型:</para>
      /// <para> <see cref="string"/>（フルパス）、<see cref="FileSystemInfo"/>（<see cref="DirectoryInfo"/>または<see cref="FileInfo"/>）、<see cref="FileSystemEntryInfo"/>インスタンス</para>
      /// <para>または例外が発生し<see cref="ContinueOnException"/>が<c>true</c>の場合はnull。</para>
      /// </returns>
      [SecurityCritical]
      public T Get<T>()
      {
         using (new NativeMethods.ChangeErrorMode(NativeMethods.ErrorMode.FailCriticalErrors))
         {
            NativeMethods.WIN32_FIND_DATA win32FindData;
            var lastError = 0;

            // フォルダとして明示的に設定されていない場合。

            if (!IsDirectory)
            {
               using var handle = FindFirstFile(InputPath, out win32FindData, out lastError);
               if (null != handle)
               {
                  if (!ContinueOnException)
                  {
                     VerifyInstanceType(win32FindData);
                  }
               }

               else
               {
                  return (T) (object) null;
               }


               return NewFileSystemEntryType<T>((win32FindData.dwFileAttributes & FileAttributes.Directory) != 0, null, null, InputPath, win32FindData);
            }


            using (var handle = FindFirstFile(InputPath, out win32FindData, out lastError, true))
            {
               if (null == handle)
               {
                  // InputPathは論理ドライブ（例: "C:\"、"D:\"）の可能性があります。

                  var attrs = new NativeMethods.WIN32_FILE_ATTRIBUTE_DATA();

                  lastError = File.FillAttributeInfoCore(Transaction, Path.GetRegularPathCore(InputPath, GetFullPathOptions.None, false), ref attrs, false, true);

                  if (lastError != Win32Errors.NO_ERROR)
                  {
                     if (!ContinueOnException)
                     {
                        switch ((uint) lastError)
                        {
                           case Win32Errors.ERROR_FILE_NOT_FOUND: // FileNotFoundException.
                           case Win32Errors.ERROR_PATH_NOT_FOUND: // DirectoryNotFoundException.
                           case Win32Errors.ERROR_NOT_READY:      // DeviceNotReadyException: Floppy device or network drive not ready.
                           case Win32Errors.ERROR_BAD_NET_NAME:

                              Directory.ExistsDriveOrFolderOrFile(Transaction, InputPath, IsDirectory, lastError, true, true);
                              break;
                        }

                        ThrowPossibleException((uint) lastError, InputPath);
                     }

                     return (T) (object) null;
                  }


                  win32FindData = new NativeMethods.WIN32_FIND_DATA
                  {
                     cFileName = Path.CurrentDirectoryPrefix,
                     dwFileAttributes = attrs.dwFileAttributes,
                     ftCreationTime = attrs.ftCreationTime,
                     ftLastAccessTime = attrs.ftLastAccessTime,
                     ftLastWriteTime = attrs.ftLastWriteTime,
                     nFileSizeHigh = attrs.nFileSizeHigh,
                     nFileSizeLow = attrs.nFileSizeLow
                  };
               }


               if (!ContinueOnException)
               {
                  VerifyInstanceType(win32FindData);
               }
            }


            return NewFileSystemEntryType<T>((win32FindData.dwFileAttributes & FileAttributes.Directory) != 0, null, null, InputPath, win32FindData);
         }
      }

      #endregion // Methods
   }
}
