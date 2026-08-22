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
using System.Security;
using System.Text;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class Path
   {
      /// <summary>指定された <paramref name="path"/> 文字列の絶対パスを取得します。</summary>
      /// <returns>"C:\MyFile.txt" のような <paramref name="path"/> の完全修飾パス。</returns>
      /// <remarks>
      /// <para>GetFullPathName は、現在のドライブとディレクトリの名前を指定されたファイル名と結合して、指定されたファイルの完全パスとファイル名を決定します。</para>
      /// <para>また、完全パスとファイル名のファイル名部分のアドレスも計算します。</para>
      /// <para>&#160;</para>
      /// <para>このメソッドは、結果のパスとファイル名が有効であるか、関連するボリューム上に既存のファイルが存在するかを検証しません。</para>
      /// <para>.NET Framework は、<c>\\.\PhysicalDrive0</c> のようなデバイス名のパスを通じた物理ディスクへの直接アクセスをサポートしていません。</para>
      /// <para>&#160;</para>
      /// <para>MSDN: マルチスレッドアプリケーションと共有ライブラリコードは GetFullPathName 関数を使用すべきではなく、</para>
      /// <para>相対パス名の使用を避けるべきです。SetCurrentDirectory 関数によって書き込まれるカレントディレクトリの状態は、各プロセスのグローバル変数として格納されるため、</para>
      /// <para>マルチスレッドアプリケーションは、この値を読み取りまたは設定している他のスレッドからのデータ破損の可能性なしにこの値を確実に使用することはできません。</para>
      /// <para>この制限は SetCurrentDirectory および GetCurrentDirectory 関数にも適用されます。例外は、アプリケーションが単一スレッドで実行されることが保証されている場合です。</para>
      /// <para>例えば、追加のスレッドを作成する前にメインスレッドでコマンドライン引数文字列からファイル名を解析する場合です。</para>
      /// <para>マルチスレッドアプリケーションや共有ライブラリコードで相対パス名を使用すると、予測不可能な結果が生じる可能性があり、サポートされていません。</para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="checkSupported">サポートされているパス形式かチェックするかどうか。</param>
      /// <param name="path">絶対パス情報を取得するファイルまたはディレクトリ。</param>
      /// <param name="options">完全パス取得を制御するオプション。</param>
      [SecurityCritical]
      internal static string GetFullPathCore(KernelTransaction transaction, bool checkSupported, string path, GetFullPathOptions options)
      {
         // Windowsカーネルのみが認識する特殊パスをスキップする。

         if (null != path)
         {
            if (path.StartsWith(GlobalRootPrefix, StringComparison.OrdinalIgnoreCase) ||
                path.StartsWith(VolumePrefix, StringComparison.OrdinalIgnoreCase) ||
                path.StartsWith(NonInterpretedPathPrefix, StringComparison.Ordinal))

            {
               return path;
            }


            if (checkSupported)
            {
               CheckInvalidUncPath(path);

               CheckSupportedPathFormat(path, true, true);
            }
         }


         if (options != GetFullPathOptions.None)
         {
            if (!checkSupported && (options & GetFullPathOptions.CheckInvalidPathChars) != 0)
            {
               var checkAdditional = (options & GetFullPathOptions.CheckAdditional) != 0;

               CheckInvalidPathChars(path, checkAdditional, false);
               

               // 重複チェックを防止する。
               options &= ~GetFullPathOptions.CheckInvalidPathChars;

               if (checkAdditional)
               {
                  options &= ~GetFullPathOptions.CheckAdditional;
               }
            }

            // "C:\" のようなドライブを指すパスの場合、末尾のディレクトリ区切り文字を削除しない。
            // 削除するとパスがカレントディレクトリを指すことになる。

            // ".", "C:", "C:\"
            if (null == path || path.Length <= 3 || !path.StartsWith(LongPathPrefix, StringComparison.Ordinal) && path[1] != VolumeSeparatorChar)
            {
               options &= ~GetFullPathOptions.RemoveTrailingDirectorySeparator;
            }
         }
         

         var pathLp = GetLongPathCore(path, options);

         uint bufferSize = NativeMethods.MaxPath;  // 大多数のパスは260文字以下。不足時はgoto startGetFullPathNameで自動拡大

         using (new NativeMethods.ChangeErrorMode(NativeMethods.ErrorMode.FailCriticalErrors))
         {
         
      startGetFullPathName:

            var buffer = new StringBuilder((int) bufferSize);

            var returnLength = null == transaction || !NativeMethods.IsAtLeastWindowsVista

               // GetFullPathName() / GetFullPathNameTransacted()
               // 2013-04-15: MSDN confirms LongPath usage.

               ? NativeMethods.GetFullPathName(pathLp, bufferSize, buffer, IntPtr.Zero)

               : NativeMethods.GetFullPathNameTransacted(pathLp, bufferSize, buffer, IntPtr.Zero, transaction.SafeHandle);


            if (returnLength != Win32Errors.NO_ERROR)
            {
               if (returnLength > bufferSize)
               {
                  bufferSize = returnLength;
                  goto startGetFullPathName;
               }
            }

            else
            {
               var lastError = System.Runtime.InteropServices.Marshal.GetLastWin32Error();

               if ((options & GetFullPathOptions.ContinueOnNonExist) != 0)
               {
                  return null;
               }

               NativeError.ThrowException(lastError, pathLp);
            }


            var finalFullPath = (options & GetFullPathOptions.AsLongPath) != 0 ? GetLongPathCore(buffer.ToString(), GetFullPathOptions.None) : GetRegularPathCore(buffer.ToString(), GetFullPathOptions.None, false);
            
            finalFullPath = NormalizePath(finalFullPath, options);
            

            if ((options & GetFullPathOptions.KeepDotOrSpace) != 0)
            {
               if (pathLp.EndsWith(CurrentDirectoryPrefix, StringComparison.Ordinal))

               {
                  finalFullPath += CurrentDirectoryPrefix;
               }


               var lastChar = pathLp[pathLp.Length - 1];

               if (char.IsWhiteSpace(lastChar))

               {
                  finalFullPath += lastChar;
               }
            }


            return finalFullPath;
         }
      }


      /// <summary><seealso cref="GetFullPathOptions"/> を <paramref name="path"/> に適用します。</summary>
      /// <returns><paramref name="options"/> が適用された <paramref name="path"/>。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <param name="path">パス。</param>
      /// <param name="options">適用するオプション。</param>
      private static string ApplyFullPathOptions(string path, GetFullPathOptions options)
      {
         if ((options & GetFullPathOptions.TrimEnd) != 0)
         {
            if ((options & GetFullPathOptions.KeepDotOrSpace) == 0)
            {
               path = path.TrimEnd();
            }
         }

         if ((options & GetFullPathOptions.AddTrailingDirectorySeparator) != 0)
         {
            path = AddTrailingDirectorySeparator(path, false);
         }

         if ((options & GetFullPathOptions.RemoveTrailingDirectorySeparator) != 0)
         {
            path = RemoveTrailingDirectorySeparator(path);
         }

         if ((options & GetFullPathOptions.CheckInvalidPathChars) != 0)
         {
            CheckInvalidPathChars(path, (options & GetFullPathOptions.CheckAdditional) != 0, false);
         }


         // 先頭の空白をトリムする。
         if ((options & GetFullPathOptions.KeepDotOrSpace) == 0)
         {
            path = path.TrimStart();
         }

         return path;
      }
   }
}
