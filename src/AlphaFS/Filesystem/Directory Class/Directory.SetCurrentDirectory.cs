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
using System.Runtime.InteropServices;
using System.Security;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class Directory
   {
      /// <summary>
      /// アプリケーションの's current working directory to the specified directory.
      /// <para>
      ///   MSDN: マルチスレッドアプリケーションと共有ライブラリコードはGetCurrentDirectory関数を使用すべきではありません。 should avoid using relative path names.
      ///   SetCurrentDirectory関数によって書き込まれる現在のディレクトリ状態は各プロセスのグローバル変数として格納され、
      ///   そのためマルチスレッドアプリケーションは他のスレッドからのデータ破損の可能性なくこの値を信頼できません。 that may also be reading or setting this value.
      ///   <para>This limitation also applies to the SetCurrentDirectory and GetFullPathName functions. The exception being when the application is guaranteed to be running in a single thread,
      ///   例えば、追加のスレッドを作成する前にメインスレッドでコマンドライン引数文字列からファイル名を解析するなど。</para>
      ///   <para>マルチスレッドアプリケーションや共有ライブラリコードで相対パス名を使用すると、予測不可能な結果が生じる可能性があり、サポートされていません。</para>
      /// </para>
      /// </summary>
      /// <param name="path">The path to which the current working directory is set.</param>
      [SuppressMessage("Microsoft.Design", "CA1062:Validate arguments of public methods", MessageId = "0", Justification = "Utils.IsNullOrWhiteSpace validates arguments.")]
      [SecurityCritical]
      public static void SetCurrentDirectory(string path)
      {
         SetCurrentDirectory(path, PathFormat.RelativePath);
      }


      /// <summary>
      /// アプリケーションの's current working directory to the specified directory.
      /// <para>
      ///   MSDN: マルチスレッドアプリケーションと共有ライブラリコードはGetCurrentDirectory関数を使用すべきではありません。 should avoid using relative path names.
      ///   SetCurrentDirectory関数によって書き込まれる現在のディレクトリ状態は各プロセスのグローバル変数として格納され、
      ///   そのためマルチスレッドアプリケーションは他のスレッドからのデータ破損の可能性なくこの値を信頼できません。 that may also be reading or setting this value.
      ///   <para>This limitation also applies to the SetCurrentDirectory and GetFullPathName functions. The exception being when the application is guaranteed to be running in a single thread,
      ///   例えば、追加のスレッドを作成する前にメインスレッドでコマンドライン引数文字列からファイル名を解析するなど。</para>
      ///   <para>マルチスレッドアプリケーションや共有ライブラリコードで相対パス名を使用すると、予測不可能な結果が生じる可能性があり、サポートされていません。</para>
      /// </para>
      /// </summary>
      /// <param name="path">The path to which the current working directory is set.</param>
      /// <param name="pathFormat">Indicates the format of the path parameter.</param>
      [SuppressMessage("Microsoft.Design", "CA1062:Validate arguments of public methods", MessageId = "0", Justification = "Utils.IsNullOrWhiteSpace validates arguments.")]
      [SecurityCritical]
      public static void SetCurrentDirectory(string path, PathFormat pathFormat)
      {
         if (Utils.IsNullOrWhiteSpace(path))
         {
            throw new ArgumentNullException("path");
         }

         var fullCheck = pathFormat == PathFormat.RelativePath;
         Path.CheckSupportedPathFormat(path, fullCheck, fullCheck);
         var pathLp = Path.GetExtendedLengthPathCore(null, path, pathFormat, GetFullPathOptions.AddTrailingDirectorySeparator);

         if (pathFormat == PathFormat.FullPath)
         {
            pathLp = Path.GetRegularPathCore(pathLp, GetFullPathOptions.None, false);
         }


         // SetCurrentDirectory()
         // 2016-09-29: MSDNはLongPathの使用を確認していないが、この関数のUnicodeバージョンが存在する。
         // 2017-05-30: MSDN confirms LongPath usage: Starting with Windows 10, version 1607

         var success = NativeMethods.SetCurrentDirectory(pathLp);

         var lastError = Marshal.GetLastWin32Error();
         if (!success)
         {
            NativeError.ThrowException(lastError, pathLp);
         }
      }
   }
}
