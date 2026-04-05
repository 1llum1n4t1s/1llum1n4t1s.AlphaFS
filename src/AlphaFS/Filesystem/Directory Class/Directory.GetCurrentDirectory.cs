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

using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class Directory
   {
      /// <summary>
      /// アプリケーションの現在の作業ディレクトリを取得します。
      /// <para>
      ///   MSDN: マルチスレッドアプリケーションと共有ライブラリコードはGetCurrentDirectory関数を使用すべきではありません。 should avoid using relative path names.
      ///   SetCurrentDirectory関数によって書き込まれる現在のディレクトリ状態は各プロセスのグローバル変数として格納され、
      ///   そのためマルチスレッドアプリケーションは他のスレッドからのデータ破損の可能性なくこの値を信頼できません。 that may also be reading or setting this value.
      ///   <para>This limitation also applies to the SetCurrentDirectory and GetFullPathName functions. The exception being when the application is guaranteed to be running in a single thread,
      ///   例えば、追加のスレッドを作成する前にメインスレッドでコマンドライン引数文字列からファイル名を解析するなど。</para>
      ///   <para>マルチスレッドアプリケーションや共有ライブラリコードで相対パス名を使用すると、予測不可能な結果が生じる可能性があり、サポートされていません。</para>
      /// </para>
      /// </summary>
      /// <returns>The path of the current working directory without a trailing directory separator.</returns>
      [SuppressMessage("Microsoft.Design", "CA1024:UsePropertiesWhereAppropriate"), SecurityCritical]
      public static string GetCurrentDirectory()
      {
         var nameBuffer = new StringBuilder(NativeMethods.MaxPathUnicode);

         // GetCurrentDirectory()
         // 2016-09-29: MSDNはLongPathの使用を確認していないが、この関数のUnicodeバージョンが存在する。
         // 2017-05-30: MSDN confirms LongPath usage: Starting with Windows 10, version 1607
         // 2018-01-15: MSDN confirmation is gone?

         var folderNameLength = NativeMethods.GetCurrentDirectory((uint) nameBuffer.Capacity, nameBuffer);
         var lastError = Marshal.GetLastWin32Error();

         if (folderNameLength == 0)
         {
            NativeError.ThrowException(lastError);
         }

         if (folderNameLength > NativeMethods.MaxPathUnicode)
         {
            throw new PathTooLongException(string.Format(CultureInfo.InvariantCulture, "Path is greater than {0} characters: {1}", NativeMethods.MaxPathUnicode, folderNameLength));
         }

         return nameBuffer.ToString();
      }
   }
}
