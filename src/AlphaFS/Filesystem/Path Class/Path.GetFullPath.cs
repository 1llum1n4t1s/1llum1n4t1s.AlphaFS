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

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class Path
   {
      #region .NET

      /// <summary>指定されたパス文字列の絶対パスを返します。</summary>
      /// <returns>"C:\MyFile.txt" のような完全修飾パス。</returns>
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
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="NotSupportedException"/>
      /// <param name="path">絶対パス情報を取得するファイルまたはディレクトリ。</param>
      [SecurityCritical]
      public static string GetFullPath(string path)
      {
         return GetFullPathCore(null, true, path, GetFullPathOptions.None);
      }

      #endregion // .NET


      /// <summary>[AlphaFS] 指定されたパス文字列の絶対パスを返します。</summary>
      /// <returns>"C:\MyFile.txt" のような完全修飾パス。</returns>
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
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="NotSupportedException"/>
      /// <param name="path">絶対パス情報を取得するファイルまたはディレクトリ。</param>
      /// <param name="options">完全パス取得を制御するオプション。</param>
      [SecurityCritical]
      public static string GetFullPath(string path, GetFullPathOptions options)
      {
         return GetFullPathCore(null, true, path, options);
      }
   }
}
