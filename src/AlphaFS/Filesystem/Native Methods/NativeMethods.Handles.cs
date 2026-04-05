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
   internal static partial class NativeMethods
   {
      /// <summary>開いているオブジェクトハンドルを閉じます。</summary>
      /// <remarks>
      ///   <para>CloseHandle 関数は以下のオブジェクトのハンドルを閉じます:</para>
      ///   <para>アクセストークン、通信デバイス、コンソール入力、コンソールスクリーンバッファ、イベント、ファイル、ファイルマッピング、I/O 完了ポート、
      ///   ジョブ、メールスロット、メモリリソース通知、ミューテックス、名前付きパイプ、パイプ、プロセス、セマフォ、スレッド、トランザクション、
      ///   待機可能タイマー。</para>
      ///   <para>SetLastError は <c>false</c> に設定されています。</para>
      ///   <para>サポートされる最小クライアント: Windows 2000 Professional [デスクトップアプリ | Windows ストアアプリ]</para>
      ///   <para>サポートされる最小サーバー: Windows 2000 Server [デスクトップアプリ | Windows ストアアプリ]</para>
      /// </remarks>
      /// <returns>
      ///   <para>関数が成功した場合、戻り値はゼロ以外です。</para>
      ///   <para>関数が失敗した場合、戻り値はゼロです。拡張エラー情報を取得するには GetLastError を呼び出してください。</para>
      ///   <para>アプリケーションがデバッガーの下で実行されている場合、無効なハンドル値または疑似ハンドル値を受け取ると例外をスローします。
      ///   これは、ハンドルを二重に閉じた場合、または FindClose 関数を呼び出す代わりに FindFirstFile 関数が返したハンドルに対して
      ///   CloseHandle を呼び出した場合に発生する可能性があります。</para>
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = false, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool CloseHandle(IntPtr hObject);
   }
}
