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
using Microsoft.Win32.SafeHandles;

namespace Alphaleonis.Win32.Filesystem
{
   /// <summary>FindFirstFile/FindNextFile Win32 API関数で使用されるハンドルのラッパークラスを表します。</summary>
   [SecurityCritical]
   public sealed class SafeFindFileHandle : SafeHandleZeroOrMinusOneIsInvalid
   {
      /// <summary>このクラスのデフォルトインスタンスの作成を防止するコンストラクタ。</summary>
      private SafeFindFileHandle() : base(true)
      {
      }

      /// <summary><see cref="SafeFindFileHandle"/>クラスの新しいインスタンスを初期化します。</summary>
      /// <param name="handle">ハンドル。</param>
      /// <param name="callerHandle">ファイナライズ段階でハンドルを確実に解放する場合は<c>true</c>、確実な解放を防止する場合は<c>false</c>（非推奨）。</param>
      public SafeFindFileHandle(IntPtr handle, bool callerHandle) : base(callerHandle)
      {
         SetHandle(handle);
      }

      
      /// <summary>派生クラスでオーバーライドされた場合、ハンドルを解放するために必要なコードを実行します。</summary>
      /// <returns>
      /// ハンドルが正常に解放された場合は<c>true</c>、致命的な障害が発生した場合は
      /// <c>false</c>。この場合、ReleaseHandleFailed マネージデバッグアシスタントが生成されます。
      /// </returns>
      protected override bool ReleaseHandle()
      {
         return NativeMethods.FindClose(handle);
      }
   }
}
