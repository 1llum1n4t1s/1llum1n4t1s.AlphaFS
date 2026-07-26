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

namespace Alphaleonis.Win32
{
   /// <summary>netapi32.dll のネットワーク管理関数 (NetShareEnum / NetSessionEnum / NetFileEnum /
   /// NetConnectionEnum / NetServerDiskEnum / NetDfsEnum / NetShareGetInfo / NetStatisticsGet 等) が
   /// 確保して返すバッファを表します。
   /// <para>これらのバッファは Win32 の契約上 <c>NetApiBufferFree</c> で解放する必要があります。
   /// LocalFree (<see cref="System.Runtime.InteropServices.Marshal.FreeHGlobal"/>) での解放は、
   /// netapi32 が内部で使うヒープの実装詳細にたまたま一致しているだけで、契約としては未定義動作です。</para>
   /// </summary>
   internal sealed class SafeNetApiBufferHandle : SafeNativeMemoryBufferHandle
   {
      /// <summary>ゼロ IntPtr で <see cref="SafeNetApiBufferHandle"/> クラスの新しいインスタンスを初期化します。</summary>
      /// <remarks>P/Invoke の out パラメータとしてマーシャラーが生成するため、既定コンストラクタが必要です。</remarks>
      public SafeNetApiBufferHandle() : base(true)
      {
      }


      /// <summary>ハンドルを解放します。</summary>
      /// <returns>常に <c>true</c>。</returns>
      protected override bool ReleaseHandle()
      {
         // 戻り値は NERR_Success (0) か、無効なポインタなら ERROR_INVALID_PARAMETER。
         // ReleaseHandle からは例外を投げられないため、結果は握って true を返す
         // (false を返すと ReleaseHandleFailed MDA が発生するだけで復旧手段は無い)。
         Network.NativeMethods.NetApiBufferFree(handle);

         return true;
      }
   }
}
