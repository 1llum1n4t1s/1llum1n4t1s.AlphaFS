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
using System.Text;

namespace Alphaleonis.Win32
{
   /// <summary>Kernel32.dllのLocalAlloc関数を使用して割り当てられた、指定されたサイズのネイティブメモリブロックを表します。</summary>
   internal sealed class SafeGlobalMemoryBufferHandle : SafeNativeMemoryBufferHandle
   {
      /// <summary>ゼロIntPtrで<see cref="SafeGlobalMemoryBufferHandle"/>クラスの新しいインスタンスを初期化します。</summary>
      public SafeGlobalMemoryBufferHandle() : base(true)
      {
      }


      /// <summary>指定されたバイト数のアンマネージメモリを割り当てて、<see cref="SafeGlobalMemoryBufferHandle"/>クラスの新しいインスタンスを初期化します。</summary>
      /// <param name="capacity">容量。</param>
      public SafeGlobalMemoryBufferHandle(int capacity) : base(capacity)
      {
         SetHandle(Marshal.AllocHGlobal(capacity));
      }


      private SafeGlobalMemoryBufferHandle(IntPtr buffer, int capacity) : base(buffer, capacity)
      {
      }


      [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope")]
      public static SafeGlobalMemoryBufferHandle FromLong(long? value)
      {
         if (value.HasValue)
         {
            var safeBuffer = new SafeGlobalMemoryBufferHandle(Marshal.SizeOf<long>());

            Marshal.WriteInt64(safeBuffer.handle, value.Value);

            return safeBuffer;
         }

         return new SafeGlobalMemoryBufferHandle();
      }


      public static SafeGlobalMemoryBufferHandle FromStringUni(string str)
      {
         if (str == null)
         {
            throw new ArgumentNullException("str");
         }

         return new SafeGlobalMemoryBufferHandle(Marshal.StringToHGlobalUni(str), str.Length * UnicodeEncoding.CharSize + UnicodeEncoding.CharSize);
      }


      /// <summary>派生クラスでオーバーライドされた場合、ハンドルを解放するために必要なコードを実行します。</summary>
      /// <returns>
      /// ハンドルが正常に解放された場合は<c>true</c>、致命的な障害が発生した場合は
      /// <c>false</c>。この場合、ReleaseHandleFailed マネージデバッグアシスタントが生成されます。
      /// </returns>
      protected override bool ReleaseHandle()
      {
         Marshal.FreeHGlobal(handle);
         return true;
      }
   }
}
