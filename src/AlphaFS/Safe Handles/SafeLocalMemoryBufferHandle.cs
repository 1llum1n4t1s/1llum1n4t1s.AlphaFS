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
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

namespace Alphaleonis.Win32.Security
{
   /// <summary>Marshal.AllocHGlobal操作の結果として使用できるIntPtrラッパー。
   /// <para>破棄またはファイナライズ時にMarshal.FreeHGlobalを呼び出します。</para>
   /// </summary>
   internal sealed class SafeLocalMemoryBufferHandle : SafeHandleZeroOrMinusOneIsInvalid
   {
      /// <summary>ゼロIntPtrで<see cref="SafeLocalMemoryBufferHandle"/>クラスの新しいインスタンスを初期化します。</summary>
      public SafeLocalMemoryBufferHandle() : base(true)
      {
      }


      /// <summary>1次元のマネージ8ビット符号なし整数配列からこのインスタンスが参照するアンマネージメモリポインターにデータをコピーします。</summary>
      /// <param name="source">コピー元の1次元配列。</param>
      /// <param name="startIndex">コピーを開始する配列のゼロベースインデックス。</param>
      /// <param name="length">コピーする配列要素の数。</param>
      public void CopyFrom(byte[] source, int startIndex, int length)
      {
         Marshal.Copy(source, startIndex, handle, length);
      }
      
      
      public void CopyTo(byte[] destination, int destinationOffset, int length)
      {
         if (destination == null)
         {
            throw new ArgumentNullException("destination");
         }

         if (destinationOffset < 0)
         {
            throw new ArgumentOutOfRangeException("destinationOffset", Resources.Negative_Destination_Offset);
         }

         if (length < 0)
         {
            throw new ArgumentOutOfRangeException("length", Resources.Negative_Length);
         }

         if (destinationOffset + length > destination.Length)
         {
            throw new ArgumentException(Resources.Destination_Buffer_Not_Large_Enough, "length");
         }

         Marshal.Copy(handle, destination, destinationOffset, length);
      }


      public byte[] ToByteArray(int startIndex, int length)
      {
         if (IsInvalid)
         {
            return null;
         }

         var arr = new byte[length];
         Marshal.Copy(new IntPtr(handle.ToInt64() + startIndex), arr, 0, length);
         return arr;
      }


      /// <summary>派生クラスでオーバーライドされた場合、ハンドルを解放するために必要なコードを実行します。</summary>
      /// <returns>
      /// ハンドルが正常に解放された場合は<c>true</c>、致命的な障害が発生した場合は
      /// <c>false</c>。この場合、ReleaseHandleFailed マネージデバッグアシスタントが生成されます。
      /// </returns>
      protected override bool ReleaseHandle()
      {
         return handle == IntPtr.Zero || NativeMethods.LocalFree(handle) == IntPtr.Zero;
      }
   }
}
