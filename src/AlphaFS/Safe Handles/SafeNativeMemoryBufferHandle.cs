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

using Microsoft.Win32.SafeHandles;
using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

namespace Alphaleonis.Win32
{
   /// <summary>アンマネージメモリのブロックを表すクラスの基底クラス。</summary>
   internal abstract class SafeNativeMemoryBufferHandle : SafeHandleZeroOrMinusOneIsInvalid
   {
      private readonly int m_capacity;


      /// <summary>メモリブロックの割り当て容量を指定して、<see cref="SafeNativeMemoryBufferHandle"/>クラスの新しいインスタンスを初期化します。</summary>
      /// <param name="callerHandle">ファイナライズ段階でハンドルを確実に解放する場合は<c>true</c>、確実な解放を防止する場合は<c>false</c>（非推奨）。</param>
      protected SafeNativeMemoryBufferHandle(bool callerHandle) : base(callerHandle)
      {
      }


      /// <summary>メモリブロックの割り当て容量を指定して、<see cref="SafeNativeMemoryBufferHandle"/>クラスの新しいインスタンスを初期化します。</summary>
      /// <param name="capacity">容量。</param>
      protected SafeNativeMemoryBufferHandle(int capacity) : this(true)
      {
         m_capacity = capacity;
      }


      protected SafeNativeMemoryBufferHandle(IntPtr memory, int capacity) : this(capacity)
      {
         SetHandle(memory);
      }


      
      
      /// <summary>容量を取得します。このインスタンスがサイズを指定するコンストラクタを使用して作成された場合にのみ有効であり、
      /// P/Invokeを使用するネイティブメソッドによって返されたハンドルの場合は正しくありません。
      /// </summary>
      public int Capacity
      {
         get { return m_capacity; }
      }




      /// <summary>1次元のマネージ8ビット符号なし整数配列からこのインスタンスが参照するアンマネージメモリポインターにデータをコピーします。</summary>
      /// <param name="source">コピー元の1次元配列。</param>
      /// <param name="startIndex">コピーを開始する配列のゼロベースインデックス。</param>
      /// <param name="length">コピーする配列要素の数。</param>
      public void CopyFrom(byte[] source, int startIndex, int length)
      {
         Marshal.Copy(source, startIndex, handle, length);
      }


      public void CopyFrom(char[] source, int startIndex, int length)
      {
         Marshal.Copy(source, startIndex, handle, length);
      }


      public void CopyFrom(char[] source, int startIndex, int length, int offset)
      {
         Marshal.Copy(source, startIndex, new IntPtr(handle.ToInt64() + offset), length);
      }


      /// <summary>このアンマネージメモリポインターからマネージ8ビット符号なし整数配列にデータをコピーします。</summary>
      /// <param name="sourceOffset">コピーを開始するバッファ内のオフセット。</param>
      /// <param name="destination">コピー先の配列。</param>
      public void CopyTo(int sourceOffset, byte[] destination)
      {
         if (null == destination || destination.Length == 0)
         {
            throw new ArgumentNullException("destination");
         }

         var length = destination.Length;

         if (sourceOffset + length > Capacity)
         {
            throw new ArgumentOutOfRangeException("sourceOffset", Resources.Source_OffsetAndLength_Outside_Bounds);
         }

         Marshal.Copy(new IntPtr(handle.ToInt64() + sourceOffset), destination, 0, length);
      }


      /// <summary>アンマネージメモリポインターからマネージ8ビット符号なし整数配列にデータをコピーします。</summary>
      /// <param name="destination">コピー先の配列。</param>
      /// <param name="destinationOffset">コピーを開始する宛先配列のゼロベースインデックス。</param>
      /// <param name="length">コピーする配列要素の数。</param>
      public void CopyTo(byte[] destination, int destinationOffset, int length)
      {
         if (null == destination)
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

         if (length > Capacity)
         {
            throw new ArgumentOutOfRangeException("length", Resources.Source_OffsetAndLength_Outside_Bounds);
         }

         Marshal.Copy(handle, destination, destinationOffset, length);
      }


      /// <summary>このアンマネージメモリポインターからマネージ8ビット符号なし整数配列にデータをコピーします。</summary>
      /// <param name="sourceOffset">コピーを開始するバッファ内のオフセット。</param>
      /// <param name="destination">コピー先の配列。</param>
      /// <param name="destinationOffset">コピーを開始する宛先配列のゼロベースインデックス。</param>
      /// <param name="length">コピーする配列要素の数。</param>
      public void CopyTo(int sourceOffset, byte[] destination, int destinationOffset, int length)
      {
         if (null == destination)
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

         if (length > Capacity)
         {
            throw new ArgumentOutOfRangeException("length", Resources.Source_OffsetAndLength_Outside_Bounds);
         }

         Marshal.Copy(new IntPtr(handle.ToInt64() + sourceOffset), destination, destinationOffset, length);
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


      #region Write

      public void WriteInt16(int offset, short value)
      {
         Marshal.WriteInt16(handle, offset, value);
      }

      public void WriteInt16(int offset, char value)
      {
         Marshal.WriteInt16(handle, offset, value);
      }

      public void WriteInt16(char value)
      {
         Marshal.WriteInt16(handle, value);
      }

      public void WriteInt16(short value)
      {
         Marshal.WriteInt16(handle, value);
      }

      public void WriteInt32(int offset, short value)
      {
         Marshal.WriteInt32(handle, offset, value);
      }

      public void WriteInt32(int value)
      {
         Marshal.WriteInt32(handle, value);
      }

      public void WriteInt64(int offset, long value)
      {
         Marshal.WriteInt64(handle, offset, value);
      }

      public void WriteInt64(long value)
      {
         Marshal.WriteInt64(handle, value);
      }

      public void WriteByte(int offset, byte value)
      {
         Marshal.WriteByte(handle, offset, value);
      }

      public void WriteByte(byte value)
      {
         Marshal.WriteByte(handle, value);
      }

      public void WriteIntPtr(int offset, IntPtr value)
      {
         Marshal.WriteIntPtr(handle, offset, value);
      }

      public void WriteIntPtr(IntPtr value)
      {
         Marshal.WriteIntPtr(handle, value);
      }

      #endregion // Write


      #region Read

      public byte ReadByte()
      {
         return Marshal.ReadByte(handle);
      }

      public byte ReadByte(int offset)
      {
         return Marshal.ReadByte(handle, offset);
      }

      public short ReadInt16()
      {
         return Marshal.ReadInt16(handle);
      }

      public short ReadInt16(int offset)
      {
         return Marshal.ReadInt16(handle, offset);
      }

      public int ReadInt32()
      {
         return Marshal.ReadInt32(handle);
      }

      public int ReadInt32(int offset)
      {
         return Marshal.ReadInt32(handle, offset);
      }

      public long ReadInt64()
      {
         return Marshal.ReadInt64(handle);
      }

      public long ReadInt64(int offset)
      {
         return Marshal.ReadInt64(handle, offset);
      }

      public IntPtr ReadIntPtr()
      {
         return Marshal.ReadIntPtr(handle);
      }

      public IntPtr ReadIntPtr(int offset)
      {
         return Marshal.ReadIntPtr(handle, offset);
      }

      #endregion // Read



      /// <summary>マネージオブジェクトからアンマネージメモリブロックにデータをマーシャリングします。</summary>
      public void StructureToPtr<T>(T structure, bool deleteOld) where T : notnull
      {
         Marshal.StructureToPtr<T>(structure, handle, deleteOld);
      }


      /// <summary>アンマネージメモリブロックから指定された型の新しく割り当てられたマネージオブジェクトにデータをマーシャリングします。</summary>
      /// <returns>ptrパラメータが指すデータを含むマネージオブジェクト。</returns>
      public T PtrToStructure<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicConstructors | DynamicallyAccessedMemberTypes.NonPublicConstructors)] T>(int offset)
      {
         return Marshal.PtrToStructure<T>(new IntPtr(handle.ToInt64() + offset));
      }


      /// <summary>マネージSystem.Stringを割り当て、アンマネージANSI文字列から指定された数の文字をコピーします。</summary>
      /// <returns>ptrパラメータの値がnullでない場合はアンマネージ文字列のコピーを保持するマネージ文字列。そうでなければnullを返します。</returns>
      public string PtrToStringAnsi(int offset)
      {
         return Marshal.PtrToStringAnsi(new IntPtr(handle.ToInt64() + offset));
      }

      /// <summary>マネージSystem.Stringを割り当て、アンマネージUnicode文字列から最初のnull文字までのすべての文字をコピーします。</summary>
      /// <returns>ptrパラメータの値がnullでない場合はアンマネージ文字列のコピーを保持するマネージ文字列。そうでなければnullを返します。</returns>
      public string PtrToStringUni()
      {
         return Marshal.PtrToStringUni(handle);
      }


      /// <summary>マネージSystem.Stringを割り当て、アンマネージUnicode文字列から指定された数の文字をコピーします。</summary>
      /// <returns>ptrパラメータの値がnullでない場合はアンマネージ文字列のコピーを保持するマネージ文字列。そうでなければnullを返します。</returns>
      public string PtrToStringUni(int offset, int length)
      {
         return Marshal.PtrToStringUni(new IntPtr(handle.ToInt64() + offset), length);
      }
   }
}
