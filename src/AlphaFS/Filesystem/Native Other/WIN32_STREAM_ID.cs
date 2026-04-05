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

namespace Alphaleonis.Win32.Filesystem
{
   internal static partial class NativeMethods
   {
      /// <summary>ストリームデータを格納します。</summary>
      [StructLayout(LayoutKind.Sequential, Pack = 1, CharSet = CharSet.Unicode)]
      [Serializable]
      internal struct WIN32_STREAM_ID
      {
         /// <summary>ストリームデータの種類。</summary>
         [MarshalAs(UnmanagedType.U4)]
         public readonly STREAM_ID dwStreamId;

         /// <summary>異なるオペレーティングシステム間の転送を容易にするためのデータ属性。</summary>
         [MarshalAs(UnmanagedType.U4)]
         public readonly STREAM_ATTRIBUTE dwStreamAttribute;

         /// <summary>データのサイズ（バイト単位）。</summary>
         [MarshalAs(UnmanagedType.U8)]
         public readonly ulong Size;

         /// <summary>代替データストリームの名前の長さ（バイト単位）。</summary>
         [MarshalAs(UnmanagedType.U4)]
         public readonly uint dwStreamNameSize;
      }
   }
}
