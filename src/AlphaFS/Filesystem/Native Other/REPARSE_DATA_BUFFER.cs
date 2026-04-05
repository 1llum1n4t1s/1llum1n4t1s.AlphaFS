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

using System.Runtime.InteropServices;

namespace Alphaleonis.Win32.Filesystem
{
   internal static partial class NativeMethods
   {
      [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
      internal struct REPARSE_DATA_BUFFER
      {
         /// <summary>リパースポイントタグ。Microsoft リパースポイントタグでなければなりません。</summary>
         public ReparsePointTag ReparseTag;

         /// <summary>Reserved メンバーの後のデータのサイズ（バイト単位）。
         /// 次のように計算できます: (4 * sizeof(ushort)) + SubstituteNameLength + PrintNameLength + (namesAreNullTerminated ? 2 * sizeof(char) : 0);
         /// </summary>
         public ushort ReparseDataLength;

         /// <summary>予約済み。使用しないでください。</summary>
         public ushort Reserved;

         /// <summary>PathBuffer 配列内の代替名文字列のオフセット（バイト単位）。</summary>
         public ushort SubstituteNameOffset;

         /// <summary>代替名文字列の長さ（バイト単位）。この文字列が null 終端の場合、SubstituteNameLength には null 文字のスペースは含まれません。</summary>
         public ushort SubstituteNameLength;

         /// <summary>PathBuffer 配列内の表示名文字列のオフセット（バイト単位）。</summary>
         public ushort PrintNameOffset;

         /// <summary>表示名文字列の長さ（バイト単位）。この文字列が null 終端の場合、PrintNameLength には null 文字のスペースは含まれません。</summary>
         public ushort PrintNameLength;

         /// <summary>Unicode エンコードされたパス文字列を含むバッファー。パス文字列には代替名文字列と表示名文字列が含まれます。</summary>
         [MarshalAs(UnmanagedType.ByValArray, SizeConst = 16368)] public byte[] PathBuffer;
      }
   }
}
