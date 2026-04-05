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
using System.IO;
using System.Runtime.InteropServices;

namespace Alphaleonis.Win32.Filesystem
{
   internal static partial class NativeMethods
   {
      /// <summary>指定されたディレクトリ内のファイルに関する情報を格納します。ディレクトリハンドルに使用されます。GetFileInformationByHandleEx の呼び出し時にのみ使用してください。</summary>
      /// <remarks>
      /// GetFileInformationByHandleEx の各呼び出しで返されるファイル数は、関数に渡されるバッファーのサイズに依存します。
      /// 同じハンドルでの後続の GetFileInformationByHandleEx 呼び出しは、最後に返されたファイルの後から列挙操作を再開します。
      /// </remarks>
      [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
      internal struct FILE_ID_BOTH_DIR_INFO
      {
         /// <summary>次に返される FILE_ID_BOTH_DIR_INFO 構造体へのオフセット。他のエントリが続かない場合はゼロ (0) を格納します。</summary>
         [MarshalAs(UnmanagedType.U4)]
         public readonly int NextEntryOffset;

         /// <summary>親ディレクトリ内のファイルのバイトオフセット。NTFS などのファイルシステムでは、
         /// 親ディレクトリ内のファイル位置が固定されておらず、ソート順を維持するためにいつでも変更される可能性があるため、このメンバーは未定義です。
         /// </summary>
         [MarshalAs(UnmanagedType.U4)]
         public readonly uint FileIndex;

         /// <summary>ファイルが作成された日時。</summary>
         public FILETIME CreationTime;

         /// <summary>ファイルが最後にアクセスされた日時。</summary>
         public FILETIME LastAccessTime;

         /// <summary>ファイルが最後に書き込まれた日時。</summary>
         public FILETIME LastWriteTime;

         /// <summary>ファイルが最後に変更された日時。</summary>
         public FILETIME ChangeTime;

         /// <summary>ファイルの先頭からファイルの末尾までのバイトオフセットとしての、絶対的な新しいファイル終端位置。
         /// この値はゼロベースであるため、実際にはファイル内の最初の空きバイトを指します。
         /// つまり、EndOfFile はファイル内の最後の有効なバイトの直後のバイトへのオフセットです。
         /// </summary>
         public readonly long EndOfFile;

         /// <summary>ファイルに割り当てられたバイト数。この値は通常、基になる物理デバイスのセクターまたはクラスターサイズの倍数です。</summary>
         public readonly long AllocationSize;

         /// <summary>ファイル属性。</summary>
         public readonly FileAttributes FileAttributes;

         /// <summary>ファイル名の長さ。</summary>
         [MarshalAs(UnmanagedType.U4)]
         public readonly uint FileNameLength;

         /// <summary>ファイルの拡張属性のサイズ。</summary>
         [MarshalAs(UnmanagedType.U4)]
         public readonly int EaSize;

         /// <summary>ShortName の長さ。</summary>
         [MarshalAs(UnmanagedType.U1)]
         public readonly byte ShortNameLength;

         /// <summary>8.3 ファイル命名規則（例: "FILENAME.TXT"）によるファイルの短い名前。</summary>
         [MarshalAs(UnmanagedType.ByValArray, SizeConst = 12, ArraySubType = UnmanagedType.U2)]
         public readonly char[] ShortName;

         /// <summary>ファイルID。</summary>
         public readonly long FileId;

         /// <summary>ファイル名文字列の最初の文字。メモリ上ではこの後に文字列の残りの部分が続きます。</summary>
         public IntPtr FileName;
      }
   }
}