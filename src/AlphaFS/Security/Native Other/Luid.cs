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

namespace Alphaleonis.Win32.Security
{
   /// <summary>LUIDは、生成されたシステム上でのみ一意であることが保証される64ビット値です。ローカル一意識別子（LUID）の一意性は、システムが再起動されるまでのみ保証されます。</summary>
   /// <remarks>
   /// <para>サポートされる最小クライアント: Windows XP [デスクトップアプリのみ]</para>
   /// <para>サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリのみ]</para>
   /// </remarks>
   [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
   internal struct LUID
   {
      /// <summary>下位ビット。</summary>
      [MarshalAs(UnmanagedType.U4)] public uint LowPart;

      /// <summary>上位ビット。</summary>
      [MarshalAs(UnmanagedType.U4)] public uint HighPart;
   }
}
