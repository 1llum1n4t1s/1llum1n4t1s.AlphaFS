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
   /// <summary>TOKEN_PRIVILEGES構造体は、アクセストークンの特権セットに関する情報を含みます。</summary>
   [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
   internal struct TOKEN_PRIVILEGES
   {
      /// <summary>Privileges配列のエントリ数に設定する必要があります。</summary>
      [MarshalAs(UnmanagedType.U4)] public uint PrivilegeCount;

      /// <summary>LUID_AND_ATTRIBUTES構造体の配列を指定します。各構造体は特権のLUIDと属性を含みます。</summary>
      public LUID Luid;

      /// <summary>特権の属性は、以下の値の組み合わせです:
      /// SE_PRIVILEGE_ENABLED: 特権が有効です。
      /// SE_PRIVILEGE_ENABLED_BY_DEFAULT: 特権がデフォルトで有効です。
      /// SE_PRIVILEGE_REMOVED: 特権の削除に使用されます。詳細はAdjustTokenPrivilegesを参照してください。
      /// SE_PRIVILEGE_USED_FOR_ACCESS: 特権がオブジェクトまたはサービスへのアクセス取得に使用されました。このフラグは、不要な特権を含む可能性があるクライアントアプリケーションから渡されたセット内の関連する特権を識別するために使用されます。
      /// </summary>
      [MarshalAs(UnmanagedType.U4)] public uint Attributes;
   }
}
