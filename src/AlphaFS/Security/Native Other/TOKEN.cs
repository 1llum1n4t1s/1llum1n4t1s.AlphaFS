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

namespace Alphaleonis.Win32.Security
{
   internal static partial class NativeMethods
   {
      /// <summary>[AlphaFS] アクセストークンオブジェクトのアクセス権。</summary>
      [Flags]
      internal enum TOKEN : uint
      {
         /// <summary>プロセスにプライマリトークンをアタッチするために必要です。このタスクを達成するにはSE_ASSIGNPRIMARYTOKEN_NAME特権も必要です。</summary>
         TOKEN_ASSIGN_PRIMARY = 1,

         /// <summary>アクセストークンの複製に必要です。</summary>
         TOKEN_DUPLICATE = 2,

         /// <summary>偽装アクセストークンをプロセスにアタッチするために必要です。</summary>
         TOKEN_IMPERSONATE = 4,

         /// <summary>アクセストークンの照会に必要です。</summary>
         TOKEN_QUERY = 8,

         /// <summary>アクセストークンのソースの照会に必要です。</summary>
         TOKEN_QUERY_SOURCE = 16,

         /// <summary>アクセストークン内の特権の有効化または無効化に必要です。</summary>
         TOKEN_ADJUST_PRIVILEGES = 32,

         /// <summary>アクセストークン内のグループの属性の調整に必要です。</summary>
         TOKEN_ADJUST_GROUPS = 64,

         /// <summary>アクセストークンのデフォルトの所有者、プライマリグループ、またはDACLの変更に必要です。</summary>
         TOKEN_ADJUST_DEFAULT = 128,

         /// <summary>アクセストークンのセッションIDの調整に必要です。SE_TCB_NAME特権が必要です。</summary>
         TOKEN_ADJUST_SESSIONID = 256,

         /// <summary>STANDARD_RIGHTS_READとTOKEN_QUERYを組み合わせたものです。</summary>
         TOKEN_READ = STANDARD_RIGHTS_READ | TOKEN_QUERY,

         /// <summary>トークンのすべての可能なアクセス権を組み合わせたものです。</summary>
         TOKEN_ALL_ACCESS = STANDARD_RIGHTS_REQUIRED | TOKEN_ASSIGN_PRIMARY | TOKEN_DUPLICATE | TOKEN_IMPERSONATE | TOKEN_QUERY | TOKEN_QUERY_SOURCE |
                            TOKEN_ADJUST_PRIVILEGES | TOKEN_ADJUST_GROUPS | TOKEN_ADJUST_DEFAULT | TOKEN_ADJUST_SESSIONID
      }
   }
}
