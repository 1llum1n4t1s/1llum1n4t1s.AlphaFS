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

namespace Alphaleonis.Win32.Network
{
   internal static partial class NativeMethods
   {
      /// <summary>共有リソースに関する情報を含みます, 名前を含む, type, and permissions of the resource, comments associated with the resource,
      /// the maximum number of concurrent connections, the number of current connections, the local path for the resource, and a password for the current connection.
      /// </summary>
      /// <remarks>Share information, NT, level 2, requires admin rights to work.</remarks>
      /// <remarks>サポートされる最小クライアント: Windows XP [desktop apps only]</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2003 [desktop apps only]</remarks>
      [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
      internal struct SHARE_INFO_2
      {
         /// <summary>共有リソースの名前。</summary>
         [MarshalAs(UnmanagedType.LPWStr)] public readonly string shi2_netname;

         /// <summary>共有の種類.</summary>
         [MarshalAs(UnmanagedType.U4)] public readonly ShareType shi2_type;

         /// <summary>共有リソースに関するオプションのコメント。</summary>
         [MarshalAs(UnmanagedType.LPWStr)] public readonly string shi2_remark;

         /// <summary>The shared resource's permissions for servers running with share-level security.</summary>
         /// <remarks>Note that Windows does not support share-level security. This member is ignored on a server running user-level security.</remarks>
         [MarshalAs(UnmanagedType.U4)] public readonly AccessPermissions shi2_permissions;

         /// <summary>The maximum number of concurrent connections that the shared resource can accommodate.</summary>
         /// <remarks>の数 connections is unlimited if the value specified in this member is –1.</remarks>
         [MarshalAs(UnmanagedType.U4)] public readonly uint shi2_max_uses;

         /// <summary>の数 current connections to the resource.</summary>
         [MarshalAs(UnmanagedType.U4)] public readonly uint shi2_current_uses;

         /// <summary>The local path for the shared resource.</summary>
         /// <remarks>For disks, this member is the path being shared. For print queues, this member is the name of the print queue being shared.</remarks>
         [MarshalAs(UnmanagedType.LPWStr)] public readonly string shi2_path;

         /// <summary>The share's password (when the server is running with share-level security).</summary>
         [MarshalAs(UnmanagedType.LPWStr)] public readonly string shi2_passwd;
      }
   }
}
