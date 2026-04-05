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
      /// <summary>接続の識別番号を含みます, number of open files, connection time, number of users on the connection, and the type of connection.</summary>
      /// <remarks>サポートされる最小クライアント: Windows XP [desktop apps only]</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2003 [desktop apps only]</remarks>
      [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
      internal struct CONNECTION_INFO_1
      {
         /// <summary>接続識別番号を指定します。</summary>
         [MarshalAs(UnmanagedType.U4)] public readonly uint coni1_id;

         /// <summary>A combination of values that specify the type of connection made from the local device name to the shared resource.</summary>
         [MarshalAs(UnmanagedType.U4)] public readonly ShareType coni1_type;

         /// <summary>接続の結果として現在開いているファイルの数を指定します。</summary>
         [MarshalAs(UnmanagedType.U4)] public readonly uint coni1_num_opens;

         /// <summary>接続上のユーザー数を指定します。</summary>
         [MarshalAs(UnmanagedType.U4)] public readonly uint coni1_num_users;

         /// <summary>接続が確立されてからの秒数を指定します。</summary>
         [MarshalAs(UnmanagedType.U4)] public readonly uint coni1_time;

         /// <summary>If the server sharing the resource is running with user-level security, the UserName member describes which user made the connection. If the server is running with share-level security, coni1_username describes which computer (computername) made the connection.</summary>
         /// <remarks>Note that Windows does not support share-level security.</remarks>
         [MarshalAs(UnmanagedType.LPWStr)] public readonly string coni1_username;

         /// <summary>String that specifies either the share name of the server's shared resource or the computername of the client. The value of this member depends on which name was specified as the qualifier parameter to the NetConnectionEnum function.</summary>
         [MarshalAs(UnmanagedType.LPWStr)] public readonly string oni1_netname;
      }
   }
}
