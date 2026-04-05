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

namespace Alphaleonis.Win32.Network
{
   internal static partial class NativeMethods
   {
      /// <summary>WNetUseConnection() 関数で使用されます; 接続を記述するビットフラグのセット. This parameter can be any combination of the following values.</summary>
      [Flags]
      internal enum Connect
      {
         /// <summary>No Connect options are used.</summary>
         None = 0,

         /// <summary>This flag instructs the operating system to store the network resource connection. If this bit flag is set, the operating system automatically attempts to restore the connection when the user logs on. The system remembers only successful connections that redirect local devices. It does not remember connections that are unsuccessful or deviceless connections.</summary>
         UpdateProfile = 1,

         /// <summary>このフラグが設定されている場合、オペレーティングシステムは認証のためにユーザーと対話する場合があります.</summary>
         Interactive = 8,

         /// <summary>このフラグは、ユーザーに代替を提供する機会なしにユーザー名やパスワードのデフォルト設定を使用しないようにシステムに指示します. This flag is ignored unless <see cref="Interactive"/> is also set.</summary>
         Prompt = 16,

         /// <summary>このフラグは接続の確立時にローカルデバイスのリダイレクトを強制します.</summary>
         Redirect = 128,

         ///// <summary>If this flag is set, the connection was made using a local device redirection. If the lpAccessName parameter points to a buffer, the local device name is copied to the buffer.</summary>
         //LocalDrive = 256,

         // <summary>If this flag is set, the operating system prompts the user for authentication using the command line instead of a graphical user interface (GUI). This flag is ignored unless <see cref="Interactive"/> is also set.</summary>
         //CommandLine = 2048,

         /// <summary>このフラグが設定されている場合、オペレーティングシステムが資格情報を要求すると, the credential should be saved by the credential manager. If the credential manager is disabled for the caller's logon session, or if the network provider does not support saving credentials, this flag is ignored. This flag is also ignored unless you set the "CommandLine" flag.</summary>
         SaveCredentialManager = 4096
      }
   }
}