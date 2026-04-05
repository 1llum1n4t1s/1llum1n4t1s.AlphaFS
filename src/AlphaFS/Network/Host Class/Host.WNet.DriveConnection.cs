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
using System.Net;

namespace Alphaleonis.Win32.Network
{
   /// <summary>ネットワークリソースへの一時的な接続を作成するために使用されます that will be disconnected once this instance is disposed.</summary>
   public sealed class DriveConnection : IDisposable
   {
      #region コンストラクター

      /// <summary>Creates a temporary connection to a network resource. この関数はローカルデバイスをネットワークリソースにリダイレクトできます, using the current user credentials.</summary>
      /// <param name="remoteName">接続先のネットワークリソース。文字列の長さは最大 <c>MAX_PATH</c> characters in length.</param>
      public DriveConnection(string remoteName)
      {
         Share = remoteName;

         LocalName = Host.ConnectDisconnectCore(new Host.ConnectDisconnectArguments
         {
            RemoteName = Share,
            IsDeviceMap = true
         });
      }

      
      /// <summary>Creates a temporary connection to a network resource. この関数はローカルデバイスをネットワークリソースにリダイレクトできます, using a user name and password.</summary>
      /// <param name="remoteName">接続先のネットワークリソース。文字列の長さは最大 <c>MAX_PATH</c> characters in length.</param>
      /// <param name="userName">
      ///   The user name for making the connection. If <paramref name="userName"/> is <c>null</c>, the function uses the default
      ///   user name. (The user context for the process provides the default user name)
      /// </param>
      /// <param name="password">
      ///   The password to be used for making the network connection. If <paramref name="password"/> is <c>null</c>, the function
      ///   uses the current default password associated with the user specified by <paramref name="userName"/>.
      /// </param>
      /// <param name="prompt"><c>true</c> は常に認証ダイアログボックスをポップアップします。</param>
      public DriveConnection(string remoteName, string userName, string password, bool prompt)
      {
         Share = remoteName;

         LocalName = Host.ConnectDisconnectCore(new Host.ConnectDisconnectArguments
         {
            RemoteName = Share,
            UserName = userName,
            Password = password,
            Prompt = prompt,
            IsDeviceMap = true
         });
      }

      
      /// <summary>Creates a temporary connection to a network resource. この関数はローカルデバイスをネットワークリソースにリダイレクトできます, <see cref="NetworkCredential"/> can be supplied.</summary>
      /// <param name="remoteName">接続先のネットワークリソース。文字列の長さは最大 <c>MAX_PATH</c> characters in length.</param>
      /// <param name="credentials">An instance of <see cref="NetworkCredential"/> which provides credentials for password-based authentication schemes such as basic, digest, NTLM, and Kerberos authentication.</param>
      /// <param name="prompt"><c>true</c> は常に認証ダイアログボックスをポップアップします。</param>
      public DriveConnection(string remoteName, NetworkCredential credentials, bool prompt)
      {
         Share = remoteName;

         LocalName = Host.ConnectDisconnectCore(new Host.ConnectDisconnectArguments
         {
            RemoteName = Share,
            Credential = credentials,
            Prompt = prompt,
            IsDeviceMap = true
         });
      }

      
      /// <summary><see cref="DriveConnection"/> class destructor.</summary>
      ~DriveConnection()
      {
         Dispose(false);
      }

      #endregion // コンストラクター


      #region プロパティ

      /// <summary>The last available drive letter used for this connection.</summary>
      /// <value>The last available drive letter used for this connection.</value>
      public string LocalName { get; private set; }

      
      /// <summary>The path originally specified by the user.</summary>
      /// <value>The path originally specified by the user.</value>
      public string Share { get; private set; }

      #endregion // プロパティ


      #region メソッド

      /// <summary>Releases all resources used by the <see cref="DriveConnection"/> class.</summary>
      public void Dispose()
      {
         GC.SuppressFinalize(this);
         Dispose(true);
      }


      private void Dispose(bool isDisposing)
      {
         if (isDisposing && !Utils.IsNullOrWhiteSpace(LocalName))
         {
            Host.ConnectDisconnectCore(new Host.ConnectDisconnectArguments
            {
               LocalName = LocalName,
               Prompt = true, // Use value of prompt variable for force value.
               IsDeviceMap = true,
               IsDisconnect = true
            });

            LocalName = null;
         }
      }


      /// <summary>Returns the last available drive letter used for this connection.</summary>
      /// <returns>このインスタンスを表す文字列。</returns>
      public override string ToString()
      {
         return LocalName;
      }

      #endregion // メソッド
   }
}
