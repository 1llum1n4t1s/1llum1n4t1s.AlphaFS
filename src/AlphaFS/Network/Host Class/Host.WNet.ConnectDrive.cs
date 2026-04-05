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
using System.Security;

namespace Alphaleonis.Win32.Network
{
   public static partial class Host
   {
      /// <summary>ネットワークリソースへの接続を作成します. この関数はローカルデバイスをネットワークリソースにリダイレクトできます.</summary>
      /// <returns>If <paramref name="localName"/> is <c>null</c> or <c>string.Empty</c>, returns the last available drive letter, <c>null</c> otherwise.</returns>
      /// <param name="localName">
      ///   The name of a local device to be redirected, such as "F:". When <paramref name="localName"/> is <c>null</c> or
      ///   <c>string.Empty</c>, the last available drive letter will be used. Letters are assigned beginning with Z:, then Y: and so on.
      /// </param>
      /// <param name="remoteName">接続先のネットワークリソース。文字列の長さは最大 <c>MAX_PATH</c> characters in length.</param>
      [SecurityCritical]
      public static string ConnectDrive(string localName, string remoteName)
      {
         return ConnectDisconnectCore(new ConnectDisconnectArguments
         {
            LocalName = localName,
            RemoteName = remoteName,
            IsDeviceMap = true
         });
      }

      
      /// <summary>ネットワークリソースへの接続を作成します. この関数はローカルデバイスをネットワークリソースにリダイレクトできます.</summary>
      /// <returns>If <paramref name="localName"/> is <c>null</c> or <c>string.Empty</c>, returns the last available drive letter, null otherwise.</returns>
      /// <param name="localName">
      ///   The name of a local device to be redirected, such as "F:". When <paramref name="localName"/> is <c>null</c> or
      ///   <c>string.Empty</c>, the last available drive letter will be used. Letters are assigned beginning with Z:, then Y: and so on.
      /// </param>
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
      /// <param name="updateProfile"><c>true</c> は成功したネットワークリソース接続を保存します。</param>
      /// <param name="saveCredentials">
      ///   オペレーティングシステムが資格情報を要求した場合、true のときに資格情報マネージャーによって資格情報が保存されます。
      /// </param>
      [SecurityCritical]
      public static string ConnectDrive(string localName, string remoteName, string userName, string password, bool prompt, bool updateProfile, bool saveCredentials)
      {
         return ConnectDisconnectCore(new ConnectDisconnectArguments
         {
            LocalName = localName,
            RemoteName = remoteName,
            UserName = userName,
            Password =  password,
            Prompt = prompt,
            UpdateProfile = updateProfile,
            SaveCredentials = saveCredentials,
            IsDeviceMap = true
         });
      }

      
      /// <summary>ネットワークリソースへの接続を作成します. この関数はローカルデバイスをネットワークリソースにリダイレクトできます.</summary>
      /// <returns>If <paramref name="localName"/> is <c>null</c> or <c>string.Empty</c>, returns the last available drive letter, null otherwise.</returns>
      /// <param name="localName">
      ///   The name of a local device to be redirected, such as "F:". When <paramref name="localName"/> is <c>null</c> or
      ///   <c>string.Empty</c>, the last available drive letter will be used. Letters are assigned beginning with Z:, then Y: and so on.
      /// </param>
      /// <param name="remoteName">接続先のネットワークリソース。文字列の長さは最大 <c>MAX_PATH</c> characters in length.</param>
      /// <param name="credentials">
      ///   An instance of <see cref="NetworkCredential"/> which provides credentials for password-based authentication schemes such as basic,
      ///   digest, NTLM, and Kerberos authentication.
      /// </param>
      /// <param name="prompt"><c>true</c> は常に認証ダイアログボックスをポップアップします。</param>
      /// <param name="updateProfile"><c>true</c> は成功したネットワークリソース接続を保存します。</param>
      /// <param name="saveCredentials">
      ///   オペレーティングシステムが資格情報を要求した場合、true のときに資格情報マネージャーによって資格情報が保存されます。
      /// </param>
      [SecurityCritical]
      public static string ConnectDrive(string localName, string remoteName, NetworkCredential credentials, bool prompt, bool updateProfile, bool saveCredentials)
      {
         return ConnectDisconnectCore(new ConnectDisconnectArguments
         {
            LocalName = localName,
            RemoteName = remoteName,
            Credential = credentials,
            Prompt = prompt,
            UpdateProfile = updateProfile,
            SaveCredentials = saveCredentials,
            IsDeviceMap = true
         });
      }

      
      /// <summary>ネットワークリソースへの接続を作成します. この関数はローカルデバイスをネットワークリソースにリダイレクトできます.</summary>
      /// <returns>If <paramref name="localName"/> is <c>null</c> or <c>string.Empty</c>, returns the last available drive letter, null otherwise.</returns>
      /// <param name="winOwner">ネットワークリソースのプロバイダーがダイアログボックスのオーナーウィンドウとして使用できるウィンドウへのハンドル。</param>
      /// <param name="localName">
      ///   The name of a local device to be redirected, such as "F:". When <paramref name="localName"/> is <c>null</c> or
      ///   <c>string.Empty</c>, the last available drive letter will be used. Letters are assigned beginning with Z:, then Y: and so on.
      /// </param>
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
      /// <param name="updateProfile"><c>true</c> は成功したネットワークリソース接続を保存します。</param>
      /// <param name="saveCredentials">
      ///   オペレーティングシステムが資格情報を要求した場合、true のときに資格情報マネージャーによって資格情報が保存されます。
      /// </param>
      [SecurityCritical]
      public static string ConnectDrive(IntPtr winOwner, string localName, string remoteName, string userName, string password, bool prompt, bool updateProfile, bool saveCredentials)
      {
         return ConnectDisconnectCore(new ConnectDisconnectArguments
         {
            WinOwner = winOwner,
            LocalName = localName,
            RemoteName = remoteName,
            UserName = userName,
            Password = password,
            Prompt = prompt,
            UpdateProfile = updateProfile,
            SaveCredentials = saveCredentials,
            IsDeviceMap = true
         });
      }

      
      /// <summary>ネットワークリソースへの接続を作成します. この関数はローカルデバイスをネットワークリソースにリダイレクトできます.</summary>
      /// <returns>If <paramref name="localName"/> is <c>null</c> or <c>string.Empty</c>, returns the last available drive letter, null otherwise.</returns>
      /// <param name="winOwner">ネットワークリソースのプロバイダーがダイアログボックスのオーナーウィンドウとして使用できるウィンドウへのハンドル。</param>
      /// <param name="localName">
      ///   The name of a local device to be redirected, such as "F:". When <paramref name="localName"/> is <c>null</c> or
      ///   <c>string.Empty</c>, the last available drive letter will be used. Letters are assigned beginning with Z:, then Y: and so on.
      /// </param>
      /// <param name="remoteName">接続先のネットワークリソース。文字列の長さは最大 <c>MAX_PATH</c> characters in length.</param>
      /// <param name="credentials">
      ///   An instance of <see cref="NetworkCredential"/> which provides credentials for password-based authentication schemes such as basic,
      ///   digest, NTLM, and Kerberos authentication.
      /// </param>
      /// <param name="prompt"><c>true</c> は常に認証ダイアログボックスをポップアップします。</param>
      /// <param name="updateProfile"><c>true</c> は成功したネットワークリソース接続を保存します。</param>
      /// <param name="saveCredentials">
      ///   オペレーティングシステムが資格情報を要求した場合、true のときに資格情報マネージャーによって資格情報が保存されます。
      /// </param>
      [SecurityCritical]
      public static string ConnectDrive(IntPtr winOwner, string localName, string remoteName, NetworkCredential credentials, bool prompt, bool updateProfile, bool saveCredentials)
      {
         return ConnectDisconnectCore(new ConnectDisconnectArguments
         {
            WinOwner = winOwner,
            LocalName = localName,
            RemoteName = remoteName,
            Credential = credentials,
            Prompt = prompt,
            UpdateProfile = updateProfile,
            SaveCredentials = saveCredentials,
            IsDeviceMap = true
         });
      }
   }
}
