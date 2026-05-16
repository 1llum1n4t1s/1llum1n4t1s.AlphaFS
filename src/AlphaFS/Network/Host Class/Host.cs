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

using Alphaleonis.Win32.Filesystem;
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.Net;
using System.Net.NetworkInformation;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;
using Path = Alphaleonis.Win32.Filesystem.Path;

namespace Alphaleonis.Win32.Network
{
   /// <summary>ローカルまたはリモートホストからネットワークリソース情報を取得するための静的メソッドを提供します。</summary>
   public static partial class Host
   {
      private static readonly NativeMethods.NetworkListManagerWrapper Manager = NativeMethods.CreateNetworkListManager();

      internal delegate uint EnumerateNetworkObjectDelegate(FunctionData functionData, out SafeGlobalMemoryBufferHandle netApiBuffer, [MarshalAs(UnmanagedType.I4)] int prefMaxLen,
         [MarshalAs(UnmanagedType.U4)] out uint entriesRead, [MarshalAs(UnmanagedType.U4)] out uint totalEntries, [MarshalAs(UnmanagedType.U4)] out uint resumeHandle);


      /// <summary>Win32 関数に追加データを渡すために使用される構造体。</summary>
      internal struct FunctionData
      {
         public int EnumType;
         public string ExtraData1;
         public string ExtraData2;
      }


      internal struct ConnectDisconnectArguments
      {
         /// <summary>ネットワークリソースのプロバイダーがダイアログボックスのオーナーウィンドウとして使用できるウィンドウへのハンドル。</summary>
         public IntPtr WinOwner;

         /// <summary>リダイレクトするローカルデバイスの名前。例: "F:"。<see cref="LocalName"/> が <c>null</c> または <c>string.Empty</c> の場合、最後に利用可能なドライブ文字が使用されます。文字は Z: から始まり、次に Y: というように割り当てられます。</summary>
         public string LocalName;

         /// <summary>接続/切断するネットワークリソース。例: \\server または \\server\share。文字列の長さは最大 <see cref="Filesystem.NativeMethods.MaxPath"/> 文字です。</summary>
         public string RemoteName;

         /// <summary><see cref="NetworkCredential"/> インスタンス。これか、<see cref="UserName"/> と <see cref="Password"/> の組み合わせのいずれかを使用します。</summary>
         public NetworkCredential Credential;

         /// <summary>接続を確立するためのユーザー名。<see cref="UserName"/> が <c>null</c> の場合、関数はデフォルトのユーザー名を使用します。（プロセスのユーザーコンテキストがデフォルトのユーザー名を提供します）</summary>
         public string UserName;

         /// <summary>ネットワーク接続の確立に使用するパスワード。<see cref="Password"/> が <c>null</c> の場合、関数は <see cref="UserName"/> で指定されたユーザーに関連付けられた現在のデフォルトパスワードを使用します。</summary>
         public string Password;

         /// <summary><c>true</c> は常に認証ダイアログボックスをポップアップします。</summary>
         public bool Prompt;

         /// <summary><c>true</c> は成功したネットワークリソース接続を保存します。</summary>
         public bool UpdateProfile;

         /// <summary>オペレーティングシステムが資格情報を要求した場合、true のときに資格情報マネージャーによって資格情報が保存されます。</summary>
         public bool SaveCredentials;

         /// <summary><c>true</c> は操作がドライブマッピングに関するものであることを示します。</summary>
         public bool IsDeviceMap;

         /// <summary><c>true</c> はネットワークリソースからの切断が必要であることを示します。それ以外の場合は接続します。</summary>
         public bool IsDisconnect;
      }


      /// <summary>ネットワークリソースへの接続/切断を行います。この関数はローカルデバイスをネットワークリソースにリダイレクトできます。</summary>
      /// <returns><see cref="ConnectDisconnectArguments.LocalName"/> が <c>null</c> または <c>string.Empty</c> の場合、最後に利用可能なドライブ文字を返します。それ以外の場合は null。</returns>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="NetworkInformationException"/>
      /// <param name="arguments"><see cref="ConnectDisconnectArguments"/>。</param>
      [SuppressMessage("Microsoft.Usage", "CA2208:InstantiateArgumentExceptionsCorrectly")]
      [SecurityCritical]
      internal static string ConnectDisconnectCore(ConnectDisconnectArguments arguments)
      {
         uint lastError;

         // 常にバックスラッシュを削除します。
         if (!Utils.IsNullOrWhiteSpace(arguments.LocalName))
         {
            arguments.LocalName = Path.RemoveTrailingDirectorySeparator(arguments.LocalName).ToUpperInvariant();
         }


         if (!Utils.IsNullOrWhiteSpace(arguments.RemoteName))
         {
            if (!arguments.RemoteName.StartsWith(Path.UncPrefix, StringComparison.Ordinal))
            {
               arguments.RemoteName = Path.UncPrefix + arguments.RemoteName;
            }


            // 常にバックスラッシュを削除します。
            if (!Utils.IsNullOrWhiteSpace(arguments.RemoteName))
            {
               arguments.RemoteName = Path.RemoveTrailingDirectorySeparator(arguments.RemoteName);
            }
         }

         
         // 切断

         if (arguments.IsDisconnect)
         {
            var force = arguments.Prompt; // Use value of prompt variable for force value.
            var target = arguments.IsDeviceMap ? arguments.LocalName : arguments.RemoteName;

            if (Utils.IsNullOrWhiteSpace(target))
            {
               throw new ArgumentNullException(arguments.IsDeviceMap ? "localName" : "remoteName");
            }


            lastError = NativeMethods.WNetCancelConnection(target, arguments.UpdateProfile ? NativeMethods.Connect.UpdateProfile : NativeMethods.Connect.None, force);

            if (lastError != Win32Errors.NO_ERROR)
            {
               throw new NetworkInformationException((int) lastError);
            }

            return null;
         }

         
         // 接続

         // arguments.LocalName は null または空が許可されています。

         if (Utils.IsNullOrWhiteSpace(arguments.RemoteName) && !arguments.IsDeviceMap)
         {
            throw new ArgumentNullException("arguments.RemoteName");
         }


         // 提供された場合、NetworkCredential インスタンスのデータを使用します。
         if (arguments.Credential != null)
         {
            arguments.UserName = Utils.IsNullOrWhiteSpace(arguments.Credential.Domain)
               ? arguments.Credential.UserName
               : string.Format(CultureInfo.InvariantCulture, @"{0}\{1}", arguments.Credential.Domain, arguments.Credential.UserName);

            arguments.Password = arguments.Credential.Password;
         }


         // 接続引数を組み立てます。
         var connect = NativeMethods.Connect.None;

         if (arguments.IsDeviceMap)
         {
            connect = connect | NativeMethods.Connect.Redirect;
         }

         if (arguments.Prompt)
         {
            connect = connect | NativeMethods.Connect.Prompt | NativeMethods.Connect.Interactive;
         }

         if (arguments.UpdateProfile)
         {
            connect = connect | NativeMethods.Connect.UpdateProfile;
         }

         if (arguments.SaveCredentials)
         {
            connect = connect | NativeMethods.Connect.SaveCredentialManager;
         }


         // 構造体を初期化します。
         var resource = new NativeMethods.NETRESOURCE
         {
            lpLocalName = arguments.LocalName,
            lpRemoteName = arguments.RemoteName,
            dwType = NativeMethods.ResourceType.Disk
         };

         // 3 文字分: "X:\0" (ドライブ X: とヌル終端文字)
         uint bufferSize = 3;
         StringBuilder buffer;

         do
         {
            buffer = new StringBuilder((int) bufferSize);

            uint result;
            lastError = NativeMethods.WNetUseConnection(arguments.WinOwner, ref resource, arguments.Password, arguments.UserName, connect, buffer, out bufferSize, out result);

            switch (lastError)
            {
               case Win32Errors.NO_ERROR:
                  break;

               case Win32Errors.ERROR_MORE_DATA:
                  // MSDN, lpBufferSize: バッファが十分でないために呼び出しが失敗した場合、
                  // 関数はこの場所に必要なバッファサイズを返します。
                  //
                  // Windows 8 x64: bufferSize は変更されません。

                  bufferSize = bufferSize * 2;
                  break;
            }

         } while (lastError == Win32Errors.ERROR_MORE_DATA);


         if (lastError != Win32Errors.NO_ERROR)
         {
            throw new NetworkInformationException((int) lastError);
         }


         return arguments.IsDeviceMap ? buffer.ToString() : null;
      }
      

      [SecurityCritical]
      internal static IEnumerable<TStruct> EnumerateNetworkObjectCore<TStruct, [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicConstructors | DynamicallyAccessedMemberTypes.NonPublicConstructors)] TNative>(FunctionData functionData, Func<TNative, SafeGlobalMemoryBufferHandle, TStruct> createTStruct, EnumerateNetworkObjectDelegate enumerateNetworkObject, bool continueOnException)
      {
         int objectSize;
         bool isString;

         switch (functionData.EnumType)
         {
            // 論理ドライブ
            case 1:
               isString = true;
               objectSize = 6; // 常に 6 であるべきです。
               break;

            default:
               isString = typeof(TNative) == typeof(string);
               objectSize = isString ? 0 : Marshal.SizeOf<TNative>();
               break;
         }


         uint lastError;
         do
         {
            uint totalEntries;
            uint resumeHandle;

            lastError = enumerateNetworkObject(functionData, out var buffer, NativeMethods.MaxPreferredLength, out var entriesRead, out totalEntries, out resumeHandle);

            using (buffer)
               switch (lastError)
               {
                  case Win32Errors.NERR_Success:
                  case Win32Errors.ERROR_MORE_DATA:
                     if (entriesRead > 0)
                     {
                        for (int i = 0, itemOffset = 0; i < entriesRead; i++, itemOffset += objectSize)

                           yield return (TStruct) (isString ? buffer.PtrToStringUni(itemOffset, 2) : (object) createTStruct(buffer.PtrToStructure<TNative>(itemOffset), buffer));
                     }
                     break;

                  
                  // SHARE_INFO_503 が要求されたがサポートされていない場合に観察されます。
                  case Win32Errors.RPC_X_BAD_STUB_DATA:
                  case Win32Errors.ERROR_NOT_SUPPORTED:
                     yield break;


                  default:
                     if (lastError != Win32Errors.NO_ERROR && !continueOnException)
                     {
                        throw new NetworkInformationException((int) lastError);
                     }
                     break;
               }

         } while (lastError == Win32Errors.ERROR_MORE_DATA);
      }


      /// <summary>このメソッドは <see cref="NativeMethods.REMOTE_NAME_INFO"/> レベルを使用して完全な REMOTE_NAME_INFO 構造体を取得します。</summary>
      /// <returns><see cref="NativeMethods.REMOTE_NAME_INFO"/> 構造体。</returns>
      /// <remarks>AlphaFS は SUBST.EXE を使用して作成されたネットワークドライブを無効として扱います。</remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="PathTooLongException"/>
      /// <exception cref="NetworkInformationException"/>
      /// <param name="path">ドライブ名を含むローカルパス。</param>
      /// <param name="continueOnException"><c>true</c> はリソース不足などの失敗から発生する可能性のある例外を抑制します。</param>
      [SecurityCritical]
      internal static NativeMethods.REMOTE_NAME_INFO GetRemoteNameInfoCore(string path, bool continueOnException)
      {
         if (Utils.IsNullOrWhiteSpace(path))
         {
            throw new ArgumentNullException("path");
         }


         path = Path.GetRegularPathCore(path, GetFullPathOptions.CheckInvalidPathChars, false);


         uint lastError;
         uint bufferSize = 1024;

         do
         {
            using var buffer = new SafeGlobalMemoryBufferHandle((int) bufferSize);
            // 構造体: UNIVERSAL_NAME_INFO_LEVEL = 1 (AlphaFS では使用されません)。
            // 構造体: REMOTE_NAME_INFO_LEVEL    = 2

            lastError = NativeMethods.WNetGetUniversalName(path, 2, buffer, out bufferSize);

            if (lastError == Win32Errors.NO_ERROR)
            {
               return buffer.PtrToStructure<NativeMethods.REMOTE_NAME_INFO>(0);
            }

         } while (lastError == Win32Errors.ERROR_MORE_DATA);


         if (lastError != Win32Errors.NO_ERROR && !continueOnException)
         {
            throw new NetworkInformationException((int) lastError);
         }


         // 空の構造体を返します（すべてのフィールドが null に設定されています）。
         return new NativeMethods.REMOTE_NAME_INFO();
      }
   }
}
