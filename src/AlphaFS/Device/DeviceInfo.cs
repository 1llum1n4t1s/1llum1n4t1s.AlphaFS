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

using Alphaleonis.Win32.Network;
using System;
using System.Collections.Generic;
using System.Security;

namespace Alphaleonis.Win32.Filesystem
{
   /// <summary>ローカルまたはリモートホスト上のデバイスの情報へのアクセスを提供します。</summary>
   [Serializable]
   [SecurityCritical]
   public sealed class DeviceInfo
   {
      #region コンストラクター

      /// <summary>DeviceInfo クラスを初期化します。</summary>
      [SecurityCritical]
      public DeviceInfo()
      {
         HostName = Host.GetUncName();
      }

      /// <summary>DeviceInfo クラスを初期化します。</summary>
      /// <param name="host">リモートサーバーの DNS 名または NetBIOS 名。<c>null</c> はローカルホストを参照します。</param>
      [SecurityCritical]
      public DeviceInfo(string host)
      {
         HostName = Host.GetUncName(host).Replace(Path.UncPrefix, string.Empty);
      }

      #endregion // コンストラクター


      #region メソッド

      /// <summary>ローカルホスト上の利用可能な全デバイスを列挙します。</summary>
      /// <param name="deviceGuid"><see cref="Filesystem.DeviceGuid"/> デバイスのいずれか。</param>
      /// <returns>ローカルホストからの <see cref="Filesystem.DeviceGuid"/> 型の <see cref="IEnumerable{DeviceInfo}"/> インスタンス。</returns>
      [SecurityCritical]
      public IEnumerable<DeviceInfo> EnumerateDevices(DeviceGuid deviceGuid)
      {
         return Device.EnumerateDevicesCore(HostName, deviceGuid, true);
      }
      
      #endregion // メソッド


      #region プロパティ

      /// <summary>ベースコンテナー識別子 (ID) の <see cref="Guid"/> 値を表します。Windows プラグアンドプレイ (PnP) マネージャーがデバイスノード (devnode) にこの値を割り当てます。</summary>
      public Guid BaseContainerId { get; internal set; }


      /// <summary>デバイスインスタンスが属するデバイスセットアップクラスの名前を表します。</summary>
      public string DeviceClass { get; internal set; }


      /// <summary>デバイスインスタンスが属するデバイスセットアップクラスの <see cref="Guid"/> を表します。</summary>
      public Guid ClassGuid { get; internal set; }


      /// <summary>デバイスインスタンスの互換性のある識別子のリストを表します。</summary>
      public string CompatibleIds { get; internal set; }


      /// <summary>デバイスインスタンスの説明を表します。</summary>
      public string DeviceDescription { get; internal set; }


      /// <summary>デバイスインターフェースのパス。</summary>
      public string DevicePath { get; internal set; }


      /// <summary>デバイスインスタンスのドライバーキーのレジストリエントリ名を表します。</summary>
      public string Driver { get; internal set; }


      /// <summary>デバイスインスタンスの列挙子の名前を表します。</summary>
      public string EnumeratorName { get; internal set; }


      /// <summary>デバイスインスタンスのフレンドリ名を表します。</summary>
      public string FriendlyName { get; internal set; }


      /// <summary>デバイスインスタンスのハードウェア識別子のリストを表します。</summary>
      public string HardwareId { get; internal set; }


      /// <summary>クラスコンストラクターに渡されたホスト名。</summary>
      public string HostName { get; internal set; }


      /// <summary>デバイスのインスタンス ID を取得します。</summary>
      public string InstanceId { get; internal set; }


      /// <summary>デバイスインスタンスのバス固有の物理的な位置を表します。</summary>
      public string LocationInformation { get; internal set; }


      /// <summary>デバイスツリー内のデバイスインスタンスの位置を表します。</summary>
      public string LocationPaths { get; internal set; }


      /// <summary>デバイスインスタンスの製造元の名前を表します。</summary>
      public string Manufacturer { get; internal set; }


      /// <summary>デバイスのファームウェアから Windows に提供される物理デバイスの位置情報をカプセル化します。</summary>
      public string PhysicalDeviceObjectName { get; internal set; }


      /// <summary>デバイスインスタンスにインストールされているサービスの名前を表します。</summary>
      public string Service { get; internal set; }

      #endregion // プロパティ
   }
}
