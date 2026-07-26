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
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Net.NetworkInformation;

namespace Alphaleonis.Win32.Network
{
   /// <summary>ネットワークへの接続を表します。</summary>
   public sealed class NetworkConnectionInfo : IDisposable
   {
      #region プライベートフィールド

      private NativeMethods.NetworkConnectionWrapper _networkConnection;

      #endregion // プライベートフィールド


      #region コンストラクター

      internal NetworkConnectionInfo(NativeMethods.NetworkConnectionWrapper networkConnection)
      {
         _networkConnection = networkConnection;
      }

      #endregion // コンストラクター


      #region IDisposable

      /// <summary>Dispose 漏れ時のセーフネットとして基になる COM 参照を解放するファイナライザ。</summary>
      ~NetworkConnectionInfo()
      {
         Dispose(false);
      }

      /// <summary>基になる COM 参照を解放します。</summary>
      public void Dispose()
      {
         Dispose(true);
         GC.SuppressFinalize(this);
      }

      private void Dispose(bool disposing)
      {
         _networkConnection?.Dispose();
         _networkConnection = null;
      }

      #endregion // IDisposable


      #region プライベートヘルパー

      private void ThrowIfDisposed()
      {
         if (null == _networkConnection)
            throw new ObjectDisposedException(GetType().FullName);
      }

      #endregion // プライベートヘルパー


      #region プロパティ

      /// <summary>この接続の一意の識別子を取得します。このプロパティの値はキャッシュされません。</summary>
      [SuppressMessage("Microsoft.Naming", "CA1709:IdentifiersShouldBeCasedCorrectly", MessageId = "ID")]
      public Guid ConnectionId
      {
         get { ThrowIfDisposed(); return _networkConnection.GetConnectionId(); }
      }


      /// <summary>この接続の接続性を示す値を取得します。このプロパティの値はキャッシュされません。</summary>
      public ConnectivityStates Connectivity
      {
         get { ThrowIfDisposed(); return _networkConnection.GetConnectivity(); }
      }


      /// <summary>この接続に関連付けられたネットワークが Active Directory ネットワークかどうか、およびマシンが Active Directory によって認証されているかどうかを示す値を取得します。このプロパティの値はキャッシュされません。</summary>
      public DomainType DomainType
      {
         get { ThrowIfDisposed(); return _networkConnection.GetDomainType(); }
      }


      /// <summary>この接続がネットワーク接続を持っているかどうかを示す値を取得します。このプロパティの値はキャッシュされません。</summary>
      public bool IsConnected
      {
         get { ThrowIfDisposed(); return _networkConnection.IsConnected; }
      }


      /// <summary>この接続がインターネットアクセスを持っているかどうかを示す値を取得します。このプロパティの値はキャッシュされません。</summary>
      public bool IsConnectedToInternet
      {
         get { ThrowIfDisposed(); return _networkConnection.IsConnectedToInternet; }
      }


      /// <summary>この接続に関連付けられたネットワークを表す新しい <see cref="NetworkInfo"/> インスタンスを作成します。
      /// <para>呼び出し元は返された <see cref="NetworkInfo"/> インスタンスの破棄に責任を持ちます。</para></summary>
      /// <returns>新しい <see cref="NetworkInfo"/> インスタンス。呼び出し元は使用後にこのオブジェクトを破棄する必要があります。</returns>
      public NetworkInfo GetNetworkInfo()
      {
         ThrowIfDisposed();
         return new NetworkInfo(_networkConnection.GetNetwork());
      }


      /// <summary>この接続のネットワークインターフェースを取得します。このプロパティの値はキャッシュされません。</summary>
      public NetworkInterface NetworkInterface
      {
         get
         {
            ThrowIfDisposed();

            var adapterId = _networkConnection.GetAdapterId();
            
            
            foreach (var nic in NetworkInterface.GetAllNetworkInterfaces())
            {
               if (!Guid.TryParse(nic.Id, out var guid))
               {
                  continue;
               }

               if (Equals(adapterId, guid))
               {
                  return nic;
               }
            }


            return null;
         }
      }

      #endregion // プロパティ


      #region メソッド

      /// <summary>この接続のネットワーク名とアダプター名を返します。
      /// <para>注意: このメソッドは COM 呼び出しを行い、すべてのネットワークインターフェースを列挙するため、コストが高い場合があります。タイトなループでの呼び出しは避けてください。</para></summary>
      /// <returns>このインスタンスを表す文字列。</returns>
      public override string ToString()
      {
         using var netInfo = GetNetworkInfo();
         var nic = NetworkInterface;

         return null != netInfo && null != nic ? string.Format(CultureInfo.CurrentCulture, "{0} {1}", netInfo.Name, nic.Name) : GetType().Name;
      }


      /// <summary>指定されたオブジェクトが現在のオブジェクトと等しいかどうかを判断します。</summary>
      /// <param name="obj">比較する別のオブジェクト。</param>
      /// <returns>指定されたオブジェクトが現在のオブジェクトと等しい場合は <c>true</c>、それ以外の場合は <c>false</c>。</returns>
      public override bool Equals(object obj)
      {
         if (null == obj || GetType() != obj.GetType())
         {
            return false;
         }

         var other = obj as NetworkConnectionInfo;

         return null != other && Equals(ConnectionId, other.ConnectionId);
      }


      /// <summary>特定の型のハッシュ関数として機能します。</summary>
      /// <returns>現在のオブジェクトのハッシュコード。</returns>
      public override int GetHashCode()
      {
         return ConnectionId.GetHashCode();
      }


      /// <summary>== 演算子を実装します。</summary>
      /// <param name="left">A。</param>
      /// <param name="right">B。</param>
      /// <returns>演算子の結果。</returns>
      public static bool operator ==(NetworkConnectionInfo left, NetworkConnectionInfo right)
      {
         return ReferenceEquals(left, null) && ReferenceEquals(right, null) || !ReferenceEquals(left, null) && !ReferenceEquals(right, null) && left.Equals(right);
      }


      /// <summary>!= 演算子を実装します。</summary>
      /// <param name="left">A。</param>
      /// <param name="right">B。</param>
      /// <returns>演算子の結果。</returns>
      public static bool operator !=(NetworkConnectionInfo left, NetworkConnectionInfo right)
      {
         return !(left == right);
      }

      #endregion // メソッド
   }
}
