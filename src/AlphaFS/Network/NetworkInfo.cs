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
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;

namespace Alphaleonis.Win32.Network
{
   /// <summary>ローカルマシン上のネットワークを表します。同様のネットワークシグネチャを持つネットワーク接続のコレクションを表すこともできます。</summary>
   [Serializable]
   public class NetworkInfo : IEquatable<NetworkInfo>, IDisposable
   {
      #region プライベートフィールド

      [NonSerialized]
      private NativeMethods.NetworkWrapper _network;

      #endregion // プライベートフィールド


      #region コンストラクター

      internal NetworkInfo(NativeMethods.NetworkWrapper network)
      {
         _network = network;
      }

      #endregion // コンストラクター


      #region IDisposable

      /// <summary>基になる COM 参照を解放します。</summary>
      public void Dispose()
      {
         _network?.Dispose();
         _network = null;
      }

      #endregion // IDisposable


      #region プライベートヘルパー

      private void ThrowIfDisposed()
      {
         if (null == _network)
            throw new ObjectDisposedException(GetType().FullName);
      }

      #endregion // プライベートヘルパー


      #region プロパティ

      /// <summary>ネットワークのカテゴリを取得します。カテゴリは信頼済み、信頼されていない、または認証済みです。このプロパティの値はキャッシュされません。</summary>
      public NetworkCategory Category
      {
         get { ThrowIfDisposed(); return _network.GetCategory(); }
      }


      /// <summary>ネットワークのネットワーク接続を取得します。このプロパティの値はキャッシュされません。返されたコレクション内の各アイテムは呼び出し元が破棄する必要があります。</summary>
      public IEnumerable<NetworkConnectionInfo> Connections
      {
         get
         {
            ThrowIfDisposed();

            // 呼び出し元が部分的に列挙した場合の COM 参照リークを防ぐため、
            // すべての COM ラッパーを NetworkConnectionInfo オブジェクト（ファイナライザーを持つ）に即座にラップします。
            var connections = _network.GetNetworkConnections();
            var result = new List<NetworkConnectionInfo>();

            try
            {
               foreach (var connection in connections)
                  result.Add(new NetworkConnectionInfo(connection));
            }
            catch
            {
               foreach (var item in result)
                  item.Dispose();

               throw;
            }

            return result;
         }
      }


      /// <summary>ネットワークが接続されたローカルの日時を取得します。このプロパティの値はキャッシュされません。</summary>
      public DateTime ConnectionTime
      {
         get { return ConnectionTimeUtc.ToLocalTime(); }
      }


      /// <summary>ネットワークが接続された UTC 日時を取得します。このプロパティの値はキャッシュされません。</summary>
      public DateTime ConnectionTimeUtc
      {
         get
         {
            ThrowIfDisposed();

            uint unused1, unused2;

            _network.GetTimeCreatedAndConnected(out unused1, out unused2, out var low, out var high);
            
            long time = high;

            // 日付情報を上位ビットにシフトします。
            time <<= 32;
            time |= low;

            return DateTime.FromFileTimeUtc(time);
         }
      }


      /// <summary>ネットワークの接続状態を取得します。このプロパティの値はキャッシュされません。</summary>
      /// <remarks>Connectivity はネットワークが接続されているかどうか、およびネットワークトラフィックに使用されているプロトコルに関する情報を提供します。</remarks>
      public ConnectivityStates Connectivity
      {
         get { ThrowIfDisposed(); return _network.GetConnectivity(); }
      }


      /// <summary>ネットワークが作成されたローカルの日時を取得します。このプロパティの値はキャッシュされません。</summary>
      public DateTime CreationTime
      {
         get { return CreationTimeUtc.ToLocalTime(); }
      }


      /// <summary>ネットワークが作成された UTC 日時を取得します。このプロパティの値はキャッシュされません。</summary>
      public DateTime CreationTimeUtc
      {
         get
         {
            ThrowIfDisposed();

            uint unused1, unused2;

            _network.GetTimeCreatedAndConnected(out var low, out var high, out unused1, out unused2);

            long time = high;

            // 値を上位ビットにシフトします。
            time <<= 32;
            time |= low;

            return DateTime.FromFileTimeUtc(time);
         }
      }


      /// <summary>ネットワークの説明を取得します。このプロパティの値はキャッシュされません。</summary>
      public string Description
      {
         get { ThrowIfDisposed(); return _network.GetDescription(); }

         // AlphaFS でこれを許可すべきか？
         //private set { _network.SetDescription(value); }
      }


      /// <summary>ネットワークのドメインタイプを取得します。このプロパティの値はキャッシュされません。</summary>
      /// <remarks>ドメインは、ネットワークが Active Directory ネットワークであるかどうか、およびマシンが Active Directory によって認証されているかどうかを示します。</remarks>
      public DomainType DomainType
      {
         get { ThrowIfDisposed(); return _network.GetDomainType(); }
      }


      /// <summary>ネットワーク接続があるかどうかを示す値を取得します。このプロパティの値はキャッシュされません。</summary>
      public bool IsConnected
      {
         get { ThrowIfDisposed(); return _network.IsConnected; }
      }


      /// <summary>インターネット接続があるかどうかを示す値を取得します。このプロパティの値はキャッシュされません。</summary>
      public bool IsConnectedToInternet
      {
         get { ThrowIfDisposed(); return _network.IsConnectedToInternet; }
      }


      /// <summary>ネットワークの名前を取得します。このプロパティの値はキャッシュされません。</summary>
      public string Name
      {
         get { ThrowIfDisposed(); return _network.GetName(); }

         // AlphaFS でこれを許可すべきか？
         //private set { _network.SetName(value); }
      }


      /// <summary>ネットワークの一意の識別子を取得します。このプロパティの値はキャッシュされません。</summary>
      [SuppressMessage("Microsoft.Naming", "CA1709:IdentifiersShouldBeCasedCorrectly", MessageId = "ID")]
      public Guid NetworkId
      {
         get { ThrowIfDisposed(); return _network.GetNetworkId(); }
      }

      #endregion // プロパティ


      #region メソッド
      
      /// <summary>ネットワーク名、説明、カテゴリを返します。
      /// <para>注意: このメソッドはネットワークプロパティを取得するために COM 呼び出しを行います。</para></summary>
      /// <returns>このインスタンスを表す文字列。</returns>
      public override string ToString()
      {
         ThrowIfDisposed();

         var description = !string.IsNullOrWhiteSpace(Description) && !Equals(Name, Description) ? $" ({Description})" : string.Empty;

         return Name is not null ? $"{Name}{description}, {Category}" : GetType().Name;
      }


      /// <summary>特定の型のハッシュ関数として機能します。</summary>
      /// <returns>現在のオブジェクトのハッシュコード。</returns>
      public override int GetHashCode()
      {
         return NetworkId.GetHashCode();
      }
      

      /// <summary>指定されたオブジェクトが現在のオブジェクトと等しいかどうかを判断します。</summary>
      /// <param name="other">比較する別の <see cref="NetworkInfo"/> インスタンス。</param>
      /// <returns>指定されたオブジェクトが現在のオブジェクトと等しい場合は <c>true</c>、それ以外の場合は <c>false</c>。</returns>
      public bool Equals(NetworkInfo other)
      {
         return null != other && GetType() == other.GetType() &&
                Equals(NetworkId, other.NetworkId);
      }


      /// <summary>指定されたオブジェクトが現在のオブジェクトと等しいかどうかを判断します。</summary>
      /// <param name="obj">比較する別のオブジェクト。</param>
      /// <returns>指定されたオブジェクトが現在のオブジェクトと等しい場合は <c>true</c>、それ以外の場合は <c>false</c>。</returns>
      public override bool Equals(object obj)
      {
         var other = obj as NetworkInfo;

         return null != other && Equals(other);
      }


      /// <summary>== 演算子を実装します。</summary>
      /// <param name="left">A。</param>
      /// <param name="right">B。</param>
      /// <returns>演算子の結果。</returns>
      public static bool operator ==(NetworkInfo left, NetworkInfo right)
      {
         return ReferenceEquals(left, null) && ReferenceEquals(right, null) ||
                !ReferenceEquals(left, null) && !ReferenceEquals(right, null) && left.Equals(right);
      }


      /// <summary>!= 演算子を実装します。</summary>
      /// <param name="left">A。</param>
      /// <param name="right">B。</param>
      /// <returns>演算子の結果。</returns>
      public static bool operator !=(NetworkInfo left, NetworkInfo right)
      {
         return !(left == right);
      }

      #endregion // メソッド
   }
}
