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
using Alphaleonis.Win32.Filesystem;

namespace Alphaleonis.Win32.Network
{
   /// <summary>接続の識別番号、開いているファイルの数、接続時間、接続上のユーザー数、および接続の種類を含みます。</summary>
   [Serializable]
   public sealed class OpenConnectionInfo
   {
      #region プライベートフィールド

      private string _netName;

      #endregion // プライベートフィールド


      #region コンストラクター

      /// <summary>OpenConnectionInfo インスタンスを作成します。</summary>
      internal OpenConnectionInfo(string hostName, NativeMethods.CONNECTION_INFO_1 connectionInfo)
      {
         HostName = hostName;
         Id = connectionInfo.coni1_id;
         ShareType = connectionInfo.coni1_type;
         TotalOpenFiles = connectionInfo.coni1_num_opens;
         TotalUsers = connectionInfo.coni1_num_users;
         ConnectedTime = TimeSpan.FromSeconds(connectionInfo.coni1_time);
         UserName = connectionInfo.coni1_username;
         NetName = connectionInfo.oni1_netname;
      }

      #endregion // コンストラクター

      
      #region メソッド

      /// <summary>共有へのフルパスを返します。</summary>
      /// <returns>このインスタンスを表す文字列。</returns>
      public override string ToString()
      {
         return Id.ToString(CultureInfo.InvariantCulture);
      }


      #endregion // メソッド

      #region プロパティ

      /// <summary>ローカルまたはリモートホスト。</summary>
      [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
      [Obsolete("Use HostName")]
      public string Host { get; private set; }

      /// <summary>この接続情報のホスト名。</summary>
      public string HostName { get; private set; }

      /// <summary>接続識別番号を指定します。</summary>
      public long Id { get; private set; }

      /// <summary>共有の種類。</summary>
      public ShareType ShareType { get; private set; }

      /// <summary>接続の結果として現在開いているファイルの数を指定します。</summary>
      public long TotalOpenFiles { get; private set; }

      /// <summary>接続上のユーザー数を指定します。</summary>
      public long TotalUsers { get; private set; }

      /// <summary>接続が確立されてからの秒数を指定します。</summary>
      [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
      [Obsolete("Use ConnectedTime property.")]
      public long ConnectedSeconds { get; private set; }

      /// <summary>接続が確立されてからの期間を指定します。</summary>
      public TimeSpan ConnectedTime { get; private set; }

      /// <summary>リソースを共有しているサーバーがユーザーレベルのセキュリティで実行されている場合、UserName メンバーはどのユーザーが接続したかを記述します。サーバーが共有レベルのセキュリティで実行されている場合、UserName はどのコンピューター（コンピューター名）が接続したかを記述します。</summary>
      public string UserName { get; private set; }
      
      /// <summary>サーバーの共有リソース名、またはクライアントのコンピューター名もしくは IP アドレスを指定します。このメンバーの値は、関数の修飾子パラメーターとして指定された名前に依存します。</summary>
      public string NetName
      {
         get { return _netName; }

         set { _netName = null != value ? value.ReplaceIgnoreCase(Path.LongPathUncPrefix, string.Empty).Replace(Path.UncPrefix, string.Empty).Trim('[', ']') : null; }
      }

      #endregion // プロパティ
   }
}
