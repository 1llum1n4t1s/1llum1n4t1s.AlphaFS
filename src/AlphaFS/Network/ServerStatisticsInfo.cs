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
   /// <summary>サーバーサービスの動作統計を含みます.</summary>
   [Serializable]
   public sealed class ServerStatisticsInfo : IEquatable<ServerStatisticsInfo>
   {
      #region Fields

      [NonSerialized] private DateTime _dateTimeNowUtc;
      [NonSerialized] private NativeMethods.STAT_SERVER_0 _serverStat;

      #endregion // Fields

      
      #region コンストラクター

      /// <summary>Create a ServerStatisticsInfo instance from the local host.</summary>
      public ServerStatisticsInfo() : this(Environment.MachineName, null)
      {
      }


      /// <summary>Create a ServerStatisticsInfo instance from the specified host name.</summary>
      /// <param name="hostName">The host name.</param>
      public ServerStatisticsInfo(string hostName) : this(hostName, null)
      {
      }


      /// <summary>Create a ServerStatisticsInfo instance from the specified host name.</summary>
      internal ServerStatisticsInfo(string hostName, NativeMethods.STAT_SERVER_0? serverStat)
      {
         HostName = !Utils.IsNullOrWhiteSpace(hostName) ? hostName : Environment.MachineName;

         if (serverStat.HasValue)
         {
            _dateTimeNowUtc = DateTime.UtcNow;

            _serverStat = (NativeMethods.STAT_SERVER_0) serverStat;
         }

         else
         {
            Refresh();
         }
      }

      #endregion // コンストラクター


      #region プロパティ

      /// <summary>の数 server access permission errors.</summary>
      public int AccessPermissionErrors
      {
         get { return (int) _serverStat.sts0_permerrors; }
      }


      /// <summary>The average server response time.</summary>
      public TimeSpan AverageResponseTime
      {
         get { return TimeSpan.FromMilliseconds(_serverStat.sts0_avresponse); }
      }


      /// <summary>の数 times the server required a big buffer but failed to allocate one. This value indicates that the server parameters may need adjustment.</summary>
      public int BufferAllocationFailed
      {
         get { return (int) _serverStat.sts0_bigbufneed; }
      }


      /// <summary>の数 times the server required a request buffer but failed to allocate one. This value indicates that the server parameters may need adjustment.</summary>
      public int BufferRequestFailed
      {
         get { return (int) _serverStat.sts0_reqbufneed; }
      }


      /// <summary>の数 server bytes received from the network.</summary>
      public long BytesReceived
      {
         get { return Filesystem.NativeMethods.ToLong(_serverStat.sts0_bytesrcvd_high, _serverStat.sts0_bytesrcvd_low); }
      }


      /// <summary>の数 server bytes received from the network, formatted as a unit size.</summary>
      public string BytesReceivedUnitSize
      {
         get { return Utils.UnitSizeToText(BytesReceived); }
      }


      /// <summary>の数 server bytes sent to the network.</summary>
      public long BytesSent
      {
         get { return Filesystem.NativeMethods.ToLong(_serverStat.sts0_bytessent_high, _serverStat.sts0_bytessent_low); }
      }


      /// <summary>の数 server bytes sent to the network, formatted as a unit size.</summary>
      public string BytesSentUnitSize
      {
         get { return Utils.UnitSizeToText(BytesSent); }
      }
      

      /// <summary>の数 times a server device is opened.</summary>
      public int DevicesOpened
      {
         get { return (int) _serverStat.sts0_devopens; }
      }


      /// <summary>の数 times a file is opened on a server. This includes the number of times named pipes are opened.</summary>
      public int FilesOpened
      {
         get { return (int) _serverStat.sts0_fopens; }
      }


      /// <summary>The host name from where the statistics are gathered.</summary>
      public string HostName { get; private set; }


      /// <summary>の数 server print jobs spooled.</summary>
      public int JobsQueued
      {
         get { return (int) _serverStat.sts0_jobsqueued; }
      }


      /// <summary>の数 server password violations.</summary>
      public int PasswordViolations
      {
         get { return (int) _serverStat.sts0_pwerrors; }
      }


      /// <summary>の数 times the server sessions failed with an error.</summary>
      public int SessionsFailed
      {
         get { return (int) _serverStat.sts0_serrorout; }
      }


      /// <summary>の数 times the server session started.</summary>
      public int SessionsStarted
      {
         get { return (int) _serverStat.sts0_sopens; }
      }


      /// <summary>の数 times the server session automatically disconnected.</summary>
      public int SessionsTimedOut
      {
         get { return (int) _serverStat.sts0_stimedout; }
      }


      /// <summary>The local time when statistics collection started or when the statistics were last cleared.</summary>
      public DateTime StatisticsStartTime
      {
         get { return StatisticsStartTimeUtc.ToLocalTime(); }
      }


      /// <summary>The time when statistics collection started or when the statistics were last cleared.</summary>
      public DateTime StatisticsStartTimeUtc
      {
         get { return new DateTime((_dateTimeNowUtc - new DateTime(_serverStat.sts0_start, DateTimeKind.Utc)).Ticks, DateTimeKind.Utc); }
      }


      /// <summary>の数 server system errors.</summary>
      public int SystemErrors
      {
         get { return (int) _serverStat.sts0_syserrors; }
      }

      #endregion // プロパティ


      #region メソッド

      /// <summary>Refreshes the state of the object.</summary>
      public void Refresh()
      {
         _dateTimeNowUtc = DateTime.UtcNow;

         _serverStat = Host.GetNetStatisticsNative<NativeMethods.STAT_SERVER_0>(true, HostName);
      }


      /// <summary>Returns the local time when statistics collection started or when the statistics were last cleared.</summary>
      /// <returns>このインスタンスを表す文字列。</returns>
      public override string ToString()
      {
         return HostName;
      }


      /// <summary>特定の型のハッシュ関数として機能します。</summary>
      /// <returns>現在のオブジェクトのハッシュコード。</returns>
      public override int GetHashCode()
      {
         return Utils.CombineHashCodesOf(HostName, BytesSent, StatisticsStartTime);
      }
      

      /// <summary>指定されたオブジェクトが現在のオブジェクトと等しいかどうかを判断します。</summary>
      /// <param name="other">Another <see cref="ServerStatisticsInfo"/> instance to compare to.</param>
      /// <returns><c>true</c> 指定されたオブジェクトが現在のオブジェクトと等しい場合。それ以外の場合は <c>false</c>.</returns>
      public bool Equals(ServerStatisticsInfo other)
      {
         return null != other && GetType() == other.GetType() &&
                Equals(HostName, other.HostName) &&
                Equals(BytesSent, other.BytesSent) &&
                Equals(BytesReceived, other.BytesReceived) &&
                Equals(StatisticsStartTimeUtc, other.StatisticsStartTimeUtc);
      }


      /// <summary>指定されたオブジェクトが現在のオブジェクトと等しいかどうかを判断します。</summary>
      /// <param name="obj">比較する別のオブジェクト。</param>
      /// <returns><c>true</c> 指定されたオブジェクトが現在のオブジェクトと等しい場合。それ以外の場合は <c>false</c>.</returns>
      public override bool Equals(object obj)
      {
         var other = obj as ServerStatisticsInfo;

         return null != other && Equals(other);
      }


      /// <summary>== 演算子を実装します</summary>
      /// <param name="left">A.</param>
      /// <param name="right">B.</param>
      /// <returns>演算子の結果。</returns>
      public static bool operator ==(ServerStatisticsInfo left, ServerStatisticsInfo right)
      {
         return ReferenceEquals(left, null) && ReferenceEquals(right, null) ||
                !ReferenceEquals(left, null) && !ReferenceEquals(right, null) && left.Equals(right);
      }


      /// <summary>!= 演算子を実装します</summary>
      /// <param name="left">A.</param>
      /// <param name="right">B.</param>
      /// <returns>演算子の結果。</returns>
      public static bool operator !=(ServerStatisticsInfo left, ServerStatisticsInfo right)
      {
         return !(left == right);
      }

      #endregion // メソッド
   }
}
