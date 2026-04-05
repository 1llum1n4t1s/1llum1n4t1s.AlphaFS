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

namespace Alphaleonis.Win32.Network
{
   /// <summary>ワークステーションサービスの動作統計を含みます.</summary>
   [Serializable]
   public sealed class WorkstationStatisticsInfo : IEquatable<WorkstationStatisticsInfo>
   {
      #region Fields

      [NonSerialized] private NativeMethods.STAT_WORKSTATION_0 _workstationStat;

      #endregion // Fields


      #region コンストラクター

      /// <summary>Create a WorkstationStatisticsInfo instance from the local host.</summary>
      public WorkstationStatisticsInfo() : this(Environment.MachineName, null)
      {
      }


      /// <summary>Create a WorkstationStatisticsInfo instance from the specified host name.</summary>
      /// <param name="hostName">The host name.</param>

      public WorkstationStatisticsInfo(string hostName) : this(hostName, null)
      {
      }


      /// <summary>Create a ServerStatisticsInfo instance from the specified host name.</summary>
      internal WorkstationStatisticsInfo(string hostName, NativeMethods.STAT_WORKSTATION_0? workstationStat)
      {
         HostName = !Utils.IsNullOrWhiteSpace(hostName) ? hostName : Environment.MachineName;

         if (workstationStat.HasValue)
         {
            _workstationStat = (NativeMethods.STAT_WORKSTATION_0) workstationStat;
         }

         else
         {
            Refresh();
         }
      }

      #endregion // コンストラクター


      #region プロパティ

      /// <summary>の合計数 bytes received by the workstation.</summary>
      public long BytesReceived
      {
         get { return _workstationStat.BytesReceived; }
      }


      /// <summary>の合計数 bytes received by the workstation, formatted as a unit size.</summary>
      public string BytesReceivedUnitSize
      {
         get { return Utils.UnitSizeToText(BytesReceived); }
      }


      /// <summary>の合計数 bytes transmitted by the workstation.</summary>
      public long BytesTransmitted
      {
         get { return _workstationStat.BytesTransmitted; }
      }


      /// <summary>の合計数 bytes transmitted by the workstation, formatted as a unit size.</summary>
      public string BytesTransmittedUnitSize
      {
         get { return Utils.UnitSizeToText(BytesTransmitted); }
      }


      /// <summary>の合計数 bytes that have been read by cache I/O requests.</summary>
      public long CacheReadBytesRequested
      {
         get { return _workstationStat.CacheReadBytesRequested; }
      }


      /// <summary>の合計数 bytes that have been read by cache I/O requests, formatted as a unit size.</summary>
      public string CacheReadBytesRequestedUnitSize
      {
         get { return Utils.UnitSizeToText(CacheReadBytesRequested); }
      }


      /// <summary>の合計数 bytes that have been written by cache I/O requests.</summary>
      public long CacheWriteBytesRequested
      {
         get { return _workstationStat.CacheWriteBytesRequested; }
      }


      /// <summary>の合計数 bytes that have been written by cache I/O requests, formatted as a unit size.</summary>
      public string CacheWriteBytesRequestedUnitSize
      {
         get { return Utils.UnitSizeToText(CacheWriteBytesRequested); }
      }


      /// <summary>の合計数 connections to servers supporting the PCNET dialect that have succeeded.</summary>
      public int CoreConnects
      {
         get { return (int) _workstationStat.CoreConnects; }
      }


      /// <summary>の数 current requests that have not been completed.</summary>
      public int CurrentCommands
      {
         get { return (int) _workstationStat.CurrentCommands; }
      }


      /// <summary>の合計数 network operations that failed to complete.</summary>
      public int FailedCompletionOperations
      {
         get { return (int) _workstationStat.FailedCompletionOperations; }
      }


      /// <summary>の数 times the workstation attempted to create a session but failed.</summary>
      public int FailedSessions
      {
         get { return (int) _workstationStat.FailedSessions; }
      }


      /// <summary>の合計数 failed network connections for the workstation.</summary>
      public int FailedUseCount
      {
         get { return (int) _workstationStat.FailedUseCount; }
      }


      /// <summary>The host name from where the statistics are gathered.</summary>
      public string HostName { get; private set; }


      /// <summary>の合計数 network operations that failed to begin.</summary>
      public int InitiallyFailedOperations
      {
         get { return (int) _workstationStat.InitiallyFailedOperations; }
      }


      /// <summary>の合計数 sessions that have expired on the workstation.</summary>
      public int HungSessions
      {
         get { return (int) _workstationStat.HungSessions; }
      }


      /// <summary>の合計数 connections to servers supporting the LanManager 2.0 dialect that have succeeded.</summary>
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Lanman")]
      public int Lanman20Connects
      {
         get { return (int) _workstationStat.Lanman20Connects; }
      }


      /// <summary>の合計数 connections to servers supporting the LanManager 2.1 dialect that have succeeded.</summary>
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Lanman")]
      public int Lanman21Connects
      {
         get { return (int) _workstationStat.Lanman21Connects; }
      }


      /// <summary>の合計数 connections to servers supporting the NTLM dialect that have succeeded.</summary>
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Nt")]
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Lanman")]
      [SuppressMessage("Microsoft.Naming", "CA1709:IdentifiersShouldBeCasedCorrectly", MessageId = "Nt")]
      public int LanmanNtConnects
      {
         get { return (int) _workstationStat.LanmanNtConnects; }
      }


      /// <summary>の合計数 read requests the workstation has sent to servers that are greater than twice the size of the server's negotiated buffer size.</summary>
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Smbs")]
      public int LargeReadSmbs
      {
         get { return (int) _workstationStat.LargeReadSmbs; }
      }


      /// <summary>の合計数 write requests the workstation has sent to servers that are greater than twice the size of the server's negotiated buffer size.</summary>
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Smbs")]
      public int LargeWriteSmbs
      {
         get { return (int) _workstationStat.LargeWriteSmbs; }
      }


      /// <summary>の合計数 network errors received by the workstation.</summary>
      public int NetworkErrors
      {
         get { return (int) _workstationStat.NetworkErrors; }
      }


      /// <summary>The total amount of bytes that have been read by disk I/O requests.</summary>
      public long NetworkReadBytesRequested
      {
         get { return _workstationStat.NetworkReadBytesRequested; }
      }


      /// <summary>The total amount of bytes that have been read by disk I/O requests, formatted as a unit size.</summary>
      public string NetworkReadBytesRequestedUnitSize
      {
         get { return Utils.UnitSizeToText(NetworkReadBytesRequested); }
      }


      /// <summary>の合計数 bytes that have been written by disk I/O requests.</summary>
      public long NetworkWriteBytesRequested
      {
         get { return _workstationStat.NetworkWriteBytesRequested; }
      }


      /// <summary>の合計数 bytes that have been written by disk I/O requests, formatted as a unit size.</summary>
      public string NetworkWriteBytesRequestedUnitSize
      {
         get { return Utils.UnitSizeToText(NetworkWriteBytesRequested); }
      }


      /// <summary>の合計数 bytes that have been read by non-paging I/O requests.</summary>
      public long NonPagingReadBytesRequested
      {
         get { return _workstationStat.NonPagingReadBytesRequested; }
      }


      /// <summary>の合計数 bytes that have been read by non-paging I/O requests, formatted as a unit size.</summary>
      public string NonPagingReadBytesRequestedUnitSize
      {
         get { return Utils.UnitSizeToText(NonPagingReadBytesRequested); }
      }


      /// <summary>の合計数 bytes that have been written by non-paging I/O requests.</summary>
      public long NonPagingWriteBytesRequested
      {
         get { return _workstationStat.NonPagingWriteBytesRequested; }
      }


      /// <summary>の合計数 bytes that have been written by non-paging I/O requests, formatted as a unit size.</summary>
      public string NonPagingWriteBytesRequestedUnitSize
      {
         get { return Utils.UnitSizeToText(NonPagingWriteBytesRequested); }
      }


      /// <summary>の合計数 bytes that have been read by paging I/O requests.</summary>
      public long PagingReadBytesRequested
      {
         get { return _workstationStat.PagingReadBytesRequested; }
      }


      /// <summary>の合計数 bytes that have been read by paging I/O requests, formatted as a unit size.</summary>
      public string PagingReadBytesRequestedUnitSize
      {
         get { return Utils.UnitSizeToText(PagingReadBytesRequested); }
      }


      /// <summary>の合計数 bytes that have been written by paging I/O requests.</summary>
      public long PagingWriteBytesRequested
      {
         get { return _workstationStat.PagingWriteBytesRequested; }
      }


      /// <summary>の合計数 bytes that have been written by paging I/O requests, formatted as a unit size.</summary>
      public string PagingWriteBytesRequestedUnitSize
      {
         get { return Utils.UnitSizeToText(PagingWriteBytesRequested); }
      }


      /// <summary>の合計数 random access reads initiated by the workstation.</summary>
      public int RandomReadOperations
      {
         get { return (int) _workstationStat.RandomReadOperations; }
      }


      /// <summary>の合計数 random access writes initiated by the workstation.</summary>
      public int RandomWriteOperations
      {
         get { return (int) _workstationStat.RandomWriteOperations; }
      }


      /// <summary>の合計数 raw read requests made by the workstation that have been denied.</summary>
      public int RawReadsDenied
      {
         get { return (int) _workstationStat.RawReadsDenied; }
      }


      /// <summary>の合計数 raw write requests made by the workstation that have been denied.</summary>
      public int RawWritesDenied
      {
         get { return (int) _workstationStat.RawWritesDenied; }
      }


      /// <summary>の合計数 read operations initiated by the workstation.</summary>
      public int ReadOperations
      {
         get { return (int) _workstationStat.ReadOperations; }
      }


      /// <summary>の合計数 read requests the workstation has sent to servers.</summary>
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Smbs")]
      public int ReadSmbs
      {
         get { return (int) _workstationStat.ReadSmbs; }
      }


      /// <summary>の合計数 connections that have failed.</summary>
      public int Reconnects
      {
         get { return (int) _workstationStat.Reconnects; }
      }


      /// <summary>の数 times the workstation was disconnected by a network server.</summary>
      public int ServerDisconnects
      {
         get { return (int) _workstationStat.ServerDisconnects; }
      }


      /// <summary>の合計数 workstation sessions that were established.</summary>
      public int Sessions
      {
         get { return (int) _workstationStat.Sessions; }
      }


      /// <summary>の合計数 read requests the workstation has sent to servers that are less than 1/4 of the size of the server's negotiated buffer size.</summary>
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Smbs")]
      public int SmallReadSmbs
      {
         get { return (int) _workstationStat.SmallReadSmbs; }
      }


      /// <summary>の合計数 write requests the workstation has sent to servers that are less than 1/4 of the size of the server's negotiated buffer size.</summary>
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Smbs")]
      public int SmallWriteSmbs
      {
         get { return (int) _workstationStat.SmallWriteSmbs; }
      }


      /// <summary>The local time statistics collection started. This member also indicates when statistics for the workstations were last cleared.</summary>
      public DateTime StatisticsStartTime
      {
         get { return StatisticsStartTimeUtc.ToLocalTime(); }
      }


      /// <summary>The time statistics collection started. This member also indicates when statistics for the workstations were last cleared.</summary>
      public DateTime StatisticsStartTimeUtc
      {
         get { return DateTime.FromFileTimeUtc(_workstationStat.StatisticsStartTime); }
      }


      /// <summary>の合計数 server message blocks (SMBs) received by the workstation.</summary>
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Smbs")]
      public long SmbsReceived
      {
         get { return _workstationStat.SmbsReceived; }
      }


      /// <summary>の合計数 SMBs transmitted by the workstation.</summary>
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Smbs")]
      public long SmbsTransmitted
      {
         get { return _workstationStat.SmbsTransmitted; }
      }
      

      /// <summary>の合計数 network connections established by the workstation.</summary>
      public int UseCount
      {
         get { return (int) _workstationStat.UseCount; }
      }


      /// <summary>の合計数 write operations initiated by the workstation.</summary>
      public int WriteOperations
      {
         get { return (int) _workstationStat.WriteOperations; }
      }


      /// <summary>の合計数 write requests the workstation has sent to servers.</summary>
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Smbs")]
      public int WriteSmbs
      {
         get { return (int) _workstationStat.WriteSmbs; }
      }
      
      #endregion // プロパティ


      #region メソッド

      /// <summary>Refreshes the state of the object.</summary>
      public void Refresh()
      {
         _workstationStat = Host.GetNetStatisticsNative<NativeMethods.STAT_WORKSTATION_0>(false, HostName);
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
         return Utils.CombineHashCodesOf(HostName, BytesTransmitted, StatisticsStartTime);
      }
      

      /// <summary>指定されたオブジェクトが現在のオブジェクトと等しいかどうかを判断します。</summary>
      /// <param name="other">Another <see cref="WorkstationStatisticsInfo"/> instance to compare to.</param>
      /// <returns><c>true</c> 指定されたオブジェクトが現在のオブジェクトと等しい場合。それ以外の場合は <c>false</c>.</returns>
      public bool Equals(WorkstationStatisticsInfo other)
      {
         return null != other && GetType() == other.GetType() &&
                Equals(HostName, other.HostName) &&
                Equals(BytesTransmitted, other.BytesTransmitted) &&
                Equals(BytesReceived, other.BytesReceived) &&
                Equals(StatisticsStartTimeUtc, other.StatisticsStartTimeUtc);
      }


      /// <summary>指定されたオブジェクトが現在のオブジェクトと等しいかどうかを判断します。</summary>
      /// <param name="obj">比較する別のオブジェクト。</param>
      /// <returns><c>true</c> 指定されたオブジェクトが現在のオブジェクトと等しい場合。それ以外の場合は <c>false</c>.</returns>
      public override bool Equals(object obj)
      {
         var other = obj as WorkstationStatisticsInfo;

         return null != other && Equals(other);
      }


      /// <summary>== 演算子を実装します</summary>
      /// <param name="left">A.</param>
      /// <param name="right">B.</param>
      /// <returns>演算子の結果。</returns>
      public static bool operator ==(WorkstationStatisticsInfo left, WorkstationStatisticsInfo right)
      {
         return ReferenceEquals(left, null) && ReferenceEquals(right, null) ||
                !ReferenceEquals(left, null) && !ReferenceEquals(right, null) && left.Equals(right);
      }


      /// <summary>!= 演算子を実装します</summary>
      /// <param name="left">A.</param>
      /// <param name="right">B.</param>
      /// <returns>演算子の結果。</returns>
      public static bool operator !=(WorkstationStatisticsInfo left, WorkstationStatisticsInfo right)
      {
         return !(left == right);
      }

      #endregion // メソッド
   }
}
