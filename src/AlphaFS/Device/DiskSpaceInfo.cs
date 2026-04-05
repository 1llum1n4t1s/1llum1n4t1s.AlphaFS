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
using System.Runtime.InteropServices;
using System.Security;

namespace Alphaleonis.Win32.Filesystem
{
   /// <summary>ディスクボリュームで利用可能な領域の量に関する情報を取得します。これには、総領域量、
   /// 総空き領域量、および呼び出しスレッドに関連付けられたユーザーが利用可能な総空き領域量が含まれます。
   /// <para>このクラスは継承できません。</para>
   /// </summary>
   [Serializable]
   [SecurityCritical]
   public sealed class DiskSpaceInfo
   {
      [NonSerialized] private readonly bool _initGetClusterInfo = true;
      [NonSerialized] private readonly bool _initGetSpaceInfo = true;
      [NonSerialized] private readonly CultureInfo _cultureInfo = CultureInfo.CurrentCulture;
      [NonSerialized] private readonly bool _continueOnAccessError;


      /// <summary>DiskSpaceInfo インスタンスを初期化します。</summary>
      /// <param name="drivePath">有効なドライブパスまたはドライブ文字。大文字または小文字の 'a' ～ 'z'、または \\server\share 形式のネットワーク共有を指定できます。</param>
      /// <Remark>これは遅延読み込みオブジェクトです。プロパティにアクセスする前に <see cref="Refresh()"/> を呼び出して全プロパティを設定してください。</Remark>
      [SuppressMessage("Microsoft.Design", "CA1062:Validate arguments of public methods", MessageId = "0", Justification = "Utils.IsNullOrWhiteSpace validates arguments.")]
      [SecurityCritical]
      public DiskSpaceInfo(string drivePath)
      {
         if (Utils.IsNullOrWhiteSpace(drivePath))
         {
            throw new ArgumentNullException("drivePath");
         }


         drivePath = drivePath.Length == 1 ? drivePath + Path.VolumeSeparatorChar : Path.GetPathRoot(drivePath, false);

         if (Utils.IsNullOrWhiteSpace(drivePath))
         {
            throw new ArgumentException(Resources.InvalidDriveLetterArgument, "drivePath");
         }


         // MSDN:
         // このパラメーターが UNC 名の場合、末尾にバックスラッシュを含める必要があります（例: "\\MyServer\MyShare\"）。
         // また、ドライブ指定には末尾にバックスラッシュが必要です（例: "C:\"）。
         // 呼び出し元アプリケーションには、このディレクトリに対する FILE_LIST_DIRECTORY アクセス権が必要です。
         DriveName = Path.AddTrailingDirectorySeparator(drivePath, false);
      }

      
      /// <summary>DiskSpaceInfo インスタンスを初期化します。</summary>
      /// <param name="drivePath">有効なドライブパスまたはドライブ文字。大文字または小文字の 'a' ～ 'z'、または \\server\share 形式のネットワーク共有を指定できます。</param>
      /// <param name="spaceInfoType"><c>null</c> はサイズ情報とディスククラスター情報の両方を取得します。<c>true</c> はディスククラスター情報のみを取得、<c>false</c> はサイズ情報のみを取得します。</param>
      /// <param name="refresh">オブジェクトの状態を更新します。</param>
      /// <param name="continueOnException"><c>true</c> はリソース不足などの失敗から発生する可能性のある例外を抑制します。</param>
      [SecurityCritical]
      public DiskSpaceInfo(string drivePath, bool? spaceInfoType, bool refresh, bool continueOnException) : this(drivePath)
      {
         if (spaceInfoType == null)
         {
            _initGetSpaceInfo = true;
            _initGetClusterInfo = true;
         }

         else
         {
            _initGetSpaceInfo = (bool) !spaceInfoType;
            _initGetClusterInfo = (bool) spaceInfoType;
         }

         _continueOnAccessError = continueOnException;

         if (refresh)
         {
            Refresh();
         }
      }


      /// <summary>ドライブの利用可能な空き領域の量をパーセンテージで示します。</summary>
      public string AvailableFreeSpacePercent
      {
         get
         {
            return PercentCalculate(FreeBytesAvailable, 0, TotalNumberOfBytes).ToString("0.##", _cultureInfo) + "%";
         }
      }


      /// <summary>ドライブの利用可能な空き領域の量を単位サイズで示します。</summary>
      public string AvailableFreeSpaceUnitSize
      {
         get { return Utils.UnitSizeToText(TotalNumberOfFreeBytes, _cultureInfo); }
      }


      /// <summary>クラスターサイズを返します。</summary>
      public long ClusterSize
      {
         get { return (long) SectorsPerCluster * BytesPerSector; }
      }


      /// <summary>ドライブの名前を取得します。</summary>
      /// <returns>ドライブの名前。</returns>
      /// <remarks>このプロパティは、C:\ や E:\ などのドライブに割り当てられた名前です。</remarks>
      public string DriveName { get; private set; }


      /// <summary>呼び出しスレッドに関連付けられたユーザーが利用可能なディスクの総バイト数を単位サイズで示します。</summary>
      public string TotalSizeUnitSize
      {
         get { return Utils.UnitSizeToText(TotalNumberOfBytes, _cultureInfo); }
      }


      /// <summary>ドライブの使用済み領域の量をパーセンテージで示します。</summary>
      public string UsedSpacePercent
      {
         get
         {
            return PercentCalculate(TotalNumberOfBytes - FreeBytesAvailable, 0, TotalNumberOfBytes).ToString("0.##", _cultureInfo) + "%";
         }
      }


      /// <summary>ドライブの使用済み領域の量を単位サイズで示します。</summary>
      public string UsedSpaceUnitSize
      {
         get { return Utils.UnitSizeToText(TotalNumberOfBytes - FreeBytesAvailable, _cultureInfo); }
      }


      /// <summary>呼び出しスレッドに関連付けられたユーザーが利用可能なディスクの空きバイト数の合計。</summary>
      public long FreeBytesAvailable { get; private set; }


      /// <summary>呼び出しスレッドに関連付けられたユーザーが利用可能なディスクの総バイト数。</summary>
      public long TotalNumberOfBytes { get; private set; }


      /// <summary>ディスクの空きバイト数の合計。</summary>
      public long TotalNumberOfFreeBytes { get; private set; }


      /// <summary>セクターあたりのバイト数。</summary>
      public int BytesPerSector { get; private set; }


      /// <summary>呼び出しスレッドに関連付けられたユーザーが利用可能なディスクの空きクラスター数の合計。</summary>
      public int NumberOfFreeClusters { get; private set; }


      /// <summary>クラスターあたりのセクター数。</summary>
      public int SectorsPerCluster { get; private set; }


      /// <summary>呼び出しスレッドに関連付けられたユーザーが利用可能なディスクのクラスター数の合計。
      /// ユーザーごとのディスククォータが使用されている場合、この値はディスクのクラスター総数より少ない場合があります。
      /// </summary>
      public long TotalNumberOfClusters { get; private set; }




      /// <summary>オブジェクトの状態を更新します。</summary>
      public void Refresh()
      {
         Reset();

         using (new NativeMethods.ChangeErrorMode(NativeMethods.ErrorMode.FailCriticalErrors))
         {
            int lastError;


            // サイズ情報を取得します。

            if (_initGetSpaceInfo)
            {

               var success = NativeMethods.GetDiskFreeSpaceEx(DriveName, out var freeBytesAvailable, out var totalNumberOfBytes, out var totalNumberOfFreeBytes);

               lastError = Marshal.GetLastWin32Error();

               if (!success && !_continueOnAccessError && lastError != Win32Errors.ERROR_NOT_READY)
               {
                  NativeError.ThrowException(lastError, DriveName);
               }


               FreeBytesAvailable = freeBytesAvailable;
               TotalNumberOfBytes = totalNumberOfBytes;
               TotalNumberOfFreeBytes = totalNumberOfFreeBytes;
            }


            // クラスター情報を取得します。

            if (_initGetClusterInfo)
            {

               var success = NativeMethods.GetDiskFreeSpace(DriveName, out var sectorsPerCluster, out var bytesPerSector, out var numberOfFreeClusters, out var totalNumberOfClusters);

               lastError = Marshal.GetLastWin32Error();

               if (!success && !_continueOnAccessError && lastError != Win32Errors.ERROR_NOT_READY)
               {
                  NativeError.ThrowException(lastError, DriveName);
               }


               BytesPerSector = bytesPerSector;
               NumberOfFreeClusters = numberOfFreeClusters;
               SectorsPerCluster = sectorsPerCluster;
               TotalNumberOfClusters = totalNumberOfClusters;
            }
         }
      }


      /// <summary>すべての <see ref="Alphaleonis.Win32.Filesystem.DiskSpaceInfo"/> プロパティを 0 に初期化します。</summary>
      private void Reset()
      {
         if (_initGetSpaceInfo)
         {
            FreeBytesAvailable = 0;
            TotalNumberOfBytes = 0;
            TotalNumberOfFreeBytes = 0;
         }


         if (_initGetClusterInfo)
         {
            BytesPerSector = 0;
            NumberOfFreeClusters = 0;
            SectorsPerCluster = 0;
            TotalNumberOfClusters = 0;
         }
      }


      /// <summary>ドライブ名を返します。</summary>
      /// <returns>このオブジェクトを表す文字列。</returns>
      public override string ToString()
      {
         return DriveName;
      }


      /// <summary>パーセンテージ値を計算します。</summary>
      private static double PercentCalculate(double currentValue, double minimumValue, double maximumValue)
      {
         return currentValue < 0 || maximumValue <= 0 ? 0 : currentValue * 100 / (maximumValue - minimumValue);
      }
   }
}
