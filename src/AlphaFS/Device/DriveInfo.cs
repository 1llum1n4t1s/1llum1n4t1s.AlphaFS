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
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Security;

namespace Alphaleonis.Win32.Filesystem
{
   /// <summary>ローカルまたはリモートドライブの情報へのアクセスを提供します。</summary>
   /// <remarks>
   /// このクラスはドライブをモデル化し、ドライブ情報を照会するためのメソッドとプロパティを提供します。
   /// DriveInfo を使用して、利用可能なドライブとその種類を判断できます。
   /// ドライブの容量と利用可能な空き領域を照会することもできます。
   /// </remarks>
   [Serializable]
   [SecurityCritical]
   public sealed class DriveInfo
   {
      [NonSerialized] private readonly VolumeInfo _volumeInfo;
      [NonSerialized] private readonly DiskSpaceInfo _dsi;
      [NonSerialized] private bool _initDsie;
      [NonSerialized] private DriveType? _driveType;
      [NonSerialized] private string _dosDeviceName;
      [NonSerialized] private DirectoryInfo _rootDirectory;
      [NonSerialized] private readonly string _name;


      #region コンストラクター

      /// <summary>指定されたドライブの情報へのアクセスを提供します。</summary>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="ArgumentException"/>
      /// <param name="driveName">
      ///   有効なドライブパスまたはドライブ文字。
      ///   <para>大文字または小文字のいずれかを使用できます。</para>
      ///   <para>'a' ～ 'z' または \\server\share 形式のネットワーク共有を指定できます。</para>
      /// </param>
      [SuppressMessage("Microsoft.Design", "CA1062:Validate arguments of public methods", MessageId = "0", Justification = "Utils.IsNullOrWhiteSpace validates arguments.")]
      [SecurityCritical]
      public DriveInfo(string driveName)
      {
         if (Utils.IsNullOrWhiteSpace(driveName))
         {
            throw new ArgumentNullException("driveName");
         }


         driveName = driveName.Length == 1 ? driveName + Path.VolumeSeparatorChar : Path.GetPathRoot(driveName, false);

         if (Utils.IsNullOrWhiteSpace(driveName))
         {
            throw new ArgumentException(Resources.InvalidDriveLetterArgument, "driveName");
         }


         _name = Path.AddTrailingDirectorySeparator(driveName, false);

         // VolumeInfo() の遅延読み込みインスタンスを初期化します。
         _volumeInfo = new VolumeInfo(_name, false, true);

         // DiskSpaceInfo() の遅延読み込みインスタンスを初期化します。
         _dsi = new DiskSpaceInfo(_name, null, false, true);
      }

      #endregion // コンストラクター


      #region プロパティ

      /// <summary>ドライブの利用可能な空き領域の量を示します。</summary>
      /// <returns>ドライブで利用可能な空き領域の量（バイト単位）。</returns>
      /// <remarks>このプロパティはドライブの利用可能な空き領域の量を示します。このプロパティはディスククォータを考慮するため、<see cref="TotalFreeSpace"/> の値とは異なる場合があることに注意してください。</remarks>
      public long AvailableFreeSpace
      {
         get
         {
            GetDeviceInfo(3, 0);
            return null == _dsi ? 0 : _dsi.FreeBytesAvailable;
         }
      }

      /// <summary>NTFS や FAT32 などのファイルシステムの名前を取得します。</summary>
      /// <remarks>ドライブが使用するフォーマットを判断するには DriveFormat を使用してください。</remarks>
      public string DriveFormat
      {
         get { return (string) GetDeviceInfo(0, 1); }
      }


      /// <summary>ドライブの種類を取得します。</summary>
      /// <returns><see cref="System.IO.DriveType"/> 値のいずれか。</returns>
      /// <remarks>
      /// DriveType プロパティは、ドライブが CDRom、Fixed、Unknown、Network、NoRootDirectory、
      /// Ram、Removable、Unknown のいずれかであるかを示します。値は <see cref="System.IO.DriveType"/> 列挙体に一覧されています。
      /// </remarks>
      public DriveType DriveType
      {
         get { return (DriveType) GetDeviceInfo(2, 0); }
      }


      /// <summary>ドライブの準備ができているかどうかを示す値を取得します。</summary>
      /// <returns>ドライブの準備ができている場合は <c>true</c>、それ以外の場合は <c>false</c>。</returns>
      /// <remarks>
      /// IsReady はドライブの準備ができているかどうかを示します。たとえば、CD ドライブに CD が入っているか、
      /// リムーバブルストレージデバイスが読み書き操作の準備ができているかを示します。ドライブの準備ができているかテストせずに
      /// DriveInfo でドライブを照会すると、IOException が発生します。
      ///
      /// AlphaFS の DriveInfo は System.IO.DriveInfo と異なり、TotalSize、TotalFreeSpace、DriveFormat などの
      /// メンバーが取得に失敗しても例外を投げず、既定値 (0 / null / 空文字) を返します。
      /// そのため「取得できなかった」と「本当に 0 / ラベル無し」を戻り値から区別できません。
      /// ドライブが利用可能かどうかは、これらのプロパティを読む前に IsReady で確認してください。
      /// なお IsReady をチェックしてから他のプロパティにアクセスするまでの間に
      /// （アクセスがチェック直後であっても）、ドライブが切断されたりディスクが取り外されたりする可能性があります。
      /// 失敗の詳細は Trace の警告として出力されます。
      /// </remarks>
      public bool IsReady
      {
         get { return File.ExistsCore(null, true, Name, PathFormat.LongFullPath); }
      }


      /// <summary>ドライブの名前を取得します。</summary>
      /// <returns>ドライブの名前。</returns>
      /// <remarks>このプロパティは、C:\ や E:\ などのドライブに割り当てられた名前です。</remarks>
      public string Name
      {
         get { return _name; }
      }


      /// <summary>ドライブのルートディレクトリを取得します。</summary>
      /// <returns>ドライブのルートディレクトリを含む DirectoryInfo オブジェクト。</returns>
      public DirectoryInfo RootDirectory
      {
         get { return (DirectoryInfo) GetDeviceInfo(2, 1); }
      }

      /// <summary>ドライブで利用可能な空き領域の総量を取得します。</summary>
      /// <returns>ドライブで利用可能な空き領域の合計（バイト単位）。</returns>
      /// <remarks>このプロパティは、現在のユーザーが利用可能な量だけでなく、ドライブの空き領域の総量を示します。</remarks>
      public long TotalFreeSpace
      {
         get
         {
            GetDeviceInfo(3, 0);
            return null == _dsi ? 0 : _dsi.TotalNumberOfFreeBytes;
         }
      }


      /// <summary>ドライブのストレージ領域の総サイズを取得します。</summary>
      /// <returns>ドライブの総サイズ（バイト単位）。</returns>
      /// <remarks>このプロパティは、現在のユーザーが利用可能な量だけでなく、ドライブの総サイズ（バイト単位）を示します。</remarks>
      public long TotalSize
      {
         get
         {
            GetDeviceInfo(3, 0);
            return null == _dsi ? 0 : _dsi.TotalNumberOfBytes;
         }
      }


      /// <summary>ドライブのボリュームラベルを取得または設定します。</summary>
      /// <returns>ボリュームラベル。</returns>
      /// <remarks>
      /// ラベルの長さはオペレーティングシステムによって決まります。たとえば、NTFS ではボリュームラベルを
      /// 最大 32 文字にすることができます。<c>null</c> は有効な VolumeLabel であることに注意してください。
      /// </remarks>
      public string VolumeLabel
      {
         get { return (string) GetDeviceInfo(0, 2); }
         set { Volume.SetVolumeLabel(Name, value); }
      }

      /// <summary>[AlphaFS] <see ref="Alphaleonis.Win32.Filesystem.DiskSpaceInfo"/> インスタンスを返します。</summary>
      public DiskSpaceInfo DiskSpaceInfo
      {
         get
         {
            GetDeviceInfo(3, 0);
            return _dsi;
         }
      }


      /// <summary>[AlphaFS] MS-DOS デバイス名。</summary>
      public string DosDeviceName
      {
         get { return (string) GetDeviceInfo(1, 0); }
      }


      /// <summary>[AlphaFS] このドライブが SUBST.EXE / DefineDosDevice ドライブマッピングかどうかを示します。</summary>
      public bool IsDosDeviceSubstitute
      {
         get { return !Utils.IsNullOrWhiteSpace(DosDeviceName) && DosDeviceName.StartsWith(Path.NonInterpretedPathPrefix, StringComparison.OrdinalIgnoreCase); }
      }


      /// <summary>[AlphaFS] このドライブが UNC パスかどうかを示します。</summary>
      public bool IsUnc
      {
         get
         {
            return !IsDosDeviceSubstitute && DriveType == DriveType.Network ||
               
                   // ファイルシステムを持つホストデバイスの処理: FAT/FAT32、UDF (CDRom) など
                   Name.StartsWith(Path.UncPrefix, StringComparison.Ordinal) && DriveType == DriveType.NoRootDirectory && DriveFormat.Equals(DriveType.Unknown.ToString(), StringComparison.OrdinalIgnoreCase);
         }
      }


      /// <summary>[AlphaFS] 指定されたボリューム名が現在のコンピューター上の定義済みボリュームかどうかを判断します。</summary>
      public bool IsVolume
      {
         get { return null != GetDeviceInfo(0, 0); }
      }


      /// <summary>[AlphaFS] ファイルシステムボリュームに関する情報を含みます。</summary>
      /// <returns>ドライブのファイルシステムボリューム情報を含む VolumeInfo オブジェクト。</returns>
      public VolumeInfo VolumeInfo
      {
         get { return (VolumeInfo) GetDeviceInfo(0, 0); }
      }


      #endregion // プロパティ


      #region メソッド

      #region .NET

      /// <summary>コンピューター上のすべての論理ドライブの <see cref="DriveInfo"/> を取得します。</summary>
      /// <returns>コンピューター上の論理ドライブを表す <see cref="Alphaleonis.Win32.Filesystem.DriveInfo"/> 型の配列。</returns>
      [SecurityCritical]
      public static DriveInfo[] GetDrives()
      {
         return Directory.EnumerateLogicalDrivesCore(false, false).ToArray();
      }


      /// <summary>ドライブ名を文字列として返します。</summary>
      /// <returns>ドライブの名前。</returns>
      /// <remarks>このメソッドは Name プロパティを返します。</remarks>
      public override string ToString()
      {
         return _name;
      }

      #endregion // .NET


      /// <summary>[AlphaFS] コンピューター上のすべての論理ドライブのドライブ名を列挙します。</summary>
      /// <param name="fromEnvironment">Environment が認識している論理ドライブを取得します。</param>
      /// <param name="isReady">アクセス可能な（IsReady な）論理ドライブのみを取得します。</param>
      /// <returns>
      ///   コンピューター上の論理ドライブを表す <see cref="Alphaleonis.Win32.Filesystem.DriveInfo"/> 型の IEnumerable。
      /// </returns>      
      [SecurityCritical]
      public static IEnumerable<DriveInfo> EnumerateDrives(bool fromEnvironment, bool isReady)
      {
         return Directory.EnumerateLogicalDrivesCore(fromEnvironment, isReady);
      }


      /// <summary>[AlphaFS] ローカルシステムで最初に利用可能なドライブ文字を取得します。</summary>
      /// <returns><see cref="char"/> としてのドライブ文字。利用可能なドライブ文字がない場合、例外がスローされます。</returns>
      /// <remarks>文字 "A" と "B" はフロッピードライブ用に予約されており、この関数では返されません。</remarks>
      public static char GetFreeDriveLetter()
      {
         return GetFreeDriveLetter(false);
      }


      /// <summary>ローカルシステムで利用可能なドライブ文字を取得します。</summary>
      /// <param name="getLastAvailable"><c>true</c> の場合、最後に利用可能なドライブ文字を取得します。<c>false</c> の場合、最初に利用可能なドライブ文字を取得します。</param>
      /// <returns><see cref="char"/> としてのドライブ文字。利用可能なドライブ文字がない場合、例外がスローされます。</returns>
      /// <remarks>文字 "A" と "B" はフロッピードライブ用に予約されており、この関数では返されません。</remarks>
      /// <exception cref="ArgumentOutOfRangeException">利用可能なドライブ文字がありません。</exception>
      [SuppressMessage("Microsoft.Design", "CA1024:UsePropertiesWhereAppropriate")]
      public static char GetFreeDriveLetter(bool getLastAvailable)
      {
         var freeDriveLetters = "CDEFGHIJKLMNOPQRSTUVWXYZ".Except(Directory.EnumerateLogicalDrivesCore(false, false).Select(d => d.Name[0]));

         try
         {
            return getLastAvailable ? freeDriveLetters.Last() : freeDriveLetters.First();
         }
         catch
         {
            throw new ArgumentOutOfRangeException(Resources.No_Drive_Letters_Available);
         }
      }

      #endregion // メソッド


      #region プライベートメソッド

      /// <summary>指定されたルートファイルまたはディレクトリストリームに関連付けられたファイルシステムとボリュームに関する情報を取得します。</summary>
      [SuppressMessage("Microsoft.Maintainability", "CA1502:AvoidExcessiveComplexity")]
      [SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes")]
      [SecurityCritical]
      private object GetDeviceInfo(int type, int mode)
      {
         try
         {
            switch (type)
            {
               #region ボリューム

               // VolumeInfo プロパティ。
               case 0:
                  if (Utils.IsNullOrWhiteSpace(_volumeInfo.FullPath))
                  {
                     _volumeInfo.Refresh();
                  }

                  switch (mode)
                  {
                     case 0:
                        // IsVolume, VolumeInfo
                        return _volumeInfo;

                     case 1:
                        // DriveFormat
                        return null == _volumeInfo ? DriveType.Unknown.ToString() : _volumeInfo.FileSystemName ?? DriveType.Unknown.ToString();

                     case 2:
                        // VolumeLabel
                        return null == _volumeInfo ? string.Empty : _volumeInfo.Name ?? string.Empty;
                  }

                  break;


               // ボリューム関連。
               case 1:
                  switch (mode)
                  {
                     case 0:
                        // DosDeviceName
                        return _dosDeviceName ?? (_dosDeviceName = Volume.GetVolumeDeviceName(Name));
                  }

                  break;

               #endregion // ボリューム


               #region ドライブ

               // ドライブ関連。
               case 2:
                  switch (mode)
                  {
                     case 0:
                        // DriveType
                        return _driveType ?? (_driveType = Volume.GetDriveType(Name));

                     case 1:
                        // RootDirectory
                        return _rootDirectory ?? (_rootDirectory = new DirectoryInfo(null, Name, PathFormat.RelativePath));
                  }

                  break;

               // DiskSpaceInfo 関連。
               case 3:
                  switch (mode)
                  {
                     case 0:
                        // AvailableFreeSpace, TotalFreeSpace, TotalSize, DiskSpaceInfo
                        if (!_initDsie)
                        {
                           _dsi.Refresh();
                           _initDsie = true;
                        }

                        break;
                  }

                  break;

               #endregion // ドライブ
            }
         }
         catch (Exception ex)
         {
            // このメソッドは「取得できなければ既定値」という契約で公開プロパティから使われるため、
            // 例外はここで飲み込む (投げるように変えると TotalSize などが 0 を返す前提の利用者を壊す)。
            // ただし何も残さないと、アクセス拒否やドライブ未準備といった本当の失敗と
            // 「本当に 0 バイト / ラベル無し」が呼び出し側から区別できず、原因に到達できない。
            // 少なくとも診断できるよう Trace に残す。
            Trace.TraceWarning("DriveInfo.GetDeviceInfo(type: {0}, mode: {1}) failed for [{2}]: {3}", type, mode, Name, ex.Message);
         }

         return type == 0 && mode > 0 ? string.Empty : null;
      }
      
      #endregion // プライベート
   }
}
