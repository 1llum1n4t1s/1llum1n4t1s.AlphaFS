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
using Microsoft.Win32.SafeHandles;
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Security;
using System.Security.AccessControl;
using System.Text;

namespace Alphaleonis.Win32.Filesystem
{
   /// <summary>ローカルまたはリモートホストからデバイスリソース情報を取得するための静的メソッドを提供します。</summary>
   public static class Device
   {
      #region デバイスの列挙

      /// <summary>[AlphaFS] ローカルホスト上の利用可能な全デバイスを列挙します。</summary>
      /// <returns>ローカルホストからの <see cref="DeviceGuid"/> 型の <see cref="IEnumerable{DeviceInfo}"/> インスタンス。</returns>
      /// <param name="deviceGuid"><see cref="DeviceGuid"/> デバイスのいずれか。</param>
      [SecurityCritical]
      public static IEnumerable<DeviceInfo> EnumerateDevices(DeviceGuid deviceGuid)
      {
         return EnumerateDevicesCore(null, deviceGuid, true);
      }


      /// <summary>[AlphaFS] ローカルまたはリモートホスト上の <see cref="DeviceGuid"/> 型の利用可能な全デバイスを列挙します。</summary>
      /// <returns>指定された <paramref name="hostName"/> に対する <see cref="DeviceGuid"/> 型の <see cref="IEnumerable{DeviceInfo}"/> インスタンス。</returns>
      /// <param name="hostName">デバイスが存在するローカルまたはリモートホストの名前。<c>null</c> はローカルホストを参照します。</param>
      /// <param name="deviceGuid"><see cref="DeviceGuid"/> デバイスのいずれか。</param>
      [SecurityCritical]
      public static IEnumerable<DeviceInfo> EnumerateDevices(string hostName, DeviceGuid deviceGuid)
      {
         return EnumerateDevicesCore(hostName, deviceGuid, true);
      }




      /// <summary>[AlphaFS] ローカルまたはリモートホスト上の利用可能な全デバイスを列挙します。</summary>
      [SecurityCritical]
      internal static IEnumerable<DeviceInfo> EnumerateDevicesCore(string hostName, DeviceGuid deviceGuid, bool getAllProperties)
      {
         if (Utils.IsNullOrWhiteSpace(hostName))
         {
            hostName = Environment.MachineName;
         }


         // CM_Connect_Machine()
         // MSDN注記: Windows 8 および Windows Server 2012 以降、リモートマシンへのアクセス機能は削除されました。
         // これらのバージョンの Windows では、リモートマシンにアクセスできません。
         // http://msdn.microsoft.com/en-us/library/windows/hardware/ff537948%28v=vs.85%29.aspx


         var lastError = NativeMethods.CM_Connect_Machine(Host.GetUncName(hostName), out var safeMachineHandle);

         NativeMethods.IsValidHandle(safeMachineHandle, lastError);


         var classGuid = new Guid(Utils.GetEnumDescription(deviceGuid));


         // 指定されたマシンのデバイスツリーの「ルート」から開始します。

         using (safeMachineHandle)
         using (var safeHandle = NativeMethods.SetupDiGetClassDevsEx(ref classGuid, IntPtr.Zero, IntPtr.Zero, NativeMethods.SetupDiGetClassDevsExFlags.Present | NativeMethods.SetupDiGetClassDevsExFlags.DeviceInterface, IntPtr.Zero, hostName, IntPtr.Zero))
         {
            NativeMethods.IsValidHandle(safeHandle, Marshal.GetLastWin32Error());

            uint memberInterfaceIndex = 0;
            var interfaceStructSize = (uint)Marshal.SizeOf<NativeMethods.SP_DEVICE_INTERFACE_DATA>();
            var dataStructSize = (uint)Marshal.SizeOf<NativeMethods.SP_DEVINFO_DATA>();


            // デバイスインターフェースの列挙を開始します。

            while (true)
            {
               var interfaceData = new NativeMethods.SP_DEVICE_INTERFACE_DATA { cbSize = interfaceStructSize };

               var success = NativeMethods.SetupDiEnumDeviceInterfaces(safeHandle, IntPtr.Zero, ref classGuid, memberInterfaceIndex++, ref interfaceData);

               lastError = Marshal.GetLastWin32Error();

               if (!success)
               {
                  if (lastError != Win32Errors.NO_ERROR && lastError != Win32Errors.ERROR_NO_MORE_ITEMS)
                  {
                     NativeError.ThrowException(lastError, hostName);
                  }

                  break;
               }


               // DeviceInfo インスタンスを作成します。

               var diData = new NativeMethods.SP_DEVINFO_DATA {cbSize = dataStructSize};

               var deviceInfo = new DeviceInfo(hostName) {DevicePath = GetDeviceInterfaceDetail(safeHandle, ref interfaceData, ref diData).DevicePath};


               if (getAllProperties)
               {
                  deviceInfo.InstanceId = GetDeviceInstanceId(safeMachineHandle, hostName, diData);

                  SetDeviceProperties(safeHandle, deviceInfo, diData);
               }

               else
               {
                  SetMinimalDeviceProperties(safeHandle, deviceInfo, diData);
               }


               yield return deviceInfo;
            }
         }
      }


      #region プライベートヘルパー

      [SecurityCritical]
      private static string GetDeviceInstanceId(SafeCmConnectMachineHandle safeMachineHandle, string hostName, NativeMethods.SP_DEVINFO_DATA diData)
      {
         uint ptrPrevious;

         var lastError = NativeMethods.CM_Get_Parent_Ex(out ptrPrevious, diData.DevInst, 0, safeMachineHandle);

         if (lastError != Win32Errors.CR_SUCCESS)
         {
            NativeError.ThrowException(lastError, hostName);
         }


         using var safeBuffer = new SafeGlobalMemoryBufferHandle(NativeMethods.DefaultNativeQueryBufferSize / 8);
         lastError = NativeMethods.CM_Get_Device_ID_Ex(diData.DevInst, safeBuffer, (uint) safeBuffer.Capacity, 0, safeMachineHandle);

         if (lastError != Win32Errors.CR_SUCCESS)
         {
            NativeError.ThrowException(lastError, hostName);
         }


         // デバイス InstanceID。例: "USB\VID_8087&PID_0A2B\5&2EDA7E1E&0&7", "SCSI\DISK&VEN_SANDISK&PROD_X400\4&288ED25&0&000200" など

         return safeBuffer.PtrToStringUni();
      }


      /// <summary>デバイスインターフェース詳細データ構造体を構築します。</summary>
      /// <returns>初期化された NativeMethods.SP_DEVICE_INTERFACE_DETAIL_DATA インスタンス。</returns>
      [SecurityCritical]
      private static NativeMethods.SP_DEVICE_INTERFACE_DETAIL_DATA GetDeviceInterfaceDetail(SafeHandle safeHandle, ref NativeMethods.SP_DEVICE_INTERFACE_DATA interfaceData, ref NativeMethods.SP_DEVINFO_DATA infoData)
      {
         var didd = new NativeMethods.SP_DEVICE_INTERFACE_DETAIL_DATA {cbSize = (uint) (IntPtr.Size == 4 ? 6 : 8)};

         var success = NativeMethods.SetupDiGetDeviceInterfaceDetail(safeHandle, ref interfaceData, ref didd, (uint) Marshal.SizeOf(didd), IntPtr.Zero, ref infoData);

         var lastError = Marshal.GetLastWin32Error();

         if (!success)
         {
            NativeError.ThrowException(lastError);
         }

         return didd;
      }


      [SecurityCritical]
      private static string GetDeviceRegistryProperty(SafeHandle safeHandle, NativeMethods.SP_DEVINFO_DATA infoData, NativeMethods.SetupDiGetDeviceRegistryPropertyEnum property)
      {
         var bufferSize = NativeMethods.DefaultNativeQueryBufferSize / 8; // 512

         while (true)
         {
            using var safeBuffer = new SafeGlobalMemoryBufferHandle(bufferSize);
            var success = NativeMethods.SetupDiGetDeviceRegistryProperty(safeHandle, ref infoData, property, IntPtr.Zero, safeBuffer, (uint) safeBuffer.Capacity, IntPtr.Zero);

            var lastError = Marshal.GetLastWin32Error();

            if (success)
            {
               var value = safeBuffer.PtrToStringUni();

               return !Utils.IsNullOrWhiteSpace(value) ? value.Trim() : null;
            }


            // MSDN: SetupDiGetDeviceRegistryProperty は、要求されたプロパティがデバイスに存在しない場合、
            // またはプロパティデータが無効な場合に ERROR_INVALID_DATA エラーコードを返します。

            if (lastError == Win32Errors.ERROR_INVALID_DATA)
            {
               return null;
            }


            bufferSize = GetDoubledBufferSizeOrThrowException(lastError, safeBuffer, bufferSize, property.ToString());
         }
      }



      [SecurityCritical]
      private static void SetDeviceProperties(SafeHandle safeHandle, DeviceInfo deviceInfo, NativeMethods.SP_DEVINFO_DATA infoData)
      {
         SetMinimalDeviceProperties(safeHandle, deviceInfo, infoData);


         deviceInfo.CompatibleIds = GetDeviceRegistryProperty(safeHandle, infoData, NativeMethods.SetupDiGetDeviceRegistryPropertyEnum.CompatibleIds);

         deviceInfo.Driver = GetDeviceRegistryProperty(safeHandle, infoData, NativeMethods.SetupDiGetDeviceRegistryPropertyEnum.Driver);

         deviceInfo.EnumeratorName = GetDeviceRegistryProperty(safeHandle, infoData, NativeMethods.SetupDiGetDeviceRegistryPropertyEnum.EnumeratorName);

         deviceInfo.HardwareId = GetDeviceRegistryProperty(safeHandle, infoData, NativeMethods.SetupDiGetDeviceRegistryPropertyEnum.HardwareId);

         deviceInfo.LocationInformation = GetDeviceRegistryProperty(safeHandle, infoData, NativeMethods.SetupDiGetDeviceRegistryPropertyEnum.LocationInformation);

         deviceInfo.LocationPaths = GetDeviceRegistryProperty(safeHandle, infoData, NativeMethods.SetupDiGetDeviceRegistryPropertyEnum.LocationPaths);

         deviceInfo.Manufacturer = GetDeviceRegistryProperty(safeHandle, infoData, NativeMethods.SetupDiGetDeviceRegistryPropertyEnum.Manufacturer);

         deviceInfo.Service = GetDeviceRegistryProperty(safeHandle, infoData, NativeMethods.SetupDiGetDeviceRegistryPropertyEnum.Service);
      }


      [SecurityCritical]
      private static void SetMinimalDeviceProperties(SafeHandle safeHandle, DeviceInfo deviceInfo, NativeMethods.SP_DEVINFO_DATA infoData)
      {
         deviceInfo.BaseContainerId = new Guid(GetDeviceRegistryProperty(safeHandle, infoData, NativeMethods.SetupDiGetDeviceRegistryPropertyEnum.BaseContainerId));

         deviceInfo.ClassGuid = new Guid(GetDeviceRegistryProperty(safeHandle, infoData, NativeMethods.SetupDiGetDeviceRegistryPropertyEnum.ClassGuid));

         deviceInfo.DeviceClass = GetDeviceRegistryProperty(safeHandle, infoData, NativeMethods.SetupDiGetDeviceRegistryPropertyEnum.Class);

         deviceInfo.DeviceDescription = GetDeviceRegistryProperty(safeHandle, infoData, NativeMethods.SetupDiGetDeviceRegistryPropertyEnum.DeviceDescription);

         deviceInfo.FriendlyName = GetDeviceRegistryProperty(safeHandle, infoData, NativeMethods.SetupDiGetDeviceRegistryPropertyEnum.FriendlyName);

         deviceInfo.PhysicalDeviceObjectName = GetDeviceRegistryProperty(safeHandle, infoData, NativeMethods.SetupDiGetDeviceRegistryPropertyEnum.PhysicalDeviceObjectName);
      }


      [SecurityCritical]
      internal static int GetDoubledBufferSizeOrThrowException(int lastError, SafeHandle safeBuffer, int bufferSize, string pathForException)
      {
         if (null != safeBuffer && !safeBuffer.IsClosed)
         {
            safeBuffer.Close();
         }


         switch ((uint) lastError)
         {
            case Win32Errors.ERROR_MORE_DATA:
            case Win32Errors.ERROR_INSUFFICIENT_BUFFER:
               bufferSize *= 2;
               break;


            default:
               NativeMethods.IsValidHandle(safeBuffer, lastError, string.Format(CultureInfo.InvariantCulture, "Buffer size: {0}. Path: {1}", bufferSize.ToString(CultureInfo.InvariantCulture), pathForException));
               break;
         }


         return bufferSize;
      }
      

      /// <summary>十分なメモリが割り当てられるまで、指定された入力で InvokeIoControl を繰り返し呼び出します。</summary>
      [SecurityCritical]
      private static void InvokeIoControlUnknownSize<T>(SafeFileHandle handle, uint controlCode, T input, uint increment = 128) where T : struct
      {
         var inputSize = (uint)Marshal.SizeOf(input);
         var outputLength = increment;

         var pInput = Marshal.AllocHGlobal((int)inputSize);
         try
         {
            Marshal.StructureToPtr(input, pInput, false);

            do
            {
               var pOutput = Marshal.AllocHGlobal((int)outputLength);
               try
               {
                  var success = NativeMethods.DeviceIoControlUnknownSize(handle, controlCode, pInput, inputSize, pOutput, outputLength, out var bytesReturned, IntPtr.Zero);

                  var lastError = Marshal.GetLastWin32Error();
                  if (!success)
                  {
                     switch ((uint)lastError)
                     {
                        case Win32Errors.ERROR_MORE_DATA:
                        case Win32Errors.ERROR_INSUFFICIENT_BUFFER:
                           outputLength += increment;
                           break;
                        default:
                           if (lastError != Win32Errors.ERROR_SUCCESS)
                           {
                              NativeError.ThrowException(lastError);
                           }
                           break;
                     }
                  }
                  else
                  {
                     break;
                  }
               }
               finally
               {
                  Marshal.FreeHGlobal(pOutput);
               }
            } while (true);
         }
         finally
         {
            Marshal.FreeHGlobal(pInput);
         }
      }

      #endregion // プライベートヘルパー


      #endregion // デバイスの列挙


      #region 圧縮

      /// <summary>[AlphaFS] ファイル単位およびディレクトリ単位の圧縮をサポートするボリューム上のファイルまたはディレクトリの NTFS 圧縮状態を設定します。</summary>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="isFolder"><paramref name="path"/> がファイルかディレクトリかを指定します。</param>
      /// <param name="path">圧縮または展開するフォルダーまたはファイルを記述するパス。</param>
      /// <param name="compress"><c>true</c> = 圧縮、<c>false</c> = 展開</param>
      /// <param name="pathFormat">パスパラメーターの形式を示します。</param>
      [SecurityCritical]
      internal static void ToggleCompressionCore(KernelTransaction transaction, bool isFolder, string path, bool compress, PathFormat pathFormat)
      {
         using var handle = File.CreateFileCore(transaction, isFolder, path, ExtendedFileAttributes.BackupSemantics, null, FileMode.Open, FileSystemRights.Modify, FileShare.None, true, false, pathFormat);
         InvokeIoControlUnknownSize(handle, NativeMethods.FSCTL_SET_COMPRESSION, compress ? 1 : 0);
      }

      #endregion // 圧縮


      #region リンク

      /// <summary>[AlphaFS] NTFS ディレクトリジャンクションを作成します（CMD コマンド "MKLINK /J" と同等）。</summary>
      internal static void CreateDirectoryJunction(SafeFileHandle safeHandle, string directoryPath)
      {
         var targetDirBytes = Encoding.Unicode.GetBytes(Path.NonInterpretedPathPrefix + Path.GetRegularPathCore(directoryPath, GetFullPathOptions.AddTrailingDirectorySeparator, false));

         var header = new NativeMethods.ReparseDataBufferHeader
         {
            ReparseTag = ReparsePointTag.MountPoint,
            ReparseDataLength = (ushort) (targetDirBytes.Length + 12)
         };

         var mountPoint = new NativeMethods.MountPointReparseBuffer
         {
            SubstituteNameOffset = 0,
            SubstituteNameLength = (ushort) targetDirBytes.Length,
            PrintNameOffset = (ushort) (targetDirBytes.Length + UnicodeEncoding.CharSize),
            PrintNameLength = 0
         };

         var reparseDataBuffer = new NativeMethods.REPARSE_DATA_BUFFER
         {
            ReparseTag = header.ReparseTag,
            ReparseDataLength = header.ReparseDataLength,

            SubstituteNameOffset = mountPoint.SubstituteNameOffset,
            SubstituteNameLength = mountPoint.SubstituteNameLength,
            PrintNameOffset = mountPoint.PrintNameOffset,
            PrintNameLength = mountPoint.PrintNameLength,

            PathBuffer = new byte[NativeMethods.MAXIMUM_REPARSE_DATA_BUFFER_SIZE - 16] // 16368
         };

         targetDirBytes.CopyTo(reparseDataBuffer.PathBuffer, 0);


         using var safeBuffer = new SafeGlobalMemoryBufferHandle(Marshal.SizeOf(reparseDataBuffer));
         safeBuffer.StructureToPtr(reparseDataBuffer, false);

         uint bytesReturned;
         var succes = NativeMethods.DeviceIoControl2(safeHandle, NativeMethods.FSCTL_SET_REPARSE_POINT, safeBuffer, (uint) (targetDirBytes.Length + 20), IntPtr.Zero, 0, out bytesReturned, IntPtr.Zero);

         var lastError = Marshal.GetLastWin32Error();
         if (!succes)
         {
            NativeError.ThrowException(lastError, directoryPath);
         }
      }


      /// <summary>[AlphaFS] NTFS ディレクトリジャンクションを削除します。</summary>
      internal static void DeleteDirectoryJunction(SafeFileHandle safeHandle)
      {
         var reparseDataBuffer = new NativeMethods.REPARSE_DATA_BUFFER
         {
            ReparseTag = ReparsePointTag.MountPoint,
            ReparseDataLength = 0,
            PathBuffer = new byte[NativeMethods.MAXIMUM_REPARSE_DATA_BUFFER_SIZE - 16] // 16368
         };


         using var safeBuffer = new SafeGlobalMemoryBufferHandle(Marshal.SizeOf(reparseDataBuffer));
         safeBuffer.StructureToPtr(reparseDataBuffer, false);

         uint bytesReturned;
         var success = NativeMethods.DeviceIoControl2(safeHandle, NativeMethods.FSCTL_DELETE_REPARSE_POINT, safeBuffer, NativeMethods.REPARSE_DATA_BUFFER_HEADER_SIZE, IntPtr.Zero, 0, out bytesReturned, IntPtr.Zero);

         var lastError = Marshal.GetLastWin32Error();
         if (!success)
         {
            NativeError.ThrowException(lastError);
         }
      }


      /// <summary>[AlphaFS] NTFS ファイルシステム上のマウントポイントまたはシンボリックリンクのターゲットに関する情報を取得します。</summary>
      /// <exception cref="NotAReparsePointException"/>
      /// <exception cref="UnrecognizedReparsePointException"/>
      [SecurityCritical]
      internal static LinkTargetInfo GetLinkTargetInfo(SafeFileHandle safeHandle, string reparsePath)
      {
         using var safeBuffer = GetLinkTargetData(safeHandle, reparsePath);
         var header = safeBuffer.PtrToStructure<NativeMethods.ReparseDataBufferHeader>(0);

         var marshalReparseBuffer = (int) Marshal.OffsetOf<NativeMethods.ReparseDataBufferHeader>("data");

         var dataOffset = (int) (marshalReparseBuffer + (header.ReparseTag == ReparsePointTag.MountPoint
            ? Marshal.OffsetOf<NativeMethods.MountPointReparseBuffer>("data")
            : Marshal.OffsetOf<NativeMethods.SymbolicLinkReparseBuffer>("data")).ToInt64());

         var dataBuffer = new byte[NativeMethods.MAXIMUM_REPARSE_DATA_BUFFER_SIZE - dataOffset];


         switch (header.ReparseTag)
         {
            // MountPoint はジャンクションまたはマウントされたドライブです（マウントされたドライブは "\??\Volume" で始まります）。

            case ReparsePointTag.MountPoint:
               var mountPoint = safeBuffer.PtrToStructure<NativeMethods.MountPointReparseBuffer>(marshalReparseBuffer);

               safeBuffer.CopyTo(dataOffset, dataBuffer);

               return new LinkTargetInfo(
                  Encoding.Unicode.GetString(dataBuffer, mountPoint.SubstituteNameOffset, mountPoint.SubstituteNameLength),
                  Encoding.Unicode.GetString(dataBuffer, mountPoint.PrintNameOffset, mountPoint.PrintNameLength));


            case ReparsePointTag.SymLink:
               var symLink = safeBuffer.PtrToStructure<NativeMethods.SymbolicLinkReparseBuffer>(marshalReparseBuffer);

               safeBuffer.CopyTo(dataOffset, dataBuffer);

               return new SymbolicLinkTargetInfo(
                  Encoding.Unicode.GetString(dataBuffer, symLink.SubstituteNameOffset, symLink.SubstituteNameLength),
                  Encoding.Unicode.GetString(dataBuffer, symLink.PrintNameOffset, symLink.PrintNameLength), symLink.Flags);


            default:
               throw new UnrecognizedReparsePointException(reparsePath);
         }
      }


      /// <summary>[AlphaFS] NTFS ファイルシステム上のマウントポイントまたはシンボリックリンクのターゲットに関する情報を取得します。</summary>
      /// <exception cref="NotAReparsePointException"/>
      /// <exception cref="UnrecognizedReparsePointException"/>
      [SuppressMessage("Microsoft.Usage", "CA2202:Do not dispose objects multiple times")]
      [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "Disposing is controlled.")]
      [SecurityCritical]
      private static SafeGlobalMemoryBufferHandle GetLinkTargetData(SafeFileHandle safeHandle, string reparsePath)
      {
         var safeBuffer = new SafeGlobalMemoryBufferHandle(NativeMethods.MAXIMUM_REPARSE_DATA_BUFFER_SIZE);

         while (true)
         {
            var success = NativeMethods.DeviceIoControl(safeHandle, NativeMethods.FSCTL_GET_REPARSE_POINT, IntPtr.Zero, 0, safeBuffer, (uint) safeBuffer.Capacity, out var bytesReturned, IntPtr.Zero);

            var lastError = Marshal.GetLastWin32Error();
            if (!success)
            {
               switch ((uint) lastError)
               {
                  case Win32Errors.ERROR_MORE_DATA:
                  case Win32Errors.ERROR_INSUFFICIENT_BUFFER:

                     // 最大サイズを既に使用しているため、通常は発生しません。

                     if (safeBuffer.Capacity < bytesReturned)
                     {
                        safeBuffer.Close();
                     }
                     break;


                  default:
                     if (lastError != Win32Errors.ERROR_SUCCESS)
                     {
                        NativeError.ThrowException(lastError, reparsePath);
                     }
                     break;
               }
            }

            else
            {
               break;
            }
         }


         return safeBuffer;
      }

      #endregion // リンク
   }
}
