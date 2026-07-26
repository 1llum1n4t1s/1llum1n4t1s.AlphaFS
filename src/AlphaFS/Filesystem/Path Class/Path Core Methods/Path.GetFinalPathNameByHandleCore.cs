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

using Microsoft.Win32.SafeHandles;
using System;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class Path
   {
      /// <summary>指定されたファイルの最終パスを <see cref="FinalPathFormats"/> 形式で取得します。</summary>
      /// <returns>文字列としての最終パス。</returns>
      /// <remarks>
      ///   最終パスとは、パスが完全に解決された際に返されるパスです。例えば、"D:\yourdir" を指すシンボリックリンク "C:\tmp\mydir" の場合、
      ///   最終パスは "D:\yourdir" となります。この関数が返す文字列は
      ///   <see cref="LongPathPrefix"/> 構文を使用します。
      /// </remarks>
      /// <param name="handle"><see cref="SafeFileHandle"/> インスタンスへのハンドル。</param>
      /// <param name="finalPath"><see cref="FinalPathFormats"/> 形式の最終パス。</param>
      [SecurityCritical]
      internal static string GetFinalPathNameByHandleCore(SafeFileHandle handle, FinalPathFormats finalPath)
      {
         NativeMethods.IsValidHandle(handle);

         var buffer = new StringBuilder(NativeMethods.MaxPathUnicode);

         using (new NativeMethods.ChangeErrorMode(NativeMethods.ErrorMode.FailCriticalErrors))
         {
            if (NativeMethods.IsAtLeastWindowsVista)
            {
               // MSDN: GetFinalPathNameByHandle(): If the function fails for any other reason, the return value is zero.

               var returnValue = NativeMethods.GetFinalPathNameByHandle(handle, buffer, (uint) buffer.Capacity, finalPath);

               if (returnValue == Win32Errors.ERROR_SUCCESS)
               {
                  NativeError.ThrowException(Marshal.GetLastWin32Error());
               }


               return buffer.ToString();
            }
         }

         
         // 古いオペレーティングシステム

         // ファイルハンドルからファイル名を取得する
         // http://msdn.microsoft.com/en-us/library/aa366789%28VS.85%29.aspx

         // 不明な "File" 型オブジェクトの hFile ハンドルのサイズを確認するために GetFileSizeEx を使用する場合は注意が必要。
         // これはファイルハンドルからファイル名を返すことに関連している。ハンドルが名前付きパイプの場合、スレッドがハングする可能性がある。
         // チェック対象: FileTypes.DiskFile

         // 0バイトのファイルはマップできない。
         if (NativeMethods.GetFileSizeEx(handle, out var fileSizeHi) && fileSizeHi == 0)
         {
            return string.Empty;
         }


         // PAGE_READONLY
         // Allows views to be mapped for read-only or copy-on-write access. An attempt to write to a specific region results in an access violation.
         // The file handle that the hFile parameter specifies must be created with the GENERIC_READ access right.
         // PageReadOnly = 0x02,
         using (var handle2 = NativeMethods.CreateFileMapping(handle, null, 2, 0, 1, null))
         {
            NativeMethods.IsValidHandle(handle2, Marshal.GetLastWin32Error());

            // FILE_MAP_READ
            // Read = 4
            using (var pMem = NativeMethods.MapViewOfFile(handle2, 4, 0, 0, (UIntPtr)1))
            {
               if (NativeMethods.IsValidHandle(pMem, Marshal.GetLastWin32Error()))
               {
                  // Process インスタンスを保持せずに .Handle だけ渡すと、ネイティブ呼び出しの前に Process が
                  // 回収されてハンドルが閉じられ得る。using で呼び出し完了まで生存させる。
                  using (var process = Process.GetCurrentProcess())
                  {
                     NativeMethods.GetMappedFileName(process.Handle, pMem, buffer, (uint) buffer.Capacity);
                  }
               }
            }
         }


         // GetMappedFileName() のデフォルト出力: "\Device\HarddiskVolumeX\path\filename.ext"
         var dosDevice = buffer.Length > 0 ? buffer.ToString() : string.Empty;


         // 出力形式を選択する。
         switch (finalPath)
         {
            // As-is: "\Device\HarddiskVolumeX\path\filename.ext"
            case FinalPathFormats.VolumeNameNT:
               return dosDevice;


            // To: "\path\filename.ext"
            case FinalPathFormats.VolumeNameNone:
               return DosDeviceToDosPath(dosDevice, string.Empty);


            // To: "\\?\Volume{GUID}\path\filename.ext"
            case FinalPathFormats.VolumeNameGuid:
               var dosPath = DosDeviceToDosPath(dosDevice, null);

               if (!Utils.IsNullOrWhiteSpace(dosPath))
               {
                  var driveLetter = RemoveTrailingDirectorySeparator(GetPathRoot(dosPath, false));
                  var file = GetFileName(dosPath, true);


                  if (!Utils.IsNullOrWhiteSpace(file))
                  { 
                     foreach (var drive in Directory.EnumerateLogicalDrivesCore(false, false)
                        
                        .Select(drv => drv.Name).Where(drv => driveLetter.Equals(RemoveTrailingDirectorySeparator(drv), StringComparison.OrdinalIgnoreCase)))

                        return CombineCore(false, Volume.GetUniqueVolumeNameForPath(drive), GetSuffixedDirectoryNameWithoutRootCore(null, dosPath, PathFormat.FullPath), file);
                  }
               }

               break;
         }


         // To: "\\?\C:\path\filename.ext"
         return !Utils.IsNullOrWhiteSpace(dosDevice) ? LongPathPrefix + DosDeviceToDosPath(dosDevice, null) : string.Empty;
      }


      /// <summary>DosDevicePath、ボリュームGUIDを変換します。例: "\Device\HarddiskVolumeX\path\filename.ext" は "\path\filename.ext" または "\\?\Volume{GUID}\path\filename.ext" に変換できます。</summary>
      /// <returns>変換されたDOSパス。</returns>
      /// <param name="dosDevice">DosDevicePath。例: \Device\HarddiskVolumeX\path\filename.ext。</param>
      /// <param name="deviceReplacement">代替パス/デバイステキスト。通常は <c>string.Empty</c> または <c>null</c>。</param>
      [SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes")]
      [SecurityCritical]
      private static string DosDeviceToDosPath(string dosDevice, string deviceReplacement)
      {
         if (Utils.IsNullOrWhiteSpace(dosDevice))
         {
            return string.Empty;
         }


         foreach (var drive in Directory.EnumerateLogicalDrivesCore(false, false).Select(drv => drv.Name))
         {
            try
            {
               var path = RemoveTrailingDirectorySeparator(drive);

               foreach (var devNt in Volume.QueryAllDosDevices().Where(device => device.StartsWith(path, StringComparison.OrdinalIgnoreCase)))

                  return dosDevice.ReplaceIgnoreCase(devNt, deviceReplacement ?? path);
            }
            catch
            {
            }
         }

         return string.Empty;
      }
   }
}
