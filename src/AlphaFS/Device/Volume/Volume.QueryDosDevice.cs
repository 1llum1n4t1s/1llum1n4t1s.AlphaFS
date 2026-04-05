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
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class Volume
   {
      /// <summary>[AlphaFS] 既存のすべての MS-DOS デバイス名のソートされたリストを取得します。</summary>
      /// <returns>既存のすべての MS-DOS デバイス名のソートされた <see cref="IEnumerable{String}"/> リスト。</returns>
      [SecurityCritical]
      public static IEnumerable<string> QueryAllDosDevices()
      {
         return QueryDosDeviceCore(null, true);
      }


      /// <summary>[AlphaFS] 特定の MS-DOS デバイス名の現在のマッピングを取得します。</summary>
      /// <returns>特定の MS-DOS デバイス名の現在のマッピング。</returns>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="FileNotFoundException"/>
      /// <param name="deviceName">クエリのターゲットを指定する MS-DOS デバイス名文字列。例: "C:"、"D:" または "\\?\Volume{GUID}"。</param>
      [SecurityCritical]
      public static string QueryDosDevice(string deviceName)
      {
         if (Utils.IsNullOrWhiteSpace(deviceName))
         {
            throw new ArgumentNullException("deviceName");
         }


         var devName = QueryDosDeviceCore(deviceName, false).ToArray()[0];

         return !Utils.IsNullOrWhiteSpace(devName) ? devName : null;
      }




      /// <summary>[AlphaFS] 特定の MS-DOS デバイス名の現在のマッピングを取得します。この関数は既存のすべての MS-DOS デバイス名のリストも取得できます。</summary>
      /// <returns>既存のすべての MS-DOS デバイス名のソートされた <see cref="IEnumerable{String}"/> リスト。</returns>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="FileNotFoundException"/>
      /// <param name="deviceName">クエリのターゲットを指定する MS-DOS デバイス名文字列。例: "C:"、"D:" または "\\?\Volume{GUID}"。</param>
      /// <param name="sort"><c>true</c> は MS-DOS デバイス名のリストをソートします。</param>
      [SecurityCritical]
      internal static IEnumerable<string> QueryDosDeviceCore(string deviceName, bool sort)
      {
         // deviceName は null が許可されています: 既存のすべての MS-DOS デバイス名のリストを取得します。
         // deviceName には末尾のバックスラッシュを含めることはできません。

         if (!Utils.IsNullOrWhiteSpace(deviceName))
         {
            if (deviceName.StartsWith(Path.GlobalRootPrefix, StringComparison.OrdinalIgnoreCase))
            {
               yield return deviceName.Substring(Path.GlobalRootPrefix.Length);

               yield break;
            }

            
            if (deviceName.StartsWith(Path.VolumePrefix, StringComparison.OrdinalIgnoreCase))

            {
               deviceName = deviceName.Substring(Path.LongPathPrefix.Length);
            }


            deviceName = Path.RemoveTrailingDirectorySeparator(deviceName);
         }




         uint returnedBufferSize = 0;

         var bufferSize = (uint) (sort ? NativeMethods.DefaultNativeQueryBufferSize : 64);

         var sortedList = new List<string>(sort ? 256 : 0);


         using (new NativeMethods.ChangeErrorMode(NativeMethods.ErrorMode.FailCriticalErrors))
            while (returnedBufferSize == 0)
            {
               var cBuffer = new char[bufferSize];

               returnedBufferSize = NativeMethods.QueryDosDevice(deviceName, cBuffer, bufferSize);

               var lastError = Marshal.GetLastWin32Error();

               if (returnedBufferSize == 0)
               {
                  switch ((uint) lastError)
                  {
                     case Win32Errors.ERROR_MORE_DATA:
                     case Win32Errors.ERROR_INSUFFICIENT_BUFFER:
                        bufferSize *= 2;
                        continue;

                     default:
                        NativeError.ThrowException(lastError, deviceName);
                        break;
                  }
               }


               var buffer = new StringBuilder((int) returnedBufferSize);


               for (var i = 0; i < returnedBufferSize; i++)
               {
                  if (cBuffer[i] != Path.StringTerminatorChar)

                  {
                     buffer.Append(cBuffer[i]);
                  }


                  else if (buffer.Length > 0)
                  {
                     var assembledPath = buffer.ToString();

                     assembledPath = !Utils.IsNullOrWhiteSpace(assembledPath) ? assembledPath : null;


                     if (sort)
                     {
                        sortedList.Add(assembledPath);
                     }

                     else
                     {
                        yield return assembledPath;
                     }


                     buffer.Length = 0;
                  }
               }
            }


         if (sort)
         {
            foreach (var devName in sortedList.OrderBy(devName => devName))

               yield return devName;
         }
      }
   }
}
