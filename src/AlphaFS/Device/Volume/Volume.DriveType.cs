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

using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Security;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class Volume
   {
      /// <summary>[AlphaFS] 現在のディレクトリのルートに基づいてディスクの <see cref="DriveType"/> を判定します。</summary>
      /// <returns><see cref="DriveType"/> 列挙値。</returns>
      [SuppressMessage("Microsoft.Design", "CA1024:UsePropertiesWhereAppropriate")]
      [SecurityCritical]
      public static DriveType GetCurrentDriveType()
      {
         return GetDriveType(null);
      }


      /// <summary>[AlphaFS] ディスクの <see cref="DriveType"/> を判定します。</summary>
      /// <param name="drivePath">ドライブへのパス。例: "C:\"、"\\server\share"、または "\\?\Volume{c0580d5e-2ad6-11dc-9924-806e6f6e6963}\"</param>
      /// <returns><see cref="DriveType"/> 列挙値。</returns>
      [SecurityCritical]
      public static DriveType GetDriveType(string drivePath)
      {
         // drivePath は null であることが許可されています。

         drivePath = Path.AddTrailingDirectorySeparator(drivePath, false);


         // ChangeErrorMode は Win32 SetThreadErrorMode() メソッド用で、ポップアップの抑制に使用されます。

         using (new NativeMethods.ChangeErrorMode(NativeMethods.ErrorMode.FailCriticalErrors))

            return NativeMethods.GetDriveType(drivePath);
      }
   }
}
