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
using System.Runtime.InteropServices;
using System.Security;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class File
   {
      /// <summary>[AlphaFS] ファイルの一意識別子を取得します。識別子は64ビットのボリュームシリアル番号と128ビットのファイルシステムエントリ識別子で構成されます。</summary>
      /// <returns>要求された情報を含む<see cref="FileIdInfo"/>インスタンス。</returns>
      /// <remarks>ファイルIDは時間の経過とともに一意であることが保証さ���ていません。ファイルシステムはそれらを再利用できるためです。場合によっては、ファイルのファイルIDが時間の経過とともに変更されることがあります。</remarks>
      /// <param name="path">ファイルへのパス。</param>
      [SecurityCritical]
      public static FileIdInfo GetFileIdInfo(string path)
      {
         return GetFileIdInfoCore(null, false, path, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] ファイルの一意識別子を取得します。識別子は64ビットのボリュームシリアル番号と128ビットのファイルシステムエントリ識別子で構成されます。</summary>
      /// <returns>要求された情報を含む<see cref="FileIdInfo"/>インスタンス。</returns>
      /// <remarks>ファイルIDは時間の経過とともに一意であることが保証さ���ていません。ファイルシステムはそれらを再利用できるためです。場合によっては、ファイルのファイルIDが時間の経過とともに変更されることがあります。</remarks>
      /// <param name="path">ファイルへのパス。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static FileIdInfo GetFileIdInfo(string path, PathFormat pathFormat)
      {
         return GetFileIdInfoCore(null, false, path, pathFormat);
      }


      /// <summary>[AlphaFS] Retrieves file information for the specified <see cref="SafeFileHandle"/>.</summary>
      /// <returns>要求された情報を含む<see cref="FileIdInfo"/>インスタンス。</returns>
      /// <remarks>ファイルIDは時間の経過とともに一意であることが保証さ���ていません。ファイルシステムはそれらを再利用できるためです。場合によっては、ファイルのファイルIDが時間の経過とともに変更されることがあります。</remarks>
      /// <param name="handle">A <see cref="SafeFileHandle"/> connected to the open file or directory from which to retrieve the information.</param>
      [SecurityCritical]
      public static FileIdInfo GetFileIdInfo(SafeFileHandle handle)
      {

         var success = NativeMethods.GetFileInformationByHandle(handle, out var info);

         var lastError = Marshal.GetLastWin32Error();
         if (!success)
         {
            NativeError.ThrowException(lastError);
         }

         return new FileIdInfo(info);
      }
   }
}
