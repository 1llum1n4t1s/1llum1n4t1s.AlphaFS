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

using System.Security;
using Microsoft.Win32.SafeHandles;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class Volume
   {
      /// <summary>[AlphaFS] 指定されたルートファイルまたはディレクトリストリームに関連付けられたファイルシステムとボリュームに関する情報を取得します。</summary>
      /// <param name="volumePath">ルートディレクトリを含むパス。</param>
      /// <returns>指定されたルートディレクトリに関連付けられたボリュームを記述する <see cref="VolumeInfo"/> インスタンス。</returns>
      [SecurityCritical]
      public static VolumeInfo GetVolumeInfo(string volumePath)
      {
         return new VolumeInfo(volumePath, true, false);
      }


      /// <summary>[AlphaFS] 指定されたルートファイルまたはディレクトリストリームに関連付けられたファイルシステムとボリュームに関する情報を取得します。</summary>
      /// <param name="volumeHandle"><see cref="SafeFileHandle"/> ハンドルのインスタンス。</param>
      /// <returns>指定されたルートディレクトリに関連付けられたボリュームを記述する <see cref="VolumeInfo"/> インスタンス。</returns>
      [SecurityCritical]
      public static VolumeInfo GetVolumeInfo(SafeFileHandle volumeHandle)
      {
         return new VolumeInfo(volumeHandle, true, true);
      }
   }
}
