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
using System.Security;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class File
   {
      /// <summary>[AlphaFS] 指定されたファイルのサイズを取得します。</summary>
      /// <returns>ファイルサイズ(バイト単位)。</returns>
      /// <param name="path">ファイルへのパス。</param>
      [SecurityCritical]
      public static long GetSize(string path)
      {
         return GetSizeCore(null, null, path, false, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] 指定されたファイルのサイズを取得します。</summary>
      /// <returns>ファイルサイズ(バイト単位)。</returns>
      /// <param name="path">ファイルへのパス。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static long GetSize(string path, PathFormat pathFormat)
      {
         return GetSizeCore(null, null, path, false, pathFormat);
      }


      /// <summary>[AlphaFS] 指定されたファイルのサイズを取得します。</summary>
      /// <returns>最初のストリームまたはすべてのストリームのファイルサイズ(バイト単位)。</returns>
      /// <param name="path">ファイルへのパス。</param>
      /// <param name="sizeOfAllStreams">すべての代替データストリームのサイズを取得する場合は<c>true</c>、最初のストリームのサイズを取得する場合は<c>false</c>。</param>
      [SecurityCritical]
      public static long GetSize(string path, bool sizeOfAllStreams)
      {
         return GetSizeCore(null, null, path, sizeOfAllStreams, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] 指定されたファイルのサイズを取得します。</summary>
      /// <returns>最初のストリームまたはすべてのストリームのファイルサイズ(バイト単位)。</returns>
      /// <param name="path">ファイルへのパス。</param>
      /// <param name="sizeOfAllStreams">すべての代替データストリームのサイズを取得する場合は<c>true</c>、最初のストリームのサイズを取得する場合は<c>false</c>。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static long GetSize(string path, bool sizeOfAllStreams, PathFormat pathFormat)
      {
         return GetSizeCore(null, null, path, sizeOfAllStreams, pathFormat);
      }


      /// <summary>[AlphaFS] 指定されたファイルのサイズを取得します。</summary>
      /// <returns>ファイルサイズ(バイト単位)。</returns>
      /// <param name="handle">ファイルへの<see cref="SafeFileHandle"/>。</param>
      [SecurityCritical]
      public static long GetSize(SafeFileHandle handle)
      {
         return GetSizeCore(handle, null, null, false, PathFormat.LongFullPath);
      }
   }
}
