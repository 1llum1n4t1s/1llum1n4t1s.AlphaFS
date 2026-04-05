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
   public static partial class Path
   {
      /// <summary>[AlphaFS] 指定されたファイルの最終パスを <see cref="FinalPathFormats"/> 形式で取得します。</summary>
      /// <returns>文字列としての最終パス。</returns>
      /// <remarks>
      ///   最終パスとは、パスが完全に解決された際に返されるパスです。例えば、"D:\yourdir" を指すシンボリックリンク "C:\tmp\mydir" の場合、
      ///   最終パスは "D:\yourdir" となります。
      /// </remarks>
      /// <param name="handle"><see cref="SafeFileHandle"/> インスタンスへのハンドル。</param>
      [SecurityCritical]
      public static string GetFinalPathNameByHandle(SafeFileHandle handle)
      {
         return GetFinalPathNameByHandleCore(handle, FinalPathFormats.None);
      }


      /// <summary>[AlphaFS] 指定されたファイルの最終パスを <see cref="FinalPathFormats"/> 形式で取得します。</summary>
      /// <returns>文字列としての最終パス。</returns>
      /// <remarks>
      ///   最終パスとは、パスが完全に解決された際に返されるパスです。例えば、"D:\yourdir" を指すシンボリックリンク "C:\tmp\mydir" の場合、
      ///   最終パスは "D:\yourdir" となります。
      /// </remarks>
      /// <param name="handle"><see cref="SafeFileHandle"/> インスタンスへのハンドル。</param>
      /// <param name="finalPath"><see cref="FinalPathFormats"/> 形式の最終パス。</param>
      [SecurityCritical]
      public static string GetFinalPathNameByHandle(SafeFileHandle handle, FinalPathFormats finalPath)
      {
         return GetFinalPathNameByHandleCore(handle, finalPath);
      }
   }
}
