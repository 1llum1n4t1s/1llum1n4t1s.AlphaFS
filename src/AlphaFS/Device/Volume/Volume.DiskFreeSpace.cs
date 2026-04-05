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

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class Volume
   {
      /// <summary>[AlphaFS]
      ///   ディスクボリュームで利用可能な領域の量に関する情報を取得します。これには、総領域量、総空き領域量、
      ///   および呼び出しスレッドに関連付けられたユーザーが利用可能な総空き領域量が含まれます。
      /// </summary>
      /// <remarks>呼び出し元アプリケーションには、このディレクトリに対する FILE_LIST_DIRECTORY アクセス権が必要です。</remarks>
      /// <param name="drivePath">
      ///   ドライブへのパス。例: "C:\"、"\\server\share"、または "\\?\Volume{c0580d5e-2ad6-11dc-9924-806e6f6e6963}\"。
      /// </param>
      /// <returns><see ref="Alphaleonis.Win32.Filesystem.DiskSpaceInfo"/> クラスインスタンス。</returns>
      [SecurityCritical]
      public static DiskSpaceInfo GetDiskFreeSpace(string drivePath)
      {
         return new DiskSpaceInfo(drivePath, null, true, true);
      }


      /// <summary>[AlphaFS]
      ///   ディスクボリュームで利用可能な領域の量に関する情報を取得します。これには、総領域量、総空き領域量、
      ///   および呼び出しスレッドに関連付けられたユーザーが利用可能な総空き領域量が含まれます。
      /// </summary>
      /// <remarks>呼び出し元アプリケーションには、このディレクトリに対する FILE_LIST_DIRECTORY アクセス権が必要です。</remarks>
      /// <param name="drivePath">
      ///   ドライブへのパス。例: "C:\"、"\\server\share"、または "\\?\Volume{c0580d5e-2ad6-11dc-9924-806e6f6e6963}\"。
      /// </param>
      /// <param name="spaceInfoType">
      ///   <c>null</c> はサイズ情報とディスククラスター情報の両方を取得します。<c>true</c> はディスククラスター情報のみを取得、
      ///   <c>false</c> はサイズ情報のみを取得します。
      /// </param>
      /// <returns><see ref="Alphaleonis.Win32.Filesystem.DiskSpaceInfo"/> クラスインスタンス。</returns>
      [SecurityCritical]
      public static DiskSpaceInfo GetDiskFreeSpace(string drivePath, bool? spaceInfoType)
      {
         return new DiskSpaceInfo(drivePath, spaceInfoType, true, true);
      }
   }
}
