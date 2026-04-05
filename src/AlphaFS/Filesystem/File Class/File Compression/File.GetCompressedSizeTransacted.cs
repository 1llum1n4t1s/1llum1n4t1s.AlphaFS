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
   public static partial class File
   {
      /// <summary>
      ///   [AlphaFS] トランザクションの一部として、指定されたファイルの格納に使用されているディスクストレージの実際のバイト数を取得します。
      ///   圧縮をサポートするボリューム上にファイルがあり、そのファイルが圧縮されている場合、取得される値は指定されたファイルの圧縮サイズです。
      ///   スパースファイルをサポートするボリューム上にファイルがあり、そのファイルがスパースファイルの場合、取得される値は指定されたファイルのスパースサイズです。
      /// </summary>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path"><para>ファイルの名前。</para></param>
      /// <returns>指定されたファイルの格納に使用されているディスクストレージの実際のバイト数。</returns>
      [SecurityCritical]
      public static long GetCompressedSizeTransacted(KernelTransaction transaction, string path)
      {
         return GetCompressedSizeCore(transaction, path, PathFormat.RelativePath);
      }


      /// <summary>
      ///   [AlphaFS] トランザクションの一部として、指定されたファイルの格納に使用されているディスクストレージの実際のバイト数を取得します。
      ///   圧縮をサポートするボリューム上にファイルがあり、そのファイルが圧縮されている場合、取得される値は指定されたファイルの圧縮サイズです。
      ///   スパースファイルをサポートするボリューム上にファイルがあり、そのファイルがスパースファイルの場合、取得される値は指定されたファイルのスパースサイズです。
      /// </summary>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path"><para>ファイルの名前。</para></param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      /// <returns>指定されたファイルの格納に使用されているディスクストレージの実際のバイト数。</returns>
      [SecurityCritical]
      public static long GetCompressedSizeTransacted(KernelTransaction transaction, string path, PathFormat pathFormat)
      {
         return GetCompressedSizeCore(transaction, path, pathFormat);
      }
   }
}
