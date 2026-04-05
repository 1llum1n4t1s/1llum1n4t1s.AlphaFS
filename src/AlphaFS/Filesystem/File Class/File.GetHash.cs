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
using System.Security;
using Alphaleonis.Win32.Security;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class File
   {
      #region Obsolete

      /// <summary>[AlphaFS] 指定された<paramref name="fileFullPath"/>のハッシュ/チェックサムを計算します。</summary>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="fileFullPath">ファイルへのパス。</param>
      /// <param name="hashType"><see cref="HashType"/>値の1つ。</param>
      /// <returns>ハッシュ値。</returns>
      [Obsolete("Use GetHashTransacted method.")]
      [SecurityCritical]
      public static string GetHash(KernelTransaction transaction, string fileFullPath, HashType hashType)
      {
         return GetHashCore(transaction, fileFullPath, hashType, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] 指定された<paramref name="fileFullPath"/>のハッシュ/チェックサムを計算します。</summary>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="fileFullPath">ファイルへのパス。</param>
      /// <param name="hashType"><see cref="HashType"/>値の1つ。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      /// <returns>ハッシュ値。</returns>
      [Obsolete("Use GetHashTransacted method.")]
      [SecurityCritical]
      public static string GetHash(KernelTransaction transaction, string fileFullPath, HashType hashType, PathFormat pathFormat)
      {
         return GetHashCore(transaction, fileFullPath, hashType, pathFormat);
      }


      #endregion // Obsolete


      /// <summary>[AlphaFS] 指定された<paramref name="fileFullPath"/>のハッシュ/チェックサムを計算します。</summary>
      /// <param name="fileFullPath">ファイルへのパス。</param>
      /// <param name="hashType"><see cref="HashType"/>値の1つ。</param>
      /// <returns>ハッシュ値。</returns>
      [SecurityCritical]
      public static string GetHash(string fileFullPath, HashType hashType)
      {
         return GetHashCore(null, fileFullPath, hashType, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] 指定された<paramref name="fileFullPath"/>のハッシュ/チェックサムを計算します。</summary>
      /// <param name="fileFullPath">ファイルへのパス。</param>
      /// <param name="hashType"><see cref="HashType"/>値の1つ。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      /// <returns>ハッシュ値。</returns>
      [SecurityCritical]
      public static string GetHash(string fileFullPath, HashType hashType, PathFormat pathFormat)
      {
         return GetHashCore(null, fileFullPath, hashType, pathFormat);
      }
   }
}
