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
   public static partial class Directory
   {
      /// <summary>[AlphaFS] Gets the <see cref="FileSystemEntryInfo"/> of the directory on the path.</summary>
      /// <returns>The <see cref="FileSystemEntryInfo"/> instance of the directory.</returns>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">ディレクトリへのパス。</param>
      [SecurityCritical]
      public static FileSystemEntryInfo GetFileSystemEntryInfoTransacted(KernelTransaction transaction, string path)
      {
         return File.GetFileSystemEntryInfoCore(transaction, true, path, false, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] Gets the <see cref="FileSystemEntryInfo"/> of the directory on the path.</summary>
      /// <returns>The <see cref="FileSystemEntryInfo"/> instance of the directory.</returns>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">ディレクトリへのパス。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static FileSystemEntryInfo GetFileSystemEntryInfoTransacted(KernelTransaction transaction, string path, PathFormat pathFormat)
      {
         return File.GetFileSystemEntryInfoCore(transaction, true, path, false, pathFormat);
      }


      /// <summary>[AlphaFS] Gets the <see cref="FileSystemEntryInfo"/> of the directory on the path.</summary>
      /// <returns>The <see cref="FileSystemEntryInfo"/> instance of the directory.</returns>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">ディレクトリへのパス。</param>
      /// <param name="continueOnException">
      ///    <para><c>true</c> suppress any Exception that might be thrown as a result from a failure,</para>
      ///    <para>ACLで保護されたディレクトリやアクセス不可なリパースポイントなど。</para>
      /// </param>
      [SecurityCritical]
      public static FileSystemEntryInfo GetFileSystemEntryInfoTransacted(KernelTransaction transaction, string path, bool continueOnException)
      {
         return File.GetFileSystemEntryInfoCore(transaction, true, path, continueOnException, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] Gets the <see cref="FileSystemEntryInfo"/> of the directory on the path.</summary>
      /// <returns>The <see cref="FileSystemEntryInfo"/> instance of the directory.</returns>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">ディレクトリへのパス。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      /// <param name="continueOnException">
      ///    <para><c>true</c> suppress any Exception that might be thrown as a result from a failure,</para>
      ///    <para>ACLで保護されたディレクトリやアクセス不可なリパースポイントなど。</para>
      /// </param>
      [SecurityCritical]
      public static FileSystemEntryInfo GetFileSystemEntryInfoTransacted(KernelTransaction transaction, string path, bool continueOnException, PathFormat pathFormat)
      {
         return File.GetFileSystemEntryInfoCore(transaction, true, path, continueOnException, pathFormat);
      }
   }
}
