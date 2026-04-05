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

using System.IO;
using System.Security;

namespace Alphaleonis.Win32.Filesystem
{
   public sealed partial class DirectoryInfo
   {
      /// <summary>[AlphaFS] <see cref="DirectoryInfo"/> インスタンスから空のサブディレクトリを削除します。</summary>
      [SecurityCritical]
      public void DeleteEmptySubdirectories()
      {
         Directory.DeleteEmptySubdirectoriesCore(EntryInfo, Transaction, null, false, false, PathFormat.LongFullPath);
      }


      /// <summary>[AlphaFS] <see cref="DirectoryInfo"/> インスタンスから空のサブディレクトリを削除します。</summary>
      /// <param name="recursive"><c>true</c> の場合、このディレクトリとそのサブディレクトリから空のサブディレクトリを削除します。</param>
      [SecurityCritical]
      public void DeleteEmptySubdirectories(bool recursive)
      {
         Directory.DeleteEmptySubdirectoriesCore(EntryInfo, Transaction, null, recursive, false, PathFormat.LongFullPath);
      }


      /// <summary>[AlphaFS] <see cref="DirectoryInfo"/> インスタンスから空のサブディレクトリを削除します。</summary>
      /// <param name="recursive"><c>true</c> の場合、このディレクトリとそのサブディレクトリから空のサブディレクトリを削除します。</param>
      /// <param name="ignoreReadOnly"><c>true</c> の場合、空のディレクトリの読み取り専用 <see cref="FileAttributes"/> をオーバーライドします。</param>
      [SecurityCritical]
      public void DeleteEmptySubdirectories(bool recursive, bool ignoreReadOnly)
      {
         Directory.DeleteEmptySubdirectoriesCore(EntryInfo, Transaction, null, recursive, ignoreReadOnly, PathFormat.LongFullPath);
      }
   }
}
