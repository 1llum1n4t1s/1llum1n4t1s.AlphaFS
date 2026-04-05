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
using System.IO;
using System.Security;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class File
   {
      /// <summary>指定されたファイルを削除します。</summary>
      /// <remarks>削除するファイルが存在しない場合、例外はスローされません。</remarks>
      /// <param name="path">削除するファイルの名前。ワイルドカード文字はサポートされません。</param>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <exception cref="FileReadOnlyException"/>
      [SecurityCritical]
      public static void Delete(string path)
      {
         DeleteFileCore(null, path, false, 0, PathFormat.RelativePath);
      }


      /// <summary>指定されたファイルを削除します。</summary>
      /// <remarks>削除するファイルが存在しない場合、例外はスローされません。</remarks>
      /// <param name="path">削除するファイルの名前。ワイルドカード文字はサポートされません。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <exception cref="FileReadOnlyException"/>
      [SecurityCritical]
      public static void Delete(string path, PathFormat pathFormat)
      {
         DeleteFileCore(null, path, false, 0, pathFormat);
      }


      /// <summary>[AlphaFS] 指定されたファイルを削除します。</summary>
      /// <remarks>削除するファイルが存在しない場合、例外はスローされません。</remarks>
      /// <param name="path">削除するファイルの名前。ワイルドカード文字はサポートされません。</param>
      /// <param name="ignoreReadOnly"><c>true</c>の場合、ファイルの読み取り専用<see cref="FileAttributes"/>を上書きします。</param>      
      /// <exception cref="ArgumentException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <exception cref="FileReadOnlyException"/>
      [SecurityCritical]
      public static void Delete(string path, bool ignoreReadOnly)
      {
         DeleteFileCore(null, path, ignoreReadOnly, 0, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] 指定されたファイルを削除します。</summary>
      /// <remarks>削除するファイルが存在しない場合、例外はスローされません。</remarks>
      /// <param name="path">削除するファイルの名前。ワイルドカード文字はサポートされません。</param>
      /// <param name="ignoreReadOnly"><c>true</c>の場合、ファイルの読み取り専用<see cref="FileAttributes"/>を上書きします。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <exception cref="FileReadOnlyException"/>
      [SecurityCritical]
      public static void Delete(string path, bool ignoreReadOnly, PathFormat pathFormat)
      {
         DeleteFileCore(null, path, ignoreReadOnly, 0, pathFormat);
      }
   }
}
