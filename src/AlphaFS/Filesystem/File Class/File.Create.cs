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
using System.Security.AccessControl;
using FileStream = System.IO.FileStream;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class File
   {
      #region .NET

      /// <summary>指定されたパスにファイルを作成するか、上書きします。</summary>
      /// <param name="path">作成するファイルのパスと名前。</param>
      /// <returns><paramref name="path"/>で指定されたファイルへの読み取り/書き込みアクセスを提供する<see cref="FileStream"/>。</returns>
      [SecurityCritical]
      public static FileStream Create(string path)
      {
         return CreateFileStreamCore(null, path, ExtendedFileAttributes.Normal, null, FileMode.Create, FileAccess.ReadWrite, FileShare.None, NativeMethods.DefaultFileBufferSize, PathFormat.RelativePath);
      }


      /// <summary>指定されたファイルを作成するか、上書きします。</summary>
      /// <param name="path">ファイルの名前。</param>
      /// <param name="bufferSize">ファイルの読み取りと書き込みのためにバッファリングされるバイト数。</param>
      /// <returns><paramref name="path"/>で指定されたファイルへの読み取り/書き込みアクセスを提供する、指定されたバッファサイズの<see cref="FileStream"/>。</returns>
      [SecurityCritical]
      public static FileStream Create(string path, int bufferSize)
      {
         return CreateFileStreamCore(null, path, ExtendedFileAttributes.Normal, null, FileMode.Create, FileAccess.ReadWrite, FileShare.None, bufferSize, PathFormat.RelativePath);
      }


      /// <summary>Creates or overwrites the specified file, specifying a buffer size and a <see cref="FileOptions"/> value that describes how to create or overwrite 閉じます。</summary>
      /// <param name="path">ファイルの名前。</param>
      /// <param name="bufferSize">ファイルの読み取りと書き込みのためにバッファリングされるバイト数。</param>
      /// <param name="options">ファイルの作成または上書き方法を記述する<see cref="FileOptions"/>値の1つ。</param>
      /// <returns>指定されたバッファサイズの新しいファイル。</returns>
      [SecurityCritical]
      public static FileStream Create(string path, int bufferSize, FileOptions options)
      {
         return CreateFileStreamCore(null, path, (ExtendedFileAttributes) options, null, FileMode.Create, FileAccess.ReadWrite, FileShare.None, bufferSize, PathFormat.RelativePath);
      }


      /// <summary>Creates or overwrites the specified file, specifying a buffer size and a <see cref="FileOptions"/> value that describes how to create or overwrite 閉じます。</summary>
      /// <param name="path">ファイルの名前。</param>
      /// <param name="bufferSize">ファイルの読み取りと書き込みのためにバッファリングされるバイト数。</param>
      /// <param name="options">ファイルの作成または上書き方法を記述する<see cref="FileOptions"/>値の1つ。</param>
      /// <param name="fileSecurity">ファイルのアクセス制御と監査セキュリティを決定する<see cref="FileSecurity"/>値の1つ。</param>
      /// <returns>指定されたバッファサイズ、ファイルオプション、およびファイルセキュリティの新しいファイル。</returns>
      [SecurityCritical]
      public static FileStream Create(string path, int bufferSize, FileOptions options, FileSecurity fileSecurity)
      {
         return CreateFileStreamCore(null, path, (ExtendedFileAttributes) options, fileSecurity, FileMode.Create, FileAccess.ReadWrite, FileShare.None, bufferSize, PathFormat.RelativePath);
      }

      #endregion // .NET


      /// <summary>[AlphaFS] 指定されたパスにファイルを作成するか、上書きします。</summary>
      /// <param name="path">作成するファイルのパスと名前。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      /// <returns><paramref name="path"/>で指定されたファイルへの読み取り/書き込みアクセスを提供する<see cref="FileStream"/>。</returns>
      [SecurityCritical]
      public static FileStream Create(string path, PathFormat pathFormat)
      {
         return CreateFileStreamCore(null, path, ExtendedFileAttributes.Normal, null, FileMode.Create, FileAccess.ReadWrite, FileShare.None, NativeMethods.DefaultFileBufferSize, pathFormat);
      }


      /// <summary>[AlphaFS] 指定されたファイルを作成するか、上書きします。</summary>
      /// <param name="path">ファイルの名前。</param>
      /// <param name="bufferSize">ファイルの読み取りと書き込みのためにバッファリングされるバイト数。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      /// <returns><paramref name="path"/>で指定されたファイルへの読み取り/書き込みアクセスを提供する、指定されたバッファサイズの<see cref="FileStream"/>。</returns>
      [SecurityCritical]
      public static FileStream Create(string path, int bufferSize, PathFormat pathFormat)
      {
         return CreateFileStreamCore(null, path, ExtendedFileAttributes.Normal, null, FileMode.Create, FileAccess.ReadWrite, FileShare.None, bufferSize, pathFormat);
      }


      /// <summary>[AlphaFS] 指定されたファイルを作成または上書きし、バッファサイズとファイルの作成または上書き方法を記述する<see cref="FileOptions"/>値を指定します。</summary>
      /// <param name="path">ファイルの名前。</param>
      /// <param name="bufferSize">ファイルの読み取りと書き込みのためにバッファリングされるバイト数。</param>
      /// <param name="options">ファイルの作成または上書き方法を記述する<see cref="FileOptions"/>値の1つ。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      /// <returns>指定されたバッファサイズの新しいファイル。</returns>
      [SecurityCritical]
      public static FileStream Create(string path, int bufferSize, FileOptions options, PathFormat pathFormat)
      {
         return CreateFileStreamCore(null, path, (ExtendedFileAttributes) options, null, FileMode.Create, FileAccess.ReadWrite, FileShare.None, bufferSize, pathFormat);
      }


      /// <summary>[AlphaFS] 指定されたファイルを作成または上書きし、バッファサイズとファイルの作成または上書き方法を記述する<see cref="FileOptions"/>値を指定します。</summary>
      /// <param name="path">ファイルの名前。</param>
      /// <param name="bufferSize">ファイルの読み取りと書き込みのためにバッファリングされるバイト数。</param>
      /// <param name="options">ファイルの作成または上書き方法を記述する<see cref="FileOptions"/>値の1つ。</param>
      /// <param name="fileSecurity">ファイルのアクセス制御と監査セキュリティを決定する<see cref="FileSecurity"/>値の1つ。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      /// <returns>指定されたバッファサイズ、ファイルオプション、およびファイルセキュリティの新しいファイル。</returns>
      [SecurityCritical]
      public static FileStream Create(string path, int bufferSize, FileOptions options, FileSecurity fileSecurity, PathFormat pathFormat)
      {
         return CreateFileStreamCore(null, path, (ExtendedFileAttributes) options, fileSecurity, FileMode.Create, FileAccess.ReadWrite, FileShare.None, bufferSize, pathFormat);
      }
   }
}
