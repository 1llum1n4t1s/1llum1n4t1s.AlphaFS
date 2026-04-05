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

namespace Alphaleonis.Win32.Filesystem
{
   internal static partial class NativeMethods
   {
      /// <summary>ファイルまたはディレクトリのコピー方法を指定するフラグ。</summary>
      internal enum COPY_FILE_FLAGS
      {
         /// <summary>対象ファイルが既に存在する場合、コピー操作は直ちに失敗します。</summary>
         COPY_FILE_FAIL_IF_EXISTS = 1,


         /// <summary>
         ///   コピーが失敗した場合に備えて、コピーの進行状況がターゲットファイルで追跡されます。失敗したコピーは、
         ///   失敗した呼び出しで使用したのと同じ既存のファイル名と新しいファイル名を指定することで、後で再開できます。
         ///   コピー操作中に新しいファイルが複数回フラッシュされる可能性があるため、コピー操作が大幅に遅くなる場合があります。
         /// </summary>
         COPY_FILE_RESTARTABLE = 2,


         /// <summary>ファイルがコピーされ、元のファイルが書き込みアクセス用に開かれます。</summary>
         COPY_FILE_OPEN_SOURCE_FOR_WRITE = 4,


         /// <summary>暗号化されたファイルのコピーは、コピー先で暗号化できない場合でも成功します。</summary>
         COPY_FILE_ALLOW_DECRYPTED_DESTINATION = 8,


         /// <summary>ソースファイルがシンボリックリンクの場合、コピー先ファイルもソースのシンボリックリンクが指しているのと同じファイルを指すシンボリックリンクになります。</summary>
         COPY_FILE_COPY_SYMLINK = 2048,


         /// <summary>コピー操作はバッファーなし I/O を使用して実行され、システム I/O キャッシュリソースをバイパスします。非常に大きなファイル転送に推奨されます。</summary>
         COPY_FILE_NO_BUFFERING = 4096
      }
   }
}
