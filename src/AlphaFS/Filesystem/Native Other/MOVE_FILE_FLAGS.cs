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
      public enum MOVE_FILE_FLAGS
      {
         /// <summary>MoveOptions を使用しません。ファイル名が既に存在する場合、失敗します。</summary>
         None = 0,

         /// <summary>MOVE_FILE_REPLACE_EXISTSING
         /// <para>コピー先のファイル名が既に存在する場合、関数はその内容をソースファイルの内容で置き換えます。</para>
         /// <para>lpNewFileName または lpExistingFileName がディレクトリを指定する場合、この値は使用できません。</para>
         /// <para>ソースまたはコピー先のいずれかがディレクトリを指定する場合、この値は使用できません。</para>
         /// </summary>
         MOVE_FILE_REPLACE_EXISTSING = 1,

         /// <summary>MOVE_FILE_COPY_ALLOWED
         /// <para>ファイルが別のボリュームに移動される場合、関数は CopyFile および DeleteFile 関数を使用して移動をシミュレートします。</para>
         /// <para>この値は <see cref="MOVE_FILE_FLAGS.MOVE_FILE_DELAY_UNTIL_REBOOT"/> と併用できません。</para>
         /// </summary>
         MOVE_FILE_COPY_ALLOWED = 2,

         /// <summary>MOVE_FILE_DELAY_UNTIL_REBOOT
         /// <para>
         /// オペレーティングシステムが再起動されるまで、システムはファイルを移動しません。
         /// システムは AUTOCHK が実行された直後、ページングファイルの作成前にファイルを移動します。
         /// </para>
         /// <para>
         /// これにより、このパラメーターは以前の起動からのページングファイルを削除する機能を有効にします。
         /// この値は、プロセスが管理者グループまたは LocalSystem アカウントに属するユーザーのコンテキスト内にある場合にのみ使用できます。
         /// </para>
         /// <para>この値は <see cref="MOVE_FILE_FLAGS.MOVE_FILE_COPY_ALLOWED"/> と併用できません。</para>
         /// </summary>
         MOVE_FILE_DELAY_UNTIL_REBOOT = 4,


         /// <summary>MOVE_FILE_WRITE_THROUGH
         /// <para>ファイルが実際にディスク上で移動されるまで、関数は戻りません。</para>
         /// <para>
         /// この値を設定すると、コピーと削除操作として実行される移動が、関数が戻る前にディスクにフラッシュされることが保証されます。
         /// フラッシュはコピー操作の終了時に行われます。
         /// </para>
         /// <para><see cref="MOVE_FILE_FLAGS.MOVE_FILE_DELAY_UNTIL_REBOOT"/> が設定されている場合、この値は効果がありません。</para>
         /// </summary>
         MOVE_FILE_WRITE_THROUGH = 8,


         /// <summary>MOVE_FILE_CREATE_HARDLINK
         /// <para>将来の使用のために予約されています。</para>
         /// </summary>
         MOVE_FILE_CREATE_HARDLINK = 16,


         /// <summary>MOVE_FILE_FAIL_IF_NOT_TRACKABLE
         /// <para>ソースファイルがリンクソースであるが、移動後にファイルを追跡できない場合、関数は失敗します。</para>
         /// <para>この状況は、コピー先が FAT ファイルシステムでフォーマットされたボリュームである場合に発生する可能性があります。</para>
         /// </summary>
         MOVE_FILE_FAIL_IF_NOT_TRACKABLE = 32
      }
   }
}
