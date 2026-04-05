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

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class File
   {
      /// <summary>[AlphaFS] 暗号化されたファイルをバックアップ(エクスポート)します。 This is one of a group of Encrypted File System (EFS) functions that is
      /// 暗号化された状態のまま、ファイルのバックアップおよび復元機能を実装することを目的としています。</summary>
      /// <remarks>
      ///   <para>
      ///      バックアップされるファイルは復号化されません。暗号化された状態のままバックアップされます。
      ///   </para>
      ///   <para>
      ///      呼び出し元がファイルのキーにアクセスできない場合、呼び出し元は
      ///      <see cref="Security.Privilege.Backup"/> to export encrypted files. See
      ///      <see cref="Security.PrivilegeEnabler"/>.
      ///   </para>
      ///   <para>
      ///      暗号化されたファイルをバックアップするには、次のいずれか���呼び出します。
      ///      <see cref="O:Alphaleonis.Win32.Filesystem.File.ExportEncryptedFileRaw"/> overloads and specify the file to backup
      ///      バックアップデータのコピー先ストリームとともに指定します。
      ///   </para>
      ///   <para>
      ///      この関数は暗号化されたファイルのみのバックアップを目的としています。暗号化されていないファイルのバックアップについては<see cref="BackupFileStream"/>を参照してください。
      ///      
      ///   </para>
      /// </remarks>
      /// <param name="fileName">バックアップするファイルの名前。</param>
      /// <param name="outputStream">バックアップデータが書き込まれるコピー先ストリーム。</param>
      /// <seealso cref="O:Alphaleonis.Win32.Filesystem.File.ImportEncryptedFileRaw"/>      
      public static void ExportEncryptedFileRaw(string fileName, Stream outputStream)
      {
         ImportExportEncryptedFileDirectoryRawCore(true, false, outputStream, fileName, false, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] 暗号化されたファイルをバックアップ(エクスポート)します。 This is one of a group of Encrypted File System (EFS) functions that is
      ///   暗号化された状態のまま、ファイルのバックアップおよび復元機能を実装することを目的としています。</summary>
      /// <remarks>
      ///   <para>
      ///      バックアップされるファイルは復号化されません。暗号化された状態のままバックアップされます。
      ///   </para>
      ///   <para>
      ///      呼び出し元がファイルのキーにアクセスできない場合、呼び出し元は
      ///      <see cref="Security.Privilege.Backup"/> to export encrypted files. See
      ///      <see cref="Security.PrivilegeEnabler"/>.
      ///   </para>
      ///   <para>
      ///      暗号化されたファイルをバックアップするには、次のいずれか���呼び出します。
      ///      <see cref="O:Alphaleonis.Win32.Filesystem.File.ExportEncryptedFileRaw"/> overloads and specify the file to backup
      ///      バックアップデータのコピー先ストリームとともに指定します。
      ///   </para>
      ///   <para>
      ///      この関数は暗号化されたファイルのみのバックアップを目的としています。暗号化されていないファイルのバックアップについては<see cref="BackupFileStream"/>を参照してください。
      ///      
      ///   </para>
      /// </remarks>
      /// <param name="fileName">バックアップするファイルの名前。</param>
      /// <param name="outputStream">バックアップデータが書き込まれるコピー先ストリーム。</param>
      /// <param name="pathFormat">The path format of the <paramref name="fileName"/> parameter.</param>
      /// <seealso cref="O:Alphaleonis.Win32.Filesystem.File.ImportEncryptedFileRaw"/>
      public static void ExportEncryptedFileRaw(string fileName, Stream outputStream, PathFormat pathFormat)
      {
         ImportExportEncryptedFileDirectoryRawCore(true, false, outputStream, fileName, false, pathFormat);
      }
   }
}
