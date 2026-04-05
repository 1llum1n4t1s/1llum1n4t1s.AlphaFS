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
      /// <summary>[AlphaFS] 暗号化されたファイルを復元(インポート)します。 This is one of a group of Encrypted File System (EFS) functions that is
      ///   暗号化された状態のまま、ファイルのバックアップおよび復元機能を実装することを目的としています。</summary>
      /// <remarks>
      ///   <para>
      ///     呼び出し元がファイルのキーにアクセスできない場合、呼び出し元は
      ///     <see cref="Security.Privilege.Backup"/> to restore encrypted files. See
      ///     <see cref="Security.PrivilegeEnabler"/>.
      ///   </para>
      ///   <para>
      ///     暗号化されたファイルを復元するには、次のいずれかを呼び出します。
      ///     <see cref="O:Alphaleonis.Win32.Filesystem.File.ImportEncryptedFileRaw"/> overloads and specify the file to restore
      ///     復元データのコピー先ストリームとともに指定します。
      ///   </para>
      ///   <para>
      ///     この関数は暗号化されたファイルのみの復元を目的としています。暗号化されていないファイルのバックアップについては<see cref="BackupFileStream"/>を参照してください。
      ///     backup 
      ///   </para>
      /// </remarks>
      /// <param name="inputStream">以前にバックアップしたデータを読み取るストリーム。</param>
      /// <param name="destinationFilePath">復元先のファイルのパス。</param>
      /// <seealso cref="O:Alphaleonis.Win32.Filesystem.File.ExportEncryptedFileRaw"/>
      public static void ImportEncryptedFileRaw(Stream inputStream, string destinationFilePath)
      {
         ImportExportEncryptedFileDirectoryRawCore(false, false, inputStream, destinationFilePath, false, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] 暗号化されたファイルを復元(インポート)します。 This is one of a group of Encrypted File System (EFS) functions that is
      ///   暗号化された状態のまま、ファイルのバックアップおよび復元機能を実装することを目的としています。</summary>
      /// <remarks>
      ///   <para>
      ///     呼び出し元がファイルのキーにアクセスできない場合、呼び出し元は
      ///     <see cref="Security.Privilege.Backup"/> to restore encrypted files. See
      ///     <see cref="Security.PrivilegeEnabler"/>.
      ///   </para>
      ///   <para>
      ///     暗号化されたファイルを復元するには、次のいずれかを呼び出します。
      ///     <see cref="O:Alphaleonis.Win32.Filesystem.File.ImportEncryptedFileRaw"/> overloads and specify the file to restore
      ///     復元データのコピー先ストリームとともに指定します。
      ///   </para>
      ///   <para>
      ///     この関数は暗号化されたファイルのみの復元を目的としています。暗号化されていないファイルのバックアップについては<see cref="BackupFileStream"/>を参照してください。
      ///     backup 
      ///   </para>
      /// </remarks>
      /// <param name="inputStream">以前にバックアップしたデータを読み取るストリーム。</param>
      /// <param name="destinationFilePath">復元先のファイルのパス。</param>
      /// <param name="pathFormat"><paramref name="destinationFilePath"/>パラメータのパス形式。</param>
      /// <seealso cref="O:Alphaleonis.Win32.Filesystem.File.ExportEncryptedFileRaw"/>
      public static void ImportEncryptedFileRaw(Stream inputStream, string destinationFilePath, PathFormat pathFormat)
      {
         ImportExportEncryptedFileDirectoryRawCore(false, false, inputStream, destinationFilePath, false, pathFormat);
      }


      /// <summary>[AlphaFS] 暗号化されたファイルを復元(インポート)します。 This is one of a group of Encrypted File System (EFS) functions that is
      ///   暗号化された状態のまま、ファイルのバックアップおよび復元機能を実装することを目的としています。</summary>
      /// <remarks>
      ///   <para>
      ///     呼び出し元がファイルのキーにアクセスできない場合、呼び出し元は
      ///     <see cref="Security.Privilege.Backup"/> to restore encrypted files. See
      ///     <see cref="Security.PrivilegeEnabler"/>.
      ///   </para>
      ///   <para>
      ///     暗号化されたファイルを復元するには、次のいずれかを呼び出します。
      ///     <see cref="O:Alphaleonis.Win32.Filesystem.File.ImportEncryptedFileRaw"/> overloads and specify the file to restore
      ///     復元データのコピー先ストリームとともに指定します。
      ///   </para>
      ///   <para>
      ///     この関数は暗号化されたファイルのみの復元を目的としています。暗号化されていないファイルのバックアップについては<see cref="BackupFileStream"/>を参照してください。
      ///     backup 
      ///   </para>
      /// </remarks>
      /// <param name="inputStream">以前にバックアップしたデータを読み取るストリーム。</param>
      /// <param name="destinationFilePath">復元先のファイルのパス。</param>
      /// <param name="overwriteHidden"><c>true</c>に設定した場合、インポート時に隠しファイルが上書きされます。</param>
      /// <seealso cref="O:Alphaleonis.Win32.Filesystem.File.ExportEncryptedFileRaw"/>
      public static void ImportEncryptedFileRaw(Stream inputStream, string destinationFilePath, bool overwriteHidden)
      {
         ImportExportEncryptedFileDirectoryRawCore(false, false, inputStream, destinationFilePath, overwriteHidden, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] 暗号化されたファイルを復元(インポート)します。 This is one of a group of Encrypted File System (EFS) functions that is
      ///   暗号化された状態のまま、ファイルのバックアップおよび復元機能を実装することを目的としています。</summary>
      /// <remarks>
      ///   <para>
      ///     呼び出し元がファイルのキーにアクセスできない場合、呼び出し元は
      ///     <see cref="Security.Privilege.Backup"/> to restore encrypted files. See
      ///     <see cref="Security.PrivilegeEnabler"/>.
      ///   </para>
      ///   <para>
      ///     暗号化されたファイルを復元するには、次のいずれかを呼び出します。
      ///     <see cref="O:Alphaleonis.Win32.Filesystem.File.ImportEncryptedFileRaw"/> overloads and specify the file to restore
      ///     復元データのコピー先ストリームとともに指定します。
      ///   </para>
      ///   <para>
      ///     この関数は暗号化されたファイルのみの復元を目的としています。暗号化されていないファイルのバックアップについては<see cref="BackupFileStream"/>を参照してください。
      ///     backup 
      ///   </para>
      /// </remarks>
      /// <param name="inputStream">以前にバックアップしたデータを読み取るストリーム。</param>
      /// <param name="destinationFilePath">復元先のファイルのパス。</param>
      /// <param name="overwriteHidden"><c>true</c>に設定した場合、インポート時に隠しファイルが上書きされます。</param>
      /// <param name="pathFormat"><paramref name="destinationFilePath"/>パラメータのパス形式。</param>
      /// <seealso cref="O:Alphaleonis.Win32.Filesystem.File.ExportEncryptedFileRaw"/>
      public static void ImportEncryptedFileRaw(Stream inputStream, string destinationFilePath, bool overwriteHidden, PathFormat pathFormat)
      {
         ImportExportEncryptedFileDirectoryRawCore(false, false, inputStream, destinationFilePath, overwriteHidden, pathFormat);
      }
   }
}
