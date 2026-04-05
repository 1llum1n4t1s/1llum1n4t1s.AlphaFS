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

using System.Diagnostics.CodeAnalysis;
using System.Security;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class File
   {
      #region .NET

      /// <summary>指定されたファイルの内容を別のファイルの内容で置換し、元のファイルを削除し、置換されたファイルのバックアップを作成します。</summary>
      /// <remarks>Replaceメソッドは、指定されたファイルの内容を別のファイルの内容で置換します。また、置換されたファイルのバックアップを作成します。</remarks>
      /// <remarks>
      ///   <paramref name="sourceFileName"/>と<paramref name="destinationFileName"/>が異なるボリューム上にある場合、このメソッドは
      ///   例外が発生します。<paramref name="destinationBackupFileName"/>がソースファイルと異なるボリューム上にある場合、バックアップ
      ///   ファイルは削除されます。
      /// </remarks>
      /// <remarks>
      ///   置換されるファイルのバックアップを作成しない場合は、<paramref name="destinationBackupFileName"/>パラメータにnullを渡します。
      ///   
      /// </remarks>
      /// <param name="sourceFileName"><paramref name="destinationFileName"/>で指定されたファイルを置換するファイルの名前。</param>
      /// <param name="destinationFileName">置換されるファイルの名前。</param>
      /// <param name="destinationBackupFileName">バックアップファイルの名前。</param>      
      [SecurityCritical]
      public static void Replace(string sourceFileName, string destinationFileName, string destinationBackupFileName)
      {
         ReplaceCore(sourceFileName, destinationFileName, destinationBackupFileName, false, PathFormat.RelativePath);
      }
      

      /// <summary>指定されたファイルの内容を別のファイルの内容で置換し、元のファイルを削除し、置換されたファイルのバックアップを作成します。オプションでマージエラーを無視します。</summary>
      /// <remarks>Replaceメソッドは、指定されたファイルの内容を別のファイルの内容で置換します。また、置換されたファイルのバックアップを作成します。</remarks>
      /// <remarks>
      ///   <paramref name="sourceFileName"/>と<paramref name="destinationFileName"/>が異なるボリューム上にある場合、このメソッドは
      ///   例外が発生します。<paramref name="destinationBackupFileName"/>がソースファイルと異なるボリューム上にある場合、バックアップ
      ///   ファイルは削除されます。
      /// </remarks>
      /// <remarks>
      ///   置換されるファイルのバックアップを作成しない場合は、<paramref name="destinationBackupFileName"/>パラメータにnullを渡します。
      ///   
      /// </remarks>
      /// <param name="sourceFileName"><paramref name="destinationFileName"/>で指定されたファイルを置換するファイルの名前。</param>
      /// <param name="destinationFileName">置換されるファイルの名前。</param>
      /// <param name="destinationBackupFileName">バックアップファイルの名前。</param>
      /// <param name="ignoreMetadataErrors">
      ///   <c>true</c> to ignore merge errors (such as attributes and access control lists (ACLs)) from the replaced file to the
      ///   replacement file; otherwise, <c>false</c>.
      /// </param>      
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "dest")]
      [SecurityCritical]
      public static void Replace(string sourceFileName, string destinationFileName, string destinationBackupFileName, bool ignoreMetadataErrors)
      {
         ReplaceCore(sourceFileName, destinationFileName, destinationBackupFileName, ignoreMetadataErrors, PathFormat.RelativePath);
      }

      #endregion // .NET


      /// <summary>[AlphaFS] 指定されたファイルの内容を別のファイルの内容で置換し、元のファイルを削除し、置換されたファイルのバックアップを作成します。オプションでマージエラーを無視します。</summary>
      /// <remarks>Replaceメソッドは、指定されたファイルの内容を別のファイルの内容で置換します。また、置換されたファイルのバックアップを作成します。</remarks>
      /// <remarks>
      ///   <paramref name="sourceFileName"/>と<paramref name="destinationFileName"/>が異なるボリューム上にある場合、このメソッドは
      ///   例外が発生します。<paramref name="destinationBackupFileName"/>がソースファイルと異なるボリューム上にある場合、バックアップ
      ///   ファイルは削除されます。
      /// </remarks>
      /// <remarks>
      ///   置換されるファイルのバックアップを作成しない場合は、<paramref name="destinationBackupFileName"/>パラメータにnullを渡します。
      ///   
      /// </remarks>
      /// <param name="sourceFileName"><paramref name="destinationFileName"/>で指定されたファイルを置換するファイルの名前。</param>
      /// <param name="destinationFileName">置換されるファイルの名前。</param>
      /// <param name="destinationBackupFileName">バックアップファイルの名前。</param>
      /// <param name="ignoreMetadataErrors">
      ///   <c>true</c> to ignore merge errors (such as attributes and access control lists (ACLs)) from the replaced file to the
      ///   replacement file; otherwise, <c>false</c>.
      /// </param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>      
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "dest")]
      [SecurityCritical]
      public static void Replace(string sourceFileName, string destinationFileName, string destinationBackupFileName, bool ignoreMetadataErrors, PathFormat pathFormat)
      {
         ReplaceCore(sourceFileName, destinationFileName, destinationBackupFileName, ignoreMetadataErrors, pathFormat);
      }
   }
}
