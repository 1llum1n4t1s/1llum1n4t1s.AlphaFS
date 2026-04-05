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
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Security;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class File
   {
      #region Obsolete

      /// <summary>[AlphaFS] ファイルへのシンボリックリンク(CMDコマンド"MKLINK"と同様)を作成します。</summary>
      /// <remarks>このメソッドを昇格された状態で実行するには、<see cref="Security.Privilege.CreateSymbolicLink"/>を参照してください。</remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="PlatformNotSupportedException">The operating system is older than Windows Vista.</exception>
      /// <param name="symlinkFileName">作成するシンボリックリンクのターゲット名。</param>
      /// <param name="targetFileName">作成するシンボリックリンク。</param>
      /// <param name="targetType">リンクターゲット<paramref name="targetFileName"/>がファイルかディレクトリかを示します。</param>      
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "symlink")]
      [SecurityCritical]
      [Obsolete("Methods with SymbolicLinkTarget parameter are obsolete.")]
      public static void CreateSymbolicLink(string symlinkFileName, string targetFileName, SymbolicLinkTarget targetType)
      {
         CreateSymbolicLinkCore(null, symlinkFileName, targetFileName, targetType, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] ファイルへのシンボリックリンク(CMDコマンド"MKLINK"と同様)を作成します。</summary>
      /// <remarks>このメソッドを昇格された状態で実行するには、<see cref="Security.Privilege.CreateSymbolicLink"/>を参照してください。</remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="PlatformNotSupportedException">The operating system is older than Windows Vista.</exception>
      /// <param name="symlinkFileName">作成するシンボリックリンクのターゲット名。</param>
      /// <param name="targetFileName">作成するシンボリックリンク。</param>
      /// <param name="targetType">リンクターゲット<paramref name="targetFileName"/>がファイルかディレクトリかを示します。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>      
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "symlink")]
      [SecurityCritical]
      [Obsolete("Methods with SymbolicLinkTarget parameter are obsolete.")]
      public static void CreateSymbolicLink(string symlinkFileName, string targetFileName, SymbolicLinkTarget targetType, PathFormat pathFormat)
      {
         CreateSymbolicLinkCore(null, symlinkFileName, targetFileName, targetType, pathFormat);
      }


      /// <summary>[AlphaFS] トランザクション操作として、ファイルへのシンボリックリンク(CMDコマンド"MKLINK"と同様)を作成します。</summary>
      /// <remarks>このメソッドを昇格された状態で実行するには、<see cref="Security.Privilege.CreateSymbolicLink"/>を参照してください。</remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="PlatformNotSupportedException">The operating system is older than Windows Vista.</exception>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="symlinkFileName">作成するシンボリックリンクのターゲット名。</param>
      /// <param name="targetFileName">作成するシンボリックリンク。</param>
      /// <param name="targetType">リンクターゲット<paramref name="targetFileName"/>がファイルかディレクトリかを示します。</param>      
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "symlink")]
      [SecurityCritical]
      [Obsolete("Methods with SymbolicLinkTarget parameter are obsolete.")]
      public static void CreateSymbolicLinkTransacted(KernelTransaction transaction, string symlinkFileName, string targetFileName, SymbolicLinkTarget targetType)
      {
         CreateSymbolicLinkCore(transaction, symlinkFileName, targetFileName, targetType, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] トランザクション操作として、ファイルへのシンボリックリンク(CMDコマンド"MKLINK"と同様)を作成します。</summary>
      /// <remarks>このメソッドを昇格された状態で実行するには、<see cref="Security.Privilege.CreateSymbolicLink"/>を参照してください。</remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="PlatformNotSupportedException">The operating system is older than Windows Vista.</exception>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="symlinkFileName">作成するシンボリックリンクのターゲット名。</param>
      /// <param name="targetFileName">作成するシンボリックリンク。</param>
      /// <param name="targetType">リンクターゲット<paramref name="targetFileName"/>がファイルかディレクトリかを示します。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>      
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "symlink")]
      [SecurityCritical]
      [Obsolete("Methods with SymbolicLinkTarget parameter are obsolete.")]
      public static void CreateSymbolicLinkTransacted(KernelTransaction transaction, string symlinkFileName, string targetFileName, SymbolicLinkTarget targetType, PathFormat pathFormat)
      {
         CreateSymbolicLinkCore(transaction, symlinkFileName, targetFileName, targetType, pathFormat);
      }

      #endregion // Obsolete


      /// <summary>[AlphaFS] ファイルへのシンボリックリンク(CMDコマンド"MKLINK"と同様)を作成します。</summary>
      /// <para>&#160;</para>
      /// <remarks>
      /// <para>シンボリックリンクは存在しないターゲットを指すことができます。</para>
      /// <para>シンボリックリンクを作成するとき、オペレーティングシステムはターゲットが存在するかどうかをチェックしません。</para>
      /// <para>シンボリックリンクはリパースポイントです。</para>
      /// <para>特定のパスで許可されるリパースポイント(したがってシンボリックリンク)は最大31個です。</para>
      /// <para>このメソッドを昇格された状態で実行するには、<see cref="Security.Privilege.CreateSymbolicLink"/>を参照してください。</para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="PlatformNotSupportedException">The operating system is older than Windows Vista.</exception>
      /// <param name="symlinkFileName">作成するシンボリックリンクのターゲット名。</param>
      /// <param name="targetFileName">作成するシンボリックリンク。</param>
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "symlink")]
      [SecurityCritical]
      public static void CreateSymbolicLink(string symlinkFileName, string targetFileName)
      {
         CreateSymbolicLinkCore(null, symlinkFileName, targetFileName, SymbolicLinkTarget.File, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] ファイルへのシンボリックリンク(CMDコマンド"MKLINK"と同様)を作成します。</summary>
      /// <para>&#160;</para>
      /// <remarks>
      /// <para>シンボリックリンクは存在しないターゲットを指すことができます。</para>
      /// <para>シンボリックリンクを作成するとき、オペレーティングシステムはターゲットが存在するかどうかをチェックしません。</para>
      /// <para>シンボリックリンクはリパースポイントです。</para>
      /// <para>特定のパスで許可されるリパースポイント(したがってシンボリックリンク)は最大31個です。</para>
      /// <para>このメソッドを昇格された状態で実行するには、<see cref="Security.Privilege.CreateSymbolicLink"/>を参照してください。</para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="PlatformNotSupportedException">The operating system is older than Windows Vista.</exception>
      /// <param name="symlinkFileName">作成するシンボリックリンクのターゲット名。</param>
      /// <param name="targetFileName">作成するシンボリックリンク。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>      
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "symlink")]
      [SecurityCritical]
      public static void CreateSymbolicLink(string symlinkFileName, string targetFileName, PathFormat pathFormat)
      {
         CreateSymbolicLinkCore(null, symlinkFileName, targetFileName, SymbolicLinkTarget.File, pathFormat);
      }
   }
}
