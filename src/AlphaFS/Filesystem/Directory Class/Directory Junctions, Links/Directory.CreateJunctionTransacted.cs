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
   public static partial class Directory
   {
      #region Obsolete

      /// <summary>[AlphaFS] NTFSディレクトリジャンクションを作成します（CMDコマンド: "MKLINK /J" に類似）。</summary>
      /// <remarks>
      /// ディレクトリは空であり、ローカルボリュームに存在する必要があります。
      /// <para>
      ///   MSDN: ジャンクション（ソフトリンクとも呼ばれる）は、参照するストレージオブジェクトが別個のディレクトリである点でハードリンクと異なり、
      ///   ジャンクションは同じコンピュータ上の異なるローカルボリュームにあるディレクトリをリンクできます。
      ///   それ以外の点では、ジャンクションはハードリンクと同様に動作します。ジャンクションはリパースポイントを通じて実装されます。
      /// </para>
      /// </remarks>
      /// <exception cref="AlreadyExistsException"/>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="junctionPath">作成するジャンクションポイントのパス。</param>
      /// <param name="directoryPath">The path to the directory. If the directory does not exist it will be created.</param>
      [SecurityCritical]
      public static void CreateJunction(KernelTransaction transaction, string junctionPath, string directoryPath)
      {
         CreateJunctionCore(transaction, junctionPath, directoryPath, false, false, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] NTFSディレクトリジャンクションを作成します（CMDコマンド: "MKLINK /J" に類似）。</summary>
      /// <remarks>
      /// ディレクトリは空であり、ローカルボリュームに存在する必要があります。
      /// <para>
      ///   MSDN: ジャンクション（ソフトリンクとも呼ばれる）は、参照するストレージオブジェクトが別個のディレクトリである点でハードリンクと異なり、
      ///   ジャンクションは同じコンピュータ上の異なるローカルボリュームにあるディレクトリをリンクできます。
      ///   それ以外の点では、ジャンクションはハードリンクと同様に動作します。ジャンクションはリパースポイントを通じて実装されます。
      /// </para>
      /// </remarks>
      /// <exception cref="AlreadyExistsException"/>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="junctionPath">作成するジャンクションポイントのパス。</param>
      /// <param name="directoryPath">The path to the directory. If the directory does not exist it will be created.</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static void CreateJunction(KernelTransaction transaction, string junctionPath, string directoryPath, PathFormat pathFormat)
      {
         CreateJunctionCore(transaction, junctionPath, directoryPath, false, false, pathFormat);
      }


      /// <summary>[AlphaFS] NTFSディレクトリジャンクションを作成します（CMDコマンド: "MKLINK /J" に類似）。同名のジャンクションポイントの上書きが許可されます。</summary>
      /// <remarks>
      /// ディレクトリは空であり、ローカルボリュームに存在する必要があります。
      /// <para>
      ///   MSDN: ジャンクション（ソフトリンクとも呼ばれる）は、参照するストレージオブジェクトが別個のディレクトリである点でハードリンクと異なり、
      ///   ジャンクションは同じコンピュータ上の異なるローカルボリュームにあるディレクトリをリンクできます。
      ///   それ以外の点では、ジャンクションはハードリンクと同様に動作します。ジャンクションはリパースポイントを通じて実装されます。
      /// </para>
      /// </remarks>
      /// <exception cref="AlreadyExistsException"/>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="junctionPath">作成するジャンクションポイントのパス。</param>
      /// <param name="directoryPath">The path to the directory. If the directory does not exist it will be created.</param>
      /// <param name="overwrite"><c>true</c> to overwrite an existing junction point. The directory is removed and recreated.</param>
      [SecurityCritical]
      public static void CreateJunction(KernelTransaction transaction, string junctionPath, string directoryPath, bool overwrite)
      {
         CreateJunctionCore(transaction, junctionPath, directoryPath, overwrite, false, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] NTFSディレクトリジャンクションを作成します（CMDコマンド: "MKLINK /J" に類似）。同名のジャンクションポイントの上書きが許可されます。</summary>
      /// <remarks>
      /// ディレクトリは空であり、ローカルボリュームに存在する必要があります。
      /// <para>
      ///   MSDN: ジャンクション（ソフトリンクとも呼ばれる）は、参照するストレージオブジェクトが別個のディレクトリである点でハードリンクと異なり、
      ///   ジャンクションは同じコンピュータ上の異なるローカルボリュームにあるディレクトリをリンクできます。
      ///   それ以外の点では、ジャンクションはハードリンクと同様に動作します。ジャンクションはリパースポイントを通じて実装されます。
      /// </para>
      /// </remarks>
      /// <exception cref="AlreadyExistsException"/>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="junctionPath">作成するジャンクションポイントのパス。</param>
      /// <param name="directoryPath">The path to the directory. If the directory does not exist it will be created.</param>
      /// <param name="overwrite"><c>true</c> to overwrite an existing junction point. The directory is removed and recreated.</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static void CreateJunction(KernelTransaction transaction, string junctionPath, string directoryPath, bool overwrite, PathFormat pathFormat)
      {
         CreateJunctionCore(transaction, junctionPath, directoryPath, overwrite, false, pathFormat);
      }


      /// <summary>[AlphaFS] NTFSディレクトリジャンクションを作成します（CMDコマンド: "MKLINK /J" に類似）。同名のジャンクションポイントの上書きが許可されます。</summary>
      /// <remarks>
      /// ディレクトリは空であり、ローカルボリュームに存在する必要があります。
      /// <paramref name="directoryPath"/>（ターゲット）のディレクトリの日付と時刻スタンプが the directory junction.
      /// <para>
      ///   MSDN: ジャンクション（ソフトリンクとも呼ばれる）は、参照するストレージオブジェクトが別個のディレクトリである点でハードリンクと異なり、
      ///   ジャンクションは同じコンピュータ上の異なるローカルボリュームにあるディレクトリをリンクできます。
      ///   それ以外の点では、ジャンクションはハードリンクと同様に動作します。ジャンクションはリパースポイントを通じて実装されます。
      /// </para>
      /// </remarks>
      /// <exception cref="AlreadyExistsException"/>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="junctionPath">作成するジャンクションポイントのパス。</param>
      /// <param name="directoryPath">The path to the directory. If the directory does not exist it will be created.</param>
      /// <param name="overwrite"><c>true</c> to overwrite an existing junction point. The directory is removed and recreated.</param>
      /// <param name="copyTargetTimestamps"><c>true</c> to copy the target date and time stamps to the directory junction.</param>
      [SecurityCritical]
      public static void CreateJunction(KernelTransaction transaction, string junctionPath, string directoryPath, bool overwrite, bool copyTargetTimestamps)
      {
         CreateJunctionCore(transaction, junctionPath, directoryPath, overwrite, copyTargetTimestamps, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] NTFSディレクトリジャンクションを作成します（CMDコマンド: "MKLINK /J" に類似）。同名のジャンクションポイントの上書きが許可されます。</summary>
      /// <remarks>
      /// ディレクトリは空であり、ローカルボリュームに存在する必要があります。
      /// <paramref name="directoryPath"/>（ターゲット）のディレクトリの日付と時刻スタンプが the directory junction.
      /// <para>
      ///   MSDN: ジャンクション（ソフトリンクとも呼ばれる）は、参照するストレージオブジェクトが別個のディレクトリである点でハードリンクと異なり、
      ///   ジャンクションは同じコンピュータ上の異なるローカルボリュームにあるディレクトリをリンクできます。
      ///   それ以外の点では、ジャンクションはハードリンクと同様に動作します。ジャンクションはリパースポイントを通じて実装されます。
      /// </para>
      /// </remarks>
      /// <exception cref="AlreadyExistsException"/>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="junctionPath">作成するジャンクションポイントのパス。</param>
      /// <param name="directoryPath">The path to the directory. If the directory does not exist it will be created.</param>
      /// <param name="overwrite"><c>true</c> to overwrite an existing junction point. The directory is removed and recreated.</param>
      /// <param name="copyTargetTimestamps"><c>true</c> to copy the target date and time stamps to the directory junction.</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static void CreateJunction(KernelTransaction transaction, string junctionPath, string directoryPath, bool overwrite, bool copyTargetTimestamps, PathFormat pathFormat)
      {
         CreateJunctionCore(transaction, junctionPath, directoryPath, overwrite, copyTargetTimestamps, pathFormat);
      }

      #endregion // Obsolete


      /// <summary>[AlphaFS] NTFSディレクトリジャンクションを作成します（CMDコマンド: "MKLINK /J" に類似）。</summary>
      /// <remarks>
      /// ディレクトリは空であり、ローカルボリュームに存在する必要があります。
      /// <para>
      ///   MSDN: ジャンクション（ソフトリンクとも呼ばれる）は、参照するストレージオブジェクトが別個のディレクトリである点でハードリンクと異なり、
      ///   ジャンクションは同じコンピュータ上の異なるローカルボリュームにあるディレクトリをリンクできます。
      ///   それ以外の点では、ジャンクションはハードリンクと同様に動作します。ジャンクションはリパースポイントを通じて実装されます。
      /// </para>
      /// </remarks>
      /// <exception cref="AlreadyExistsException"/>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="junctionPath">作成するジャンクションポイントのパス。</param>
      /// <param name="directoryPath">The path to the directory. If the directory does not exist it will be created.</param>
      [SecurityCritical]
      public static void CreateJunctionTransacted(KernelTransaction transaction, string junctionPath, string directoryPath)
      {
         CreateJunctionCore(transaction, junctionPath, directoryPath, false, false, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] NTFSディレクトリジャンクションを作成します（CMDコマンド: "MKLINK /J" に類似）。</summary>
      /// <remarks>
      /// ディレクトリは空であり、ローカルボリュームに存在する必要があります。
      /// <para>
      ///   MSDN: ジャンクション（ソフトリンクとも呼ばれる）は、参照するストレージオブジェクトが別個のディレクトリである点でハードリンクと異なり、
      ///   ジャンクションは同じコンピュータ上の異なるローカルボリュームにあるディレクトリをリンクできます。
      ///   それ以外の点では、ジャンクションはハードリンクと同様に動作します。ジャンクションはリパースポイントを通じて実装されます。
      /// </para>
      /// </remarks>
      /// <exception cref="AlreadyExistsException"/>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="junctionPath">作成するジャンクションポイントのパス。</param>
      /// <param name="directoryPath">The path to the directory. If the directory does not exist it will be created.</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static void CreateJunctionTransacted(KernelTransaction transaction, string junctionPath, string directoryPath, PathFormat pathFormat)
      {
         CreateJunctionCore(transaction, junctionPath, directoryPath, false, false, pathFormat);
      }


      /// <summary>[AlphaFS] NTFSディレクトリジャンクションを作成します（CMDコマンド: "MKLINK /J" に類似）。同名のジャンクションポイントの上書きが許可されます。</summary>
      /// <remarks>
      /// ディレクトリは空であり、ローカルボリュームに存在する必要があります。
      /// <para>
      ///   MSDN: ジャンクション（ソフトリンクとも呼ばれる）は、参照するストレージオブジェクトが別個のディレクトリである点でハードリンクと異なり、
      ///   ジャンクションは同じコンピュータ上の異なるローカルボリュームにあるディレクトリをリンクできます。
      ///   それ以外の点では、ジャンクションはハードリンクと同様に動作します。ジャンクションはリパースポイントを通じて実装されます。
      /// </para>
      /// </remarks>
      /// <exception cref="AlreadyExistsException"/>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="junctionPath">作成するジャンクションポイントのパス。</param>
      /// <param name="directoryPath">The path to the directory. If the directory does not exist it will be created.</param>
      /// <param name="overwrite"><c>true</c> to overwrite an existing junction point. The directory is removed and recreated.</param>
      [SecurityCritical]
      public static void CreateJunctionTransacted(KernelTransaction transaction, string junctionPath, string directoryPath, bool overwrite)
      {
         CreateJunctionCore(transaction, junctionPath, directoryPath, overwrite, false, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] NTFSディレクトリジャンクションを作成します（CMDコマンド: "MKLINK /J" に類似）。同名のジャンクションポイントの上書きが許可されます。</summary>
      /// <remarks>
      /// ディレクトリは空であり、ローカルボリュームに存在する必要があります。
      /// <para>
      ///   MSDN: ジャンクション（ソフトリンクとも呼ばれる）は、参照するストレージオブジェクトが別個のディレクトリである点でハードリンクと異なり、
      ///   ジャンクションは同じコンピュータ上の異なるローカルボリュームにあるディレクトリをリンクできます。
      ///   それ以外の点では、ジャンクションはハードリンクと同様に動作します。ジャンクションはリパースポイントを通じて実装されます。
      /// </para>
      /// </remarks>
      /// <exception cref="AlreadyExistsException"/>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="junctionPath">作成するジャンクションポイントのパス。</param>
      /// <param name="directoryPath">The path to the directory. If the directory does not exist it will be created.</param>
      /// <param name="overwrite"><c>true</c> to overwrite an existing junction point. The directory is removed and recreated.</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static void CreateJunctionTransacted(KernelTransaction transaction, string junctionPath, string directoryPath, bool overwrite, PathFormat pathFormat)
      {
         CreateJunctionCore(transaction, junctionPath, directoryPath, overwrite, false, pathFormat);
      }


      /// <summary>[AlphaFS] NTFSディレクトリジャンクションを作成します（CMDコマンド: "MKLINK /J" に類似）。同名のジャンクションポイントの上書きが許可されます。</summary>
      /// <remarks>
      /// ディレクトリは空であり、ローカルボリュームに存在する必要があります。
      /// <paramref name="directoryPath"/>（ターゲット）のディレクトリの日付と時刻スタンプが the directory junction.
      /// <para>
      ///   MSDN: ジャンクション（ソフトリンクとも呼ばれる）は、参照するストレージオブジェクトが別個のディレクトリである点でハードリンクと異なり、
      ///   ジャンクションは同じコンピュータ上の異なるローカルボリュームにあるディレクトリをリンクできます。
      ///   それ以外の点では、ジャンクションはハードリンクと同様に動作します。ジャンクションはリパースポイントを通じて実装されます。
      /// </para>
      /// </remarks>
      /// <exception cref="AlreadyExistsException"/>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="junctionPath">作成するジャンクションポイントのパス。</param>
      /// <param name="directoryPath">The path to the directory. If the directory does not exist it will be created.</param>
      /// <param name="overwrite"><c>true</c> to overwrite an existing junction point. The directory is removed and recreated.</param>
      /// <param name="copyTargetTimestamps"><c>true</c> to copy the target date and time stamps to the directory junction.</param>
      [SecurityCritical]
      public static void CreateJunctionTransacted(KernelTransaction transaction, string junctionPath, string directoryPath, bool overwrite, bool copyTargetTimestamps)
      {
         CreateJunctionCore(transaction, junctionPath, directoryPath, overwrite, copyTargetTimestamps, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] NTFSディレクトリジャンクションを作成します（CMDコマンド: "MKLINK /J" に類似）。同名のジャンクションポイントの上書きが許可されます。</summary>
      /// <remarks>
      /// ディレクトリは空であり、ローカルボリュームに存在する必要があります。
      /// <paramref name="directoryPath"/>（ターゲット）のディレクトリの日付と時刻スタンプが the directory junction.
      /// <para>
      ///   MSDN: ジャンクション（ソフトリンクとも呼ばれる）は、参照するストレージオブジェクトが別個のディレクトリである点でハードリンクと異なり、
      ///   ジャンクションは同じコンピュータ上の異なるローカルボリュームにあるディレクトリをリンクできます。
      ///   それ以外の点では、ジャンクションはハードリンクと同様に動作します。ジャンクションはリパースポイントを通じて実装されます。
      /// </para>
      /// </remarks>
      /// <exception cref="AlreadyExistsException"/>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="junctionPath">作成するジャンクションポイントのパス。</param>
      /// <param name="directoryPath">The path to the directory. If the directory does not exist it will be created.</param>
      /// <param name="overwrite"><c>true</c> to overwrite an existing junction point. The directory is removed and recreated.</param>
      /// <param name="copyTargetTimestamps"><c>true</c> to copy the target date and time stamps to the directory junction.</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static void CreateJunctionTransacted(KernelTransaction transaction, string junctionPath, string directoryPath, bool overwrite, bool copyTargetTimestamps, PathFormat pathFormat)
      {
         CreateJunctionCore(transaction, junctionPath, directoryPath, overwrite, copyTargetTimestamps, pathFormat);
      }
   }
}
