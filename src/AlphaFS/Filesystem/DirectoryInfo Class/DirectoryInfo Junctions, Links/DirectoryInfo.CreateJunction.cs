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
   public sealed partial class DirectoryInfo
   {
      /// <summary>[AlphaFS] <see cref="DirectoryInfo"/> インスタンスをディレクトリジャンクションインスタンスに変換します（CMD コマンド "MKLINK /J" と同様）。</summary>
      /// <remarks>
      /// <para>&#160;</para>
      /// <para>ディレクトリは空でローカルボリューム上に存在する必要があります。</para>
      /// <para></para>
      /// <para></para>
      /// <para>&#160;</para>
      /// <para>MSDN: ジャンクション（ソフトリンクとも呼ばれます）は、参照するストレージオブジェクトが別のディレクトリであるという点でハードリンクと異なり、</para>
      /// <para>同じコンピューター上の異なるローカルボリュームにあるディレクトリをリンクできます。</para>
      /// <para>それ以外の点では、ジャンクションはハードリンクと同じように動作します。ジャンクションはリパースポイントを通じて実装されます。</para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotAReparsePointException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="junctionPath">作成するジャンクションポイントのパス。</param>
      [SecurityCritical]
      public void CreateJunction(string junctionPath)
      {
         UpdateSourcePath(junctionPath, Directory.CreateJunctionCore(Transaction, junctionPath, LongFullName, false, false, PathFormat.RelativePath));

         RefreshEntryInfo();
      }


      /// <summary>[AlphaFS] <see cref="DirectoryInfo"/> インスタンスをディレクトリジャンクションインスタンスに変換します（CMD コマンド "MKLINK /J" と同様）。</summary>
      /// <remarks>
      /// <para>&#160;</para>
      /// <para>ディレクトリは空でローカルボリューム上に存在する必要があります。</para>
      /// <para></para>
      /// <para></para>
      /// <para>&#160;</para>
      /// <para>MSDN: ジャンクション（ソフトリンクとも呼ばれます）は、参照するストレージオブジェクトが別のディレクトリであるという点でハードリンクと異なり、</para>
      /// <para>同じコンピューター上の異なるローカルボリュームにあるディレクトリをリンクできます。</para>
      /// <para>それ以外の点では、ジャンクションはハードリンクと同じように動作します。ジャンクションはリパースポイントを通じて実装されます。</para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotAReparsePointException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="junctionPath">作成するジャンクションポイントのパス。</param>
      /// <param name="pathFormat">パスパラメーターの形式を示します。</param>
      [SecurityCritical]
      public void CreateJunction(string junctionPath, PathFormat pathFormat)
      {
         UpdateSourcePath(junctionPath, Directory.CreateJunctionCore(Transaction, junctionPath, LongFullName, false, false, pathFormat));

         RefreshEntryInfo();
      }


      /// <summary>[AlphaFS] <see cref="DirectoryInfo"/> インスタンスをディレクトリジャンクションインスタンスに変換します（CMD コマンド "MKLINK /J" と同様）。</summary>
      /// <remarks>
      /// <para>&#160;</para>
      /// <para>ディレクトリは空でローカルボリューム上に存在する必要があります。</para>
      /// <para></para>
      /// <para></para>
      /// <para>&#160;</para>
      /// <para>MSDN: ジャンクション（ソフトリンクとも呼ばれます）は、参照するストレージオブジェクトが別のディレクトリであるという点でハードリンクと異なり、</para>
      /// <para>同じコンピューター上の異なるローカルボリュームにあるディレクトリをリンクできます。</para>
      /// <para>それ以外の点では、ジャンクションはハードリンクと同じように動作します。ジャンクションはリパースポイントを通じて実装されます。</para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotAReparsePointException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="junctionPath">作成するジャンクションポイントのパス。</param>
      /// <param name="overwrite">既存のジャンクションポイントを上書きする場合は <c>true</c>。ディレクトリは削除されて再作成されます。</param>
      [SecurityCritical]
      public void CreateJunction(string junctionPath, bool overwrite)
      {
         UpdateSourcePath(junctionPath, Directory.CreateJunctionCore(Transaction, junctionPath, LongFullName, overwrite, false, PathFormat.RelativePath));

         RefreshEntryInfo();
      }


      /// <summary>[AlphaFS] <see cref="DirectoryInfo"/> インスタンスをディレクトリジャンクションインスタンスに変換します（CMD コマンド "MKLINK /J" と同様）。</summary>
      /// <remarks>
      /// <para>&#160;</para>
      /// <para>ディレクトリは空でローカルボリューム上に存在する必要があります。</para>
      /// <para></para>
      /// <para></para>
      /// <para>&#160;</para>
      /// <para>MSDN: ジャンクション（ソフトリンクとも呼ばれます）は、参照するストレージオブジェクトが別のディレクトリであるという点でハードリンクと異なり、</para>
      /// <para>同じコンピューター上の異なるローカルボリュームにあるディレクトリをリンクできます。</para>
      /// <para>それ以外の点では、ジャンクションはハードリンクと同じように動作します。ジャンクションはリパースポイントを通じて実装されます。</para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotAReparsePointException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="junctionPath">作成するジャンクションポイントのパス。</param>
      /// <param name="overwrite">既存のジャンクションポイントを上書きする場合は <c>true</c>。ディレクトリは削除されて再作成されます。</param>
      /// <param name="pathFormat">パスパラメーターの形式を示します。</param>
      [SecurityCritical]
      public void CreateJunction(string junctionPath, bool overwrite, PathFormat pathFormat)
      {
         UpdateSourcePath(junctionPath, Directory.CreateJunctionCore(Transaction, junctionPath, LongFullName, overwrite, false, pathFormat));

         RefreshEntryInfo();
      }


      /// <summary>[AlphaFS] <see cref="DirectoryInfo"/> インスタンスをディレクトリジャンクションインスタンスに変換します（CMD コマンド "MKLINK /J" と同様）。</summary>
      /// <remarks>
      /// <para>&#160;</para>
      /// <para>ディレクトリは空でローカルボリューム上に存在する必要があります。</para>
      /// <para></para>
      /// <para></para>
      /// <para>&#160;</para>
      /// <para>MSDN: ジャンクション（ソフトリンクとも呼ばれます）は、参照するストレージオブジェクトが別のディレクトリであるという点でハードリンクと異なり、</para>
      /// <para>同じコンピューター上の異なるローカルボリュームにあるディレクトリをリンクできます。</para>
      /// <para>それ以外の点では、ジャンクションはハードリンクと同じように動作します。ジャンクションはリパースポイントを通じて実装されます。</para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotAReparsePointException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="junctionPath">作成するジャンクションポイントのパス。</param>
      /// <param name="overwrite">既存のジャンクションポイントを上書きする場合は <c>true</c>。ディレクトリは削除されて再作成されます。</param>
      /// <param name="copyTargetTimestamps">ターゲットの日時スタンプをディレクトリジャンクションにコピーする場合は <c>true</c>。</param>
      [SecurityCritical]
      public void CreateJunction(string junctionPath, bool overwrite, bool copyTargetTimestamps)
      {
         UpdateSourcePath(junctionPath, Directory.CreateJunctionCore(Transaction, junctionPath, LongFullName, overwrite, copyTargetTimestamps, PathFormat.RelativePath));

         RefreshEntryInfo();
      }


      /// <summary>[AlphaFS] <see cref="DirectoryInfo"/> インスタンスをディレクトリジャンクションインスタンスに変換します（CMD コマンド "MKLINK /J" と同様）。</summary>
      /// <remarks>
      /// <para>&#160;</para>
      /// <para>ディレクトリは空でローカルボリューム上に存在する必要があります。</para>
      /// <para></para>
      /// <para></para>
      /// <para>&#160;</para>
      /// <para>MSDN: ジャンクション（ソフトリンクとも呼ばれます）は、参照するストレージオブジェクトが別のディレクトリであるという点でハードリンクと異なり、</para>
      /// <para>同じコンピューター上の異なるローカルボリュームにあるディレクトリをリンクできます。</para>
      /// <para>それ以外の点では、ジャンクションはハードリンクと同じように動作します。ジャンクションはリパースポイントを通じて実装されます。</para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotAReparsePointException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="junctionPath">作成するジャンクションポイントのパス。</param>
      /// <param name="overwrite">既存のジャンクションポイントを上書きする場合は <c>true</c>。ディレクトリは削除されて再作成されます。</param>
      /// <param name="copyTargetTimestamps">ターゲットの日時スタンプをディレクトリジャンクションにコピーする場合は <c>true</c>。</param>
      /// <param name="pathFormat">パスパラメーターの形式を示します。</param>
      [SecurityCritical]
      public void CreateJunction(string junctionPath, bool overwrite, bool copyTargetTimestamps, PathFormat pathFormat)
      {
         UpdateSourcePath(junctionPath, Directory.CreateJunctionCore(Transaction, junctionPath, LongFullName, overwrite, copyTargetTimestamps, pathFormat));

         RefreshEntryInfo();
      }
   }
}
