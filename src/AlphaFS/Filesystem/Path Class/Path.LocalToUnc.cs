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
using System.Net.NetworkInformation;
using System.Security;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class Path
   {
      #region Obsolete

      /// <summary>[AlphaFS] ローカルパスをネットワーク共有パスに変換します。オプションで長いパス形式での返却や末尾のバックスラッシュの追加・削除が可能です。
      ///   <para>ローカルパス（例: "C:\Windows" または "C:\Windows\"）は "\\localhost\C$\Windows" として返されます。</para>
      ///   <para>論理ドライブがネットワーク共有パス（マップドドライブ）を指している場合、共有パスは末尾の <see cref="DirectorySeparator"/> 文字なしで返されます。</para>
      /// </summary>
      /// <returns>変換が成功するとUNCパスが返されます。
      ///   <para>変換に失敗した場合、<paramref name="localPath"/> が返されます。</para>
      ///   <para><paramref name="localPath"/> が空文字列または <c>null</c> の場合、<c>null</c> が返されます。</para>
      /// </returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="PathTooLongException"/>
      /// <exception cref="NetworkInformationException"/>
      /// <param name="localPath">ローカルパス。例: "C:\Windows"。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      /// <param name="fullPathOptions">完全パス取得を制御するオプション。</param>
      [Obsolete]
      [SecurityCritical]
      public static string LocalToUnc(string localPath, PathFormat pathFormat, GetFullPathOptions fullPathOptions)
      {
         return LocalToUncCore(localPath, fullPathOptions, pathFormat);
      }

      #endregion // Obsolete


      /// <summary>[AlphaFS] ローカルパスをネットワーク共有パスに変換します。
      ///   <para>ローカルパス（例: "C:\Windows" または "C:\Windows\"）は "\\localhost\C$\Windows" として返されます。</para>
      ///   <para>論理ドライブがネットワーク共有パス（マップドドライブ）を指している場合、共有パスは末尾の <see cref="DirectorySeparator"/> 文字なしで返されます。</para>
      /// </summary>
      /// <returns>変換が成功するとUNCパスが返されます。
      ///   <para>変換に失敗した場合、<paramref name="localPath"/> が返されます。</para>
      ///   <para><paramref name="localPath"/> が空文字列または <c>null</c> の場合、<c>null</c> が返されます。</para>
      /// </returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="PathTooLongException"/>
      /// <exception cref="NetworkInformationException"/>
      /// <param name="localPath">ローカルパス。例: "C:\Windows"。</param>
      [SecurityCritical]
      public static string LocalToUnc(string localPath)
      {
         return LocalToUncCore(localPath, GetFullPathOptions.None, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] ローカルパスをネットワーク共有パスに変換します。
      ///   <para>ローカルパス（例: "C:\Windows" または "C:\Windows\"）は "\\localhost\C$\Windows" として返されます。</para>
      ///   <para>論理ドライブがネットワーク共有パス（マップドドライブ）を指している場合、共有パスは末尾の <see cref="DirectorySeparator"/> 文字なしで返されます。</para>
      /// </summary>
      /// <returns>変換が成功するとUNCパスが返されます。
      ///   <para>変換に失敗した場合、<paramref name="localPath"/> が返されます。</para>
      ///   <para><paramref name="localPath"/> が空文字列または <c>null</c> の場合、<c>null</c> が返されます。</para>
      /// </returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="PathTooLongException"/>
      /// <exception cref="NetworkInformationException"/>
      /// <param name="localPath">ローカルパス。例: "C:\Windows"。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static string LocalToUnc(string localPath, PathFormat pathFormat)
      {
         return LocalToUncCore(localPath, GetFullPathOptions.None, pathFormat);
      }


      /// <summary>[AlphaFS] ローカルパスをネットワーク共有パスに変換します。オプションで長いパス形式での返却や末尾のバックスラッシュの追加・削除が可能です。
      ///   <para>ローカルパス（例: "C:\Windows" または "C:\Windows\"）は "\\localhost\C$\Windows" として返されます。</para>
      ///   <para>論理ドライブがネットワーク共有パス（マップドドライブ）を指している場合、共有パスは末尾の <see cref="DirectorySeparator"/> 文字なしで返されます。</para>
      /// </summary>
      /// <returns>変換が成功するとUNCパスが返されます。
      ///   <para>変換に失敗した場合、<paramref name="localPath"/> が返されます。</para>
      ///   <para><paramref name="localPath"/> が空文字列または <c>null</c> の場合、<c>null</c> が返されます。</para>
      /// </returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="PathTooLongException"/>
      /// <exception cref="NetworkInformationException"/>
      /// <param name="localPath">ローカルパス。例: "C:\Windows"。</param>
      /// <param name="fullPathOptions">完全パス取得を制御するオプション。</param>
      [SecurityCritical]
      public static string LocalToUnc(string localPath, GetFullPathOptions fullPathOptions)
      {
         return LocalToUncCore(localPath, fullPathOptions, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] ローカルパスをネットワーク共有パスに変換します。オプションで長いパス形式での返却や末尾のバックスラッシュの追加・削除が可能です。
      ///   <para>ローカルパス（例: "C:\Windows" または "C:\Windows\"）は "\\localhost\C$\Windows" として返されます。</para>
      ///   <para>論理ドライブがネットワーク共有パス（マップドドライブ）を指している場合、共有パスは末尾の <see cref="DirectorySeparator"/> 文字なしで返されます。</para>
      /// </summary>
      /// <returns>変換が成功するとUNCパスが返されます。
      ///   <para>変換に失敗した場合、<paramref name="localPath"/> が返されます。</para>
      ///   <para><paramref name="localPath"/> が空文字列または <c>null</c> の場合、<c>null</c> が返されます。</para>
      /// </returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="PathTooLongException"/>
      /// <exception cref="NetworkInformationException"/>
      /// <param name="localPath">ローカルパス。例: "C:\Windows"。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      /// <param name="fullPathOptions">完全パス取得を制御するオプション。</param>
      [SecurityCritical]
      public static string LocalToUnc(string localPath, GetFullPathOptions fullPathOptions, PathFormat pathFormat)
      {
         return LocalToUncCore(localPath, fullPathOptions, pathFormat);
      }
   }
}
