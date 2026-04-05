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

using Alphaleonis.Win32.Network;
using System;
using System.Globalization;
using System.IO;
using System.Net.NetworkInformation;
using System.Security;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class Path
   {
      /// <summary>ローカルパスをネットワーク共有パスに変換します。オプションで長いパス形式での返却や末尾のバックスラッシュの追加・削除が可能です。
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
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      internal static string LocalToUncCore(string localPath, GetFullPathOptions fullPathOptions, PathFormat pathFormat)
      {
         if (Utils.IsNullOrWhiteSpace(localPath))
         {
            return null;
         }

         if (pathFormat == PathFormat.RelativePath)
         {
            CheckSupportedPathFormat(localPath, true, true);
         }


         var addTrailingDirectorySeparator = (fullPathOptions & GetFullPathOptions.AddTrailingDirectorySeparator) != 0;
         var removeTrailingDirectorySeparator = (fullPathOptions & GetFullPathOptions.RemoveTrailingDirectorySeparator) != 0;

         if (addTrailingDirectorySeparator && removeTrailingDirectorySeparator)
         {
            throw new ArgumentException(Resources.GetFullPathOptions_Add_And_Remove_DirectorySeparator_Invalid, "fullPathOptions");
         }


         if (!removeTrailingDirectorySeparator && !addTrailingDirectorySeparator)
         {
            // "localPath" がバックスラッシュで終わる場合、末尾にバックスラッシュを追加する。
            if (localPath.EndsWith(DirectorySeparator, StringComparison.Ordinal))
            {
               fullPathOptions &= ~GetFullPathOptions.RemoveTrailingDirectorySeparator; // 末尾バックスラッシュの削除を無効化する。
               fullPathOptions |= GetFullPathOptions.AddTrailingDirectorySeparator;     // 末尾バックスラッシュの追加を有効化する。
            }
         }


         var getAsLongPath = (fullPathOptions & GetFullPathOptions.AsLongPath) != 0;

         var returnUncPath = GetRegularPathCore(localPath, fullPathOptions | GetFullPathOptions.CheckInvalidPathChars, false);
         

         if (!IsUncPathCore(returnUncPath, true, false))
         {
            if (returnUncPath[0] == CurrentDirectoryPrefixChar || !IsPathRooted(returnUncPath, false))
            {
               returnUncPath = GetFullPathCore(null, false, returnUncPath, GetFullPathOptions.None);
            }


            var drive = GetPathRoot(returnUncPath, false);

            if (Utils.IsNullOrWhiteSpace(drive))
            {
               return returnUncPath;
            }


            var remoteInfo = Host.GetRemoteNameInfoCore(returnUncPath, true);


            // ネットワーク共有。
            if (!Utils.IsNullOrWhiteSpace(remoteInfo.lpUniversalName))
            {
               return getAsLongPath ? GetLongPathCore(remoteInfo.lpUniversalName, fullPathOptions) : GetRegularPathCore(remoteInfo.lpUniversalName, fullPathOptions, false);
            }


            // ネットワークルート。
            if (!Utils.IsNullOrWhiteSpace(remoteInfo.lpConnectionName))
            {
               return getAsLongPath ? GetLongPathCore(remoteInfo.lpConnectionName, fullPathOptions) : GetRegularPathCore(remoteInfo.lpConnectionName, fullPathOptions, false);
            }


            // Split: localDrive[0] = "C", localDrive[1] = "\Windows"
            var localDrive = returnUncPath.Split(VolumeSeparatorChar);

            // Return: "\\localhost\C$\Windows"
            returnUncPath = string.Format(CultureInfo.InvariantCulture, "{0}{1}{2}{3}{4}", Host.GetUncName(), DirectorySeparator, localDrive[0], NetworkDriveSeparator, localDrive[1]);
         }


         return getAsLongPath ? GetLongPathCore(returnUncPath, fullPathOptions) : GetRegularPathCore(returnUncPath, fullPathOptions, false);
      }
   }
}
